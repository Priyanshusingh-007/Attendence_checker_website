import os
import io
import uuid
import datetime
from functools import wraps
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, jsonify, send_file, abort
)
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import qrcode
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

# Explicitly configure 'statics' folder mapping
app = Flask(
    __name__,
    static_folder='statics',
    static_url_path='/statics',
    template_folder='templates'
)

app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'default-dev-secret-key-123')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///attendease.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

GOOGLE_CLIENT_ID = os.getenv('GOOGLE_CLIENT_ID')
GOOGLE_CLIENT_SECRET = os.getenv('GOOGLE_CLIENT_SECRET')
SCOPES = [
    'openid',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

# ==============================================================================
# DATABASE MODELS
# ==============================================================================

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    google_id = db.Column(db.String(100), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    name = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # 'teacher' or 'student'
    college = db.Column(db.String(150), nullable=True)
    enrollment_no = db.Column(db.String(64), unique=True, nullable=True)
    avatar_url = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.datetime.utcnow)

    classes_taught = db.relationship('ClassRoom', backref='teacher', lazy=True)
    enrollments = db.relationship('ClassEnrollment', backref='student', lazy=True)

class ClassRoom(db.Model):
    __tablename__ = 'class_rooms'
    id = db.Column(db.Integer, primary_key=True)
    teacher_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    subject = db.Column(db.String(120), nullable=False)
    section = db.Column(db.String(50), nullable=True, default='General')
    sheet_id = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.datetime.utcnow)

    enrollments = db.relationship('ClassEnrollment', backref='classroom', lazy=True, cascade='all, delete-orphan')
    sessions = db.relationship('AttendanceSession', backref='classroom', lazy=True, cascade='all, delete-orphan')

class ClassEnrollment(db.Model):
    __tablename__ = 'class_enrollments'
    id = db.Column(db.Integer, primary_key=True)
    class_id = db.Column(db.Integer, db.ForeignKey('class_rooms.id'), nullable=False)
    student_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    enrolled_at = db.Column(db.DateTime, default=datetime.datetime.utcnow)

class AttendanceSession(db.Model):
    __tablename__ = 'attendance_sessions'
    id = db.Column(db.Integer, primary_key=True)
    session_uuid = db.Column(db.String(36), unique=True, nullable=False, default=lambda: str(uuid.uuid4()))
    class_id = db.Column(db.Integer, db.ForeignKey('class_rooms.id'), nullable=False)
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    started_at = db.Column(db.DateTime, default=datetime.datetime.utcnow)
    ended_at = db.Column(db.DateTime, nullable=True)

# ==============================================================================
# AUTHORIZATION WRAPPERS
# ==============================================================================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please log in with Google to continue.', 'warning')
            return redirect(url_for('home'))
        return f(*args, **kwargs)
    return decorated_function

def teacher_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('user_role') != 'teacher':
            abort(403)
        return f(*args, **kwargs)
    return decorated_function

def student_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('user_role') != 'student':
            abort(403)
        return f(*args, **kwargs)
    return decorated_function

# ==============================================================================
# GOOGLE SHEETS HELPER FUNCTIONS
# ==============================================================================

def get_google_credentials():
    if 'oauth_token' not in session:
        return None
    token_data = session['oauth_token']
    return Credentials(
        token=token_data.get('access_token'),
        refresh_token=token_data.get('refresh_token'),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        scopes=SCOPES
    )

def provision_course_spreadsheet(creds, class_name, section):
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_body = {
        'properties': {'title': f"AttendEase_{class_name}_{section}_{datetime.date.today()}"},
        'sheets': [
            {'properties': {'title': 'Roster', 'gridProperties': {'rowCount': 250, 'columnCount': 5}}},
            {'properties': {'title': 'Attendance_Log', 'gridProperties': {'rowCount': 2500, 'columnCount': 8}}}
        ]
    }
    response = service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
    sheet_id = response.get('spreadsheetId')

    header_data = [
        {
            'range': 'Roster!A1:D1',
            'values': [['Enrollment No', 'Name', 'Email', 'Registered At']]
        },
        {
            'range': 'Attendance_Log!A1:G1',
            'values': [['Session UUID', 'Enrollment No', 'Name', 'Email', 'Date', 'Time', 'Status']]
        }
    ]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={'valueInputOption': 'USER_ENTERED', 'data': header_data}
    ).execute()
    return sheet_id

def append_sheet_roster(creds, sheet_id, student_list):
    service = build('sheets', 'v4', credentials=creds)
    values = []
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for s in student_list:
        values.append([s['enrollment_no'], s['name'], s.get('email', ''), now_str])

    service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range='Roster!A2:D',
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': values}
    ).execute()

def append_attendance_record(creds, sheet_id, record):
    service = build('sheets', 'v4', credentials=creds)
    row = [[
        record['session_uuid'],
        record['enrollment_no'],
        record['name'],
        record['email'],
        record['date'],
        record['time'],
        record['status']
    ]]
    service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range='Attendance_Log!A2:G',
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body={'values': row}
    ).execute()

def get_session_attendance_records(creds, sheet_id, session_uuid):
    service = build('sheets', 'v4', credentials=creds)
    res = service.spreadsheets().values().get(spreadsheetId=sheet_id, range='Attendance_Log!A2:G').execute()
    rows = res.get('values', [])
    records = []
    for r in rows:
        if len(r) >= 7 and r[0] == session_uuid:
            records.append({
                'session_uuid': r[0],
                'enrollment_no': r[1],
                'name': r[2],
                'email': r[3],
                'date': r[4],
                'time': r[5],
                'status': r[6]
            })
    return records

def get_full_attendance_dataset(creds, sheet_id):
    service = build('sheets', 'v4', credentials=creds)
    res = service.spreadsheets().values().get(spreadsheetId=sheet_id, range='Attendance_Log!A2:G').execute()
    return res.get('values', [])

# ==============================================================================
# AUTHENTICATION ROUTES
# ==============================================================================

@app.route('/')
def home():
    if 'user_id' in session:
        if session.get('user_role') == 'teacher':
            return redirect(url_for('teacher_dashboard'))
        return redirect(url_for('student_dashboard'))
    return render_template('front_page.html')

@app.route('/auth/google/<role>')
def google_auth_initiate(role):
    if role not in ['teacher', 'student']:
        flash('Invalid role specified.', 'danger')
        return redirect(url_for('home'))

    session['pending_role'] = role
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": GOOGLE_CLIENT_ID,
                "client_secret": GOOGLE_CLIENT_SECRET,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token"
            }
        },
        scopes=SCOPES,
        redirect_uri=url_for('google_auth_callback', _external=True)
    )
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='select_account'
    )
    session['oauth_state'] = state
    return redirect(authorization_url)

@app.route('/auth/google/callback')
def google_auth_callback():
    state = session.get('oauth_state')
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": GOOGLE_CLIENT_ID,
                "client_secret": GOOGLE_CLIENT_SECRET,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token"
            }
        },
        scopes=SCOPES,
        state=state,
        redirect_uri=url_for('google_auth_callback', _external=True)
    )

    try:
        flow.fetch_token(authorization_response=request.url)
    except Exception as e:
        flash(f'Google OAuth Authentication Failed: {str(e)}', 'danger')
        return redirect(url_for('home'))

    creds = flow.credentials
    session['oauth_token'] = {
        'access_token': creds.token,
        'refresh_token': creds.refresh_token,
        'scopes': creds.scopes
    }

    user_info_service = build('oauth2', 'v2', credentials=creds)
    user_info = user_info_service.userinfo().get().execute()

    google_id = user_info.get('id')
    email = user_info.get('email').lower()
    name = user_info.get('name')
    avatar_url = user_info.get('picture')
    role = session.pop('pending_role', 'student')

    user = User.query.filter_by(google_id=google_id).first()
    if not user:
        user = User.query.filter_by(email=email).first()
        if not user:
            user = User(
                google_id=google_id,
                email=email,
                name=name,
                role=role,
                avatar_url=avatar_url
            )
            db.session.add(user)
            db.session.commit()

    session['user_id'] = user.id
    session['user_role'] = user.role
    session['user_name'] = user.name
    session['user_email'] = user.email
    session['user_avatar'] = user.avatar_url

    # Check for required profile data
    if user.role == 'teacher' and not user.college:
        return redirect(url_for('complete_profile'))
    if user.role == 'student' and (not user.college or not user.enrollment_no):
        return redirect(url_for('complete_profile'))

    redirect_target = session.pop('redirect_after_login', None)
    if redirect_target:
        return redirect(redirect_target)

    return redirect(url_for('teacher_dashboard' if user.role == 'teacher' else 'student_dashboard'))

@app.route('/complete-profile', methods=['GET', 'POST'])
@login_required
def complete_profile():
    user = User.query.get_or_404(session['user_id'])
    if request.method == 'POST':
        college = request.form.get('college', '').strip()
        if not college:
            flash('College name is required.', 'warning')
            return render_template('profile.html', user=user)

        user.college = college
        if user.role == 'student':
            enrollment_no = request.form.get('enrollment_no', '').strip().upper()
            if not enrollment_no:
                flash('Enrollment number is required.', 'warning')
                return render_template('profile.html', user=user)
            
            existing = User.query.filter(User.enrollment_no == enrollment_no, User.id != user.id).first()
            if existing:
                flash('This Enrollment Number is already registered.', 'danger')
                return render_template('profile.html', user=user)
            user.enrollment_no = enrollment_no

        db.session.commit()
        flash('Profile configuration saved.', 'success')
        return redirect(url_for('teacher_dashboard' if user.role == 'teacher' else 'student_dashboard'))

    return render_template('profile.html', user=user)

@app.route('/logout')
def logout():
    session.clear()
    flash('Logged out successfully.', 'info')
    return redirect(url_for('home'))

# ==============================================================================
# TEACHER DASHBOARD & OPERATIONS
# ==============================================================================

@app.route('/teacher/dashboard')
@teacher_required
def teacher_dashboard():
    teacher_id = session['user_id']
    classes = ClassRoom.query.filter_by(teacher_id=teacher_id).all()

    total_classes = len(classes)
    total_students = db.session.query(ClassEnrollment).join(ClassRoom).filter(ClassRoom.teacher_id == teacher_id).count()

    active_sessions_count = db.session.query(AttendanceSession).join(ClassRoom).filter(
        ClassRoom.teacher_id == teacher_id,
        AttendanceSession.is_active == True
    ).count()

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_present = 0

    creds = get_google_credentials()
    for c in classes:
        active_sess = AttendanceSession.query.filter_by(class_id=c.id, is_active=True).first()
        c.is_active = bool(active_sess)
        c.active_session_uuid = active_sess.session_uuid if active_sess else None
        c.student_count = len(c.enrollments)

        if creds and c.sheet_id:
            try:
                sheet_rows = get_full_attendance_dataset(creds, c.sheet_id)
                today_records = [r for r in sheet_rows if len(r) >= 7 and r[4] == today_str and r[6] == 'Present']
                today_present += len(today_records)
            except Exception:
                pass

    today_rate = round((today_present / total_students) * 100, 1) if total_students > 0 and today_present > 0 else 0.0

    return render_template(
        'teacher_layout.html',
        classes=classes,
        total_classes=total_classes,
        active_classes=active_sessions_count,
        total_students=total_students,
        today_rate=today_rate
    )

@app.route('/teacher/class/create', methods=['POST'])
@teacher_required
def create_class():
    name = request.form.get('name', '').strip()
    subject = request.form.get('subject', '').strip()
    section = request.form.get('section', '').strip() or 'A'
    excel_file = request.files.get('roster_file')

    if not name or not subject or not excel_file:
        flash('Class Name, Subject, and an Excel Roster file are required.', 'danger')
        return redirect(url_for('teacher_dashboard'))

    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        df.columns = [str(c).strip().title() for c in df.columns]

        req_cols = {'Enrollment No', 'Name'}
        if not req_cols.issubset(set(df.columns)):
            flash('Uploaded file is missing required columns: "Enrollment No" and "Name".', 'danger')
            return redirect(url_for('teacher_dashboard'))

        df.dropna(subset=['Enrollment No', 'Name'], inplace=True)
        df['Enrollment No'] = df['Enrollment No'].astype(str).str.strip().str.upper()
        df['Name'] = df['Name'].astype(str).str.strip()

        if df.empty:
            flash('The Excel roster contains no valid data rows.', 'warning')
            return redirect(url_for('teacher_dashboard'))

        if df['Enrollment No'].duplicated().any():
            flash('Duplicate Enrollment Numbers found inside the spreadsheet.', 'danger')
            return redirect(url_for('teacher_dashboard'))

    except Exception as e:
        flash(f'Failed to process Excel spreadsheet: {str(e)}', 'danger')
        return redirect(url_for('teacher_dashboard'))

    creds = get_google_credentials()
    if not creds:
        flash('Google authorization expired. Please log in again.', 'warning')
        return redirect(url_for('logout'))

    try:
        sheet_id = provision_course_spreadsheet(creds, name, section)
    except HttpError as err:
        flash(f'Google Sheets Provisioning Error: {err.reason}', 'danger')
        return redirect(url_for('teacher_dashboard'))

    new_class = ClassRoom(
        teacher_id=session['user_id'],
        name=name,
        subject=subject,
        section=section,
        sheet_id=sheet_id
    )
    db.session.add(new_class)
    db.session.commit()

    student_records = []
    for _, row in df.iterrows():
        enr = row['Enrollment No']
        stu_name = row['Name']
        email_val = row.get('Email', '')
        email_str = str(email_val).strip().lower() if pd.notna(email_val) else f"{enr.lower()}@college.edu"

        student = User.query.filter_by(enrollment_no=enr).first()
        if not student:
            student = User(
                google_id=f"pending_{enr}",
                email=email_str,
                name=stu_name,
                role='student',
                enrollment_no=enr,
                college=session.get('user_college', 'Institutional Campus')
            )
            db.session.add(student)
            db.session.flush()

        if not ClassEnrollment.query.filter_by(class_id=new_class.id, student_id=student.id).first():
            enrollment = ClassEnrollment(class_id=new_class.id, student_id=student.id)
            db.session.add(enrollment)

        student_records.append({'enrollment_no': enr, 'name': stu_name, 'email': email_str})

    db.session.commit()

    try:
        append_sheet_roster(creds, sheet_id, student_records)
    except Exception as e:
        flash(f'Class created, but initial roster write to Google Sheet failed: {str(e)}', 'warning')

    flash(f'Class "{name} ({section})" successfully created with {len(student_records)} students.', 'success')
    return redirect(url_for('view_class_page', class_id=new_class.id))

@app.route('/teacher/class/<int:class_id>')
@teacher_required
def view_class_page(class_id):
    classroom = ClassRoom.query.get_or_404(class_id)
    if classroom.teacher_id != session['user_id']:
        abort(403)

    active_session = AttendanceSession.query.filter_by(class_id=classroom.id, is_active=True).first()
    enrollments = ClassEnrollment.query.filter_by(class_id=classroom.id).all()
    students = [e.student for e in enrollments]

    creds = get_google_credentials()
    sheet_records = []
    total_sessions_held = AttendanceSession.query.filter_by(class_id=classroom.id).count()

    if creds and classroom.sheet_id:
        try:
            sheet_records = get_full_attendance_dataset(creds, classroom.sheet_id)
        except Exception:
            sheet_records = []

    student_stats = []
    for s in students:
        presents = len([r for r in sheet_records if len(r) >= 7 and r[1] == s.enrollment_no and r[6] == 'Present'])
        pct = round((presents / total_sessions_held) * 100, 1) if total_sessions_held > 0 else 0.0
        student_stats.append({
            'enrollment_no': s.enrollment_no,
            'name': s.name,
            'email': s.email,
            'presents': presents,
            'total_sessions': total_sessions_held,
            'percentage': pct
        })

    return render_template(
        'class_page.html',
        classroom=classroom,
        active_session=active_session,
        students=student_stats,
        total_students=len(students)
    )

@app.route('/teacher/class/<int:class_id>/start-session', methods=['POST'])
@teacher_required
def start_attendance_session(class_id):
    classroom = ClassRoom.query.get_or_404(class_id)
    if classroom.teacher_id != session['user_id']:
        abort(403)

    AttendanceSession.query.filter_by(class_id=classroom.id, is_active=True).update({
        'is_active': False,
        'ended_at': datetime.datetime.utcnow()
    })
    
    new_session = AttendanceSession(class_id=classroom.id, is_active=True)
    db.session.add(new_session)
    db.session.commit()

    flash('Attendance session started. The QR code is active.', 'success')
    return redirect(url_for('view_class_page', class_id=classroom.id))

@app.route('/teacher/class/<int:class_id>/end-session', methods=['POST'])
@teacher_required
def end_attendance_session(class_id):
    classroom = ClassRoom.query.get_or_404(class_id)
    if classroom.teacher_id != session['user_id']:
        abort(403)

    active_session = AttendanceSession.query.filter_by(class_id=classroom.id, is_active=True).first()
    if active_session:
        active_session.is_active = False
        active_session.ended_at = datetime.datetime.utcnow()
        db.session.commit()
        flash('Attendance session ended. The QR code has been invalidated.', 'info')

    return redirect(url_for('view_class_page', class_id=classroom.id))

@app.route('/teacher/class/<int:class_id>/session-status')
@teacher_required
def session_status(class_id):
    classroom = ClassRoom.query.get_or_404(class_id)
    if classroom.teacher_id != session['user_id']:
        return jsonify({'error': 'Unauthorized'}), 403

    active_session = AttendanceSession.query.filter_by(class_id=classroom.id, is_active=True).first()
    if not active_session:
        return jsonify({'is_active': False, 'present_count': 0, 'records': []})

    creds = get_google_credentials()
    records = get_session_attendance_records(creds, classroom.sheet_id, active_session.session_uuid) if creds else []

    return jsonify({
        'is_active': True,
        'session_uuid': active_session.session_uuid,
        'present_count': len(records),
        'records': records
    })

# ==============================================================================
# STUDENT DASHBOARD & SELF-ATTENDANCE
# ==============================================================================

@app.route('/student/dashboard')
@student_required
def student_dashboard():
    student_id = session['user_id']
    student = User.query.get_or_404(student_id)
    enrollments = ClassEnrollment.query.filter_by(student_id=student_id).all()

    classes_data = []
    total_attended_overall = 0
    total_sessions_overall = 0

    creds = get_google_credentials()

    for e in enrollments:
        c = e.classroom
        sessions_count = AttendanceSession.query.filter_by(class_id=c.id).count()
        presents = 0

        if creds and c.sheet_id:
            try:
                sheet_rows = get_full_attendance_dataset(creds, c.sheet_id)
                student_records = [
                    r for r in sheet_rows 
                    if len(r) >= 7 and (r[1] == student.enrollment_no or r[3] == student.email) and r[6] == 'Present'
                ]
                presents = len(student_records)
            except Exception:
                presents = 0

        pct = round((presents / sessions_count) * 100, 1) if sessions_count > 0 else 0.0
        total_attended_overall += presents
        total_sessions_overall += sessions_count

        classes_data.append({
            'class_id': c.id,
            'name': c.name,
            'subject': c.subject,
            'section': c.section,
            'teacher_name': c.teacher.name,
            'presents': presents,
            'total_sessions': sessions_count,
            'percentage': pct
        })

    overall_pct = round((total_attended_overall / total_sessions_overall) * 100, 1) if total_sessions_overall > 0 else 0.0

    return render_template(
        'student_layout.html',
        student=student,
        classes=classes_data,
        total_classes=len(enrollments),
        total_attended=total_attended_overall,
        overall_percentage=overall_pct
    )

@app.route('/student/class/<int:class_id>/attendance')
@student_required
def student_class_attendance_view(class_id):
    student_id = session['user_id']
    student = User.query.get_or_404(student_id)
    
    membership = ClassEnrollment.query.filter_by(class_id=class_id, student_id=student_id).first()
    if not membership:
        abort(403)

    classroom = membership.classroom
    creds = get_google_credentials()
    
    sheet_rows = get_full_attendance_dataset(creds, classroom.sheet_id) if creds else []
    my_records = []
    for r in sheet_rows:
        if len(r) >= 7 and (r[1] == student.enrollment_no or r[3] == student.email):
            my_records.append({
                'date': r[4],
                'time': r[5],
                'status': r[6]
            })

    total_sessions = AttendanceSession.query.filter_by(class_id=classroom.id).count()
    presents = len([rec for rec in my_records if rec['status'] == 'Present'])
    pct = round((presents / total_sessions) * 100, 1) if total_sessions > 0 else 0.0

    return render_template(
        'student_attendance_view.html',
        classroom=classroom,
        student=student,
        records=my_records,
        presents=presents,
        total_sessions=total_sessions,
        percentage=pct
    )

# ==============================================================================
# QR INTAKE & ATTENDANCE RECORDING
# ==============================================================================

@app.route('/join/<session_uuid>')
def join_attendance_session(session_uuid):
    if 'user_id' not in session:
        session['redirect_after_login'] = url_for('join_attendance_session', session_uuid=session_uuid)
        flash('Please authenticate with Google to mark attendance.', 'info')
        return redirect(url_for('google_auth_initiate', role='student'))

    if session.get('user_role') != 'student':
        return render_template('join_class.html', success=False, message="Only students can check in to sessions.")

    student = User.query.get(session['user_id'])
    attendance_session = AttendanceSession.query.filter_by(session_uuid=session_uuid).first()

    if not attendance_session or not attendance_session.is_active:
        return render_template('join_class.html', success=False, message="This attendance session has ended or is invalid.")

    classroom = attendance_session.classroom

    membership = ClassEnrollment.query.filter_by(class_id=classroom.id, student_id=student.id).first()
    if not membership:
        membership = ClassEnrollment.query.join(User).filter(
            ClassEnrollment.class_id == classroom.id,
            User.enrollment_no == student.enrollment_no
        ).first()
        if not membership:
            return render_template('join_class.html', success=False, message=f"You are not enrolled in {classroom.name}.")

    creds = get_google_credentials()
    now = datetime.datetime.now()
    now_date = now.strftime("%Y-%m-%d")
    now_time = now.strftime("%I:%M %p")

    try:
        existing_records = get_session_attendance_records(creds, classroom.sheet_id, session_uuid)
        if any(r['enrollment_no'] == student.enrollment_no for r in existing_records):
            return render_template(
                'join_class.html',
                success=False,
                already_marked=True,
                classroom=classroom,
                message="You have already been marked present for this session."
            )

        record_payload = {
            'session_uuid': session_uuid,
            'enrollment_no': student.enrollment_no or 'N/A',
            'name': student.name,
            'email': student.email,
            'date': now_date,
            'time': now_time,
            'status': 'Present'
        }
        append_attendance_record(creds, classroom.sheet_id, record_payload)

    except Exception as e:
        return render_template('join_class.html', success=False, message=f"Google Sheet write failed: {str(e)}")

    return render_template(
        'join_class.html',
        success=True,
        classroom=classroom,
        time=now_time,
        student=student,
        message="Attendance Marked Successfully"
    )

@app.route('/qr/<session_uuid>')
@login_required
def generate_qr(session_uuid):
    target_url = url_for('join_attendance_session', session_uuid=session_uuid, _external=True)
    qr = qrcode.QRCode(version=1, box_size=8, border=2)
    qr.add_data(target_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    img_io = io.BytesIO()
    img.save(img_io, 'PNG')
    img_io.seek(0)
    return send_file(img_io, mimetype='image/png')

# ==============================================================================
# APPLICATION BOOTSTRAP
# ==============================================================================

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5000)