from flask import Flask, render_template_string, request, redirect, url_for, session, flash, make_response
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, date
import uuid
import os
import pandas as pd
from io import BytesIO

app = Flask(__name__)
app.secret_key = "b926a2f9c7ce4f7b33c997dc06bf2e738abb575d0c6faf109b90bb8d5a11c3e6"
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///crm.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# -------------------------------
# MODELS
# -------------------------------
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(128), nullable=False)
    role = db.Column(db.String(20), nullable=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        if not self.password_hash:
            return False
        return check_password_hash(self.password_hash, password)

class Lead(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lead_id = db.Column(db.String(50), unique=True, nullable=False)
    architect_name = db.Column(db.String(100))
    firm_name = db.Column(db.String(100))
    grade = db.Column(db.String(10))
    client_type = db.Column(db.String(10))
    bd_name = db.Column(db.String(50))
    meeting_date = db.Column(db.Date)
    meeting_time = db.Column(db.Time)
    remark = db.Column(db.String(200))
    assigned_to = db.Column(db.String(50))
    reschedule_date = db.Column(db.Date)
    reschedule_time = db.Column(db.Time)
    reschedule_remark = db.Column(db.String(200))
    not_interested = db.Column(db.Boolean, default=False)
    require_letter = db.Column(db.Boolean, default=False)
    email_catalogue = db.Column(db.Boolean, default=False)
    quotation_sent = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

# -------------------------------
# HELPER FUNCTIONS
# -------------------------------
def login_required(f):
    """Decorator to require login for protected routes"""
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

def admin_required(f):
    """Decorator to require admin role"""
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            flash("Access denied. Admin privileges required.")
            return redirect(url_for("dashboard"))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# -------------------------------
# SAFE DB INIT & DEFAULT ADMIN
# -------------------------------
def init_db():
    with app.app_context():
        db.create_all()
        
        # Check if admin user exists and fix password if needed
        admin = User.query.filter_by(username="admin").first()
        if not admin:
            # Create new admin user
            admin = User(username="admin", role="admin")
            admin.set_password("admin")
            db.session.add(admin)
            db.session.commit()
            print("Created new admin user")
        elif not admin.password_hash:
            # Fix existing admin user with null password
            admin.set_password("admin")
            db.session.commit()
            print("Fixed admin user password")
        else:
            print("Admin user already exists with valid password")

init_db()

# -------------------------------
# BASE TEMPLATE FUNCTION
# -------------------------------
def render_page(content, title="CRM"):
    """Helper function to render pages with base template"""
    username = session.get('username', 'User')
    role = session.get('role', 'Unknown')
    
    base_template = f"""
<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: Arial, sans-serif; margin:0; padding:0; background-color: #f4f4f4; }}
        .sidebar {{ width:220px; background:#2c3e50; height:100vh; position:fixed; color:white; overflow-y: auto; }}
        .sidebar h2 {{ text-align:center; padding: 20px 10px; margin: 0; background: #34495e; }}
        .sidebar a {{ 
            color:white; display:block; padding:12px 15px; text-decoration:none; 
            border-bottom: 1px solid #34495e; transition: background 0.3s;
        }}
        .sidebar a:hover {{ background:#34495e; }}
        .content {{ margin-left:230px; padding:20px; min-height: 100vh; }}
        .card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        .form-group {{ margin-bottom:15px; }}
        label {{ display:block; margin-bottom:5px; font-weight: bold; color: #2c3e50; }}
        input, select, textarea {{ 
            width:100%; padding:10px; border: 1px solid #ddd; border-radius: 4px; 
            font-size: 14px; transition: border-color 0.3s;
        }}
        input:focus, select:focus, textarea:focus {{ 
            outline: none; border-color: #3498db; box-shadow: 0 0 5px rgba(52, 152, 219, 0.3);
        }}
        button {{ 
            background:#3498db; color:white; padding:10px 20px; border:none; border-radius:4px; 
            cursor:pointer; font-size: 14px; transition: background 0.3s;
        }}
        button:hover {{ background:#2980b9; }}
        .flash {{ 
            padding: 10px; margin: 10px 0; border-radius: 4px; 
            background: #e74c3c; color: white; 
        }}
        .flash.success {{ background: #27ae60; }}
        table {{ 
            width: 100%; border-collapse: collapse; background: white; 
            border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background: #34495e; color: white; font-weight: bold; }}
        tr:hover {{ background: #f8f9fa; }}
        .user-info {{ 
            position: absolute; top: 10px; right: 20px; color: #7f8c8d; 
            font-size: 14px; 
        }}
        @media (max-width: 768px) {{
            .sidebar {{ width: 100%; height: auto; position: relative; }}
            .content {{ margin-left: 0; }}
        }}
    </style>
</head>
<body>
<div class="sidebar">
    <h2>CRM System</h2>
    <a href="{url_for('dashboard')}">📊 Dashboard</a>
    <a href="{url_for('new_lead')}">➕ New Lead</a>
    <a href="{url_for('assign_lead')}">📋 Assign Lead</a>
    <a href="{url_for('reschedule_meeting')}">📅 Reschedule Meeting</a>
    <a href="{url_for('meeting_stats')}">📈 Meeting Stats</a>
    <a href="{url_for('export_data')}">📤 Export Data</a>
    <a href="{url_for('manage_users')}">👥 Manage Users</a>
    <a href="{url_for('logout')}">🚪 Logout</a>
</div>
<div class="content">
    <div class="user-info">
        Welcome, {username} ({role})
    </div>
    {{{{ flash_messages }}}}
    {content}
</div>
</body>
</html>
"""
    
    # Handle flash messages
    flash_html = ""
    flashed_messages = session.get('_flashes', [])
    if flashed_messages:
        for category, msg in flashed_messages:
            flash_class = 'success' if category == 'success' else ''
            flash_html += f'<div class="flash {flash_class}">{msg}</div>'
        session.pop('_flashes', None)
    
    return base_template.replace('{{ flash_messages }}', flash_html)

# -------------------------------
# LOGIN / LOGOUT
# -------------------------------
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        
        if not username or not password:
            flash("Please enter both username and password")
        else:
            user = User.query.filter_by(username=username).first()
            if user:
                # Check if password hash exists
                if not user.password_hash:
                    flash("User account needs to be reset. Contact administrator.")
                elif user.check_password(password):
                    session["user_id"] = user.id
                    session["username"] = user.username
                    session["role"] = user.role
                    flash("Login successful!", "success")
                    return redirect(url_for("dashboard"))
                else:
                    flash("Invalid username or password")
            else:
                flash("Invalid username or password")
    
    content = """
    <div class="card" style="max-width: 400px; margin: 100px auto;">
        <h2 style="text-align: center; color: #2c3e50; margin-bottom: 30px;">Login to CRM</h2>
        <form method="post">
            <div class="form-group">
                <label for="username">Username:</label>
                <input type="text" id="username" name="username" required>
            </div>
            <div class="form-group">
                <label for="password">Password:</label>
                <input type="password" id="password" name="password" required>
            </div>
            <button type="submit" style="width: 100%;">Login</button>
        </form>
        <div style="margin-top: 20px; text-align: center; color: #7f8c8d; font-size: 12px;">
            Default: admin / admin
        </div>
    </div>
    """
    return render_page(content, "Login - CRM")

@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out successfully!", "success")
    return redirect(url_for("login"))

# -------------------------------
# DASHBOARD
# -------------------------------
@app.route("/dashboard")
@login_required
def dashboard():
    leads = Lead.query.order_by(Lead.updated_at.desc()).limit(50).all()
    total_leads = Lead.query.count()
    assigned_leads = Lead.query.filter(Lead.assigned_to.isnot(None)).count()
    pending_leads = Lead.query.filter_by(assigned_to=None).count()
    
    content = f"""
    <div class="card">
        <h2>Dashboard Overview</h2>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 20px 0;">
            <div style="background: #3498db; color: white; padding: 20px; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0;">{total_leads}</h3>
                <p style="margin: 5px 0 0 0;">Total Leads</p>
            </div>
            <div style="background: #27ae60; color: white; padding: 20px; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0;">{assigned_leads}</h3>
                <p style="margin: 5px 0 0 0;">Assigned Leads</p>
            </div>
            <div style="background: #e74c3c; color: white; padding: 20px; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0;">{pending_leads}</h3>
                <p style="margin: 5px 0 0 0;">Pending Assignment</p>
            </div>
        </div>
    </div>
    
    <div class="card">
        <h3>Recent Leads</h3>
        <div style="overflow-x: auto;">
            <table>
                <thead>
                    <tr>
                        <th>Lead ID</th>
                        <th>Architect</th>
                        <th>Firm</th>
                        <th>Grade</th>
                        <th>BD Name</th>
                        <th>Meeting Date</th>
                        <th>Assigned To</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for lead in leads:
        status = "Assigned" if lead.assigned_to else "Pending"
        meeting_date_str = lead.meeting_date.strftime("%Y-%m-%d") if lead.meeting_date else "N/A"
        content += f"""
                    <tr>
                        <td>{lead.lead_id}</td>
                        <td>{lead.architect_name or 'N/A'}</td>
                        <td>{lead.firm_name or 'N/A'}</td>
                        <td>{lead.grade or 'N/A'}</td>
                        <td>{lead.bd_name or 'N/A'}</td>
                        <td>{meeting_date_str}</td>
                        <td>{lead.assigned_to or 'Not Assigned'}</td>
                        <td>{status}</td>
                    </tr>
        """
    
    content += """
                </tbody>
            </table>
        </div>
    </div>
    """
    
    return render_page(content, "Dashboard - CRM")

# -------------------------------
# NEW LEAD
# -------------------------------
@app.route("/new_lead", methods=["GET", "POST"])
@login_required
def new_lead():
    if request.method == "POST":
        try:
            # Parse dates properly
            meeting_date = request.form.get("meeting_date")
            meeting_time = request.form.get("meeting_time")
            
            meeting_date_obj = datetime.strptime(meeting_date, "%Y-%m-%d").date() if meeting_date else None
            meeting_time_obj = datetime.strptime(meeting_time, "%H:%M").time() if meeting_time else None
            
            lead = Lead(
                lead_id=str(uuid.uuid4())[:8].upper(),
                architect_name=request.form.get("architect_name", "").strip(),
                firm_name=request.form.get("firm_name", "").strip(),
                grade=request.form.get("grade"),
                client_type=request.form.get("client_type"),
                bd_name=request.form.get("bd_name", "").strip(),
                meeting_date=meeting_date_obj,
                meeting_time=meeting_time_obj,
                remark=request.form.get("remark", "").strip()
            )
            db.session.add(lead)
            db.session.commit()
            flash(f"Lead {lead.lead_id} added successfully!", "success")
            return redirect(url_for("new_lead"))
        except Exception as e:
            db.session.rollback()
            flash(f"Error adding lead: {str(e)}")
    
    content = """
    <div class="card">
        <h2>Add New Lead</h2>
        <form method="post">
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <div class="form-group">
                    <label for="architect_name">Architect Name *:</label>
                    <input type="text" id="architect_name" name="architect_name" required>
                </div>
                <div class="form-group">
                    <label for="firm_name">Firm Name:</label>
                    <input type="text" id="firm_name" name="firm_name">
                </div>
                <div class="form-group">
                    <label for="grade">Grade:</label>
                    <select id="grade" name="grade">
                        <option value="A+">A+</option>
                        <option value="A">A</option>
                        <option value="B">B</option>
                        <option value="C">C</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="client_type">Client Type:</label>
                    <select id="client_type" name="client_type">
                        <option value="CRR">CRR</option>
                        <option value="NBD">NBD</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="bd_name">BD Name:</label>
                    <input type="text" id="bd_name" name="bd_name">
                </div>
                <div class="form-group">
                    <label for="meeting_date">Meeting Date:</label>
                    <input type="date" id="meeting_date" name="meeting_date">
                </div>
                <div class="form-group">
                    <label for="meeting_time">Meeting Time:</label>
                    <input type="time" id="meeting_time" name="meeting_time">
                </div>
            </div>
            <div class="form-group">
                <label for="remark">Remark:</label>
                <textarea id="remark" name="remark" rows="4"></textarea>
            </div>
            <button type="submit">Add Lead</button>
        </form>
    </div>
    """
    
    return render_page(content, "New Lead - CRM")

# -------------------------------
# ASSIGN LEAD
# -------------------------------
@app.route("/assign_lead", methods=["GET", "POST"])
@login_required
def assign_lead():
    leads = Lead.query.filter_by(assigned_to=None).all()
    
    if request.method == "POST":
        try:
            lead_id = request.form.get("lead_id")
            assigned_to = request.form.get("assigned_to", "").strip()
            
            if not assigned_to:
                flash("Please enter the name to assign the lead to")
            else:
                lead = Lead.query.filter_by(lead_id=lead_id).first()
                if lead:
                    lead.assigned_to = assigned_to
                    db.session.commit()
                    flash(f"Lead {lead_id} assigned to {assigned_to} successfully!", "success")
                else:
                    flash("Lead not found")
                return redirect(url_for("assign_lead"))
        except Exception as e:
            db.session.rollback()
            flash(f"Error assigning lead: {str(e)}")
    
    content = f"""
    <div class="card">
        <h2>Assign Lead</h2>
        {f'<p><strong>Unassigned Leads:</strong> {len(leads)}</p>' if leads else '<p>No unassigned leads available.</p>'}
        
        {'<form method="post">' if leads else ''}
    """
    
    if leads:
        content += """
            <div class="form-group">
                <label for="lead_id">Select Lead:</label>
                <select id="lead_id" name="lead_id" required>
        """
        for lead in leads:
            content += f'<option value="{lead.lead_id}">{lead.lead_id} - {lead.architect_name or "Unknown"} ({lead.firm_name or "No Firm"})</option>'
        
        content += """
                </select>
            </div>
            <div class="form-group">
                <label for="assigned_to">Assign To *:</label>
                <input type="text" id="assigned_to" name="assigned_to" required placeholder="Enter name">
            </div>
            <button type="submit">Assign Lead</button>
        </form>
        """
    
    content += "</div>"
    
    return render_page(content, "Assign Lead - CRM")

# -------------------------------
# RESCHEDULE MEETING
# -------------------------------
@app.route("/reschedule_meeting", methods=["GET", "POST"])
@login_required
def reschedule_meeting():
    leads = Lead.query.all()
    
    if request.method == "POST":
        try:
            lead_id = request.form.get("lead_id")
            lead = Lead.query.filter_by(lead_id=lead_id).first()
            
            if lead:
                reschedule_date = request.form.get("reschedule_date")
                reschedule_time = request.form.get("reschedule_time")
                
                reschedule_date_obj = datetime.strptime(reschedule_date, "%Y-%m-%d").date() if reschedule_date else None
                reschedule_time_obj = datetime.strptime(reschedule_time, "%H:%M").time() if reschedule_time else None
                
                lead.reschedule_date = reschedule_date_obj
                lead.reschedule_time = reschedule_time_obj
                lead.reschedule_remark = request.form.get("remark", "").strip()
                db.session.commit()
                flash(f"Meeting for lead {lead_id} rescheduled successfully!", "success")
            else:
                flash("Lead not found")
            return redirect(url_for("reschedule_meeting"))
        except Exception as e:
            db.session.rollback()
            flash(f"Error rescheduling meeting: {str(e)}")
    
    content = """
    <div class="card">
        <h2>Reschedule Meeting</h2>
        <form method="post">
            <div class="form-group">
                <label for="lead_id">Select Lead:</label>
                <select id="lead_id" name="lead_id" required>
    """
    
    for lead in leads:
        content += f'<option value="{lead.lead_id}">{lead.lead_id} - {lead.architect_name or "Unknown"}</option>'
    
    content += """
                </select>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <div class="form-group">
                    <label for="reschedule_date">New Meeting Date:</label>
                    <input type="date" id="reschedule_date" name="reschedule_date">
                </div>
                <div class="form-group">
                    <label for="reschedule_time">New Meeting Time:</label>
                    <input type="time" id="reschedule_time" name="reschedule_time">
                </div>
            </div>
            <div class="form-group">
                <label for="remark">Remark:</label>
                <textarea id="remark" name="remark" rows="4" placeholder="Reason for rescheduling..."></textarea>
            </div>
            <button type="submit">Reschedule Meeting</button>
        </form>
    </div>
    """
    
    return render_page(content, "Reschedule Meeting - CRM")

# -------------------------------
# MEETING STATS
# -------------------------------
@app.route("/meeting_stats", methods=["GET", "POST"])
@login_required
def meeting_stats():
    leads = Lead.query.all()
    
    if request.method == "POST":
        try:
            lead_id = request.form.get("lead_id")
            lead = Lead.query.filter_by(lead_id=lead_id).first()
            
            if lead:
                lead.not_interested = "not_interested" in request.form
                lead.require_letter = "require_letter" in request.form
                lead.email_catalogue = "email_catalogue" in request.form
                lead.quotation_sent = "quotation_sent" in request.form
                db.session.commit()
                flash(f"Meeting stats for lead {lead_id} updated successfully!", "success")
            else:
                flash("Lead not found")
            return redirect(url_for("meeting_stats"))
        except Exception as e:
            db.session.rollback()
            flash(f"Error updating meeting stats: {str(e)}")
    
    content = """
    <div class="card">
        <h2>Update Meeting Stats</h2>
        <form method="post">
            <div class="form-group">
                <label for="lead_id">Select Lead:</label>
                <select id="lead_id" name="lead_id" required>
    """
    
    for lead in leads:
        content += f'<option value="{lead.lead_id}">{lead.lead_id} - {lead.architect_name or "Unknown"}</option>'
    
    content += """
                </select>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <div>
                    <label><input type="checkbox" name="not_interested" style="width: auto; margin-right: 10px;"> Not Interested</label>
                </div>
                <div>
                    <label><input type="checkbox" name="require_letter" style="width: auto; margin-right: 10px;"> Require Letter</label>
                </div>
                <div>
                    <label><input type="checkbox" name="email_catalogue" style="width: auto; margin-right: 10px;"> Email Catalogue</label>
                </div>
                <div>
                    <label><input type="checkbox" name="quotation_sent" style="width: auto; margin-right: 10px;"> Quotation Sent</label>
                </div>
            </div>
            <button type="submit" style="margin-top: 20px;">Update Stats</button>
        </form>
    </div>
    """
    
    return render_page(content, "Meeting Stats - CRM")

# -------------------------------
# EXPORT DATA
# -------------------------------
@app.route("/export_data", methods=["GET", "POST"])
@login_required
def export_data():
    if request.method == "POST":
        try:
            export_type = request.form.get("export_type", "all")
            start_date = request.form.get("start_date")
            end_date = request.form.get("end_date")
            
            # Build query based on filters
            query = Lead.query
            
            if export_type == "date_range" and start_date and end_date:
                start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
                end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
                query = query.filter(Lead.meeting_date.between(start_date_obj, end_date_obj))
            elif export_type == "created_range" and start_date and end_date:
                start_datetime = datetime.strptime(start_date, "%Y-%m-%d")
                end_datetime = datetime.strptime(end_date, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
                query = query.filter(Lead.created_at.between(start_datetime, end_datetime))
            
            leads = query.order_by(Lead.created_at.desc()).all()
            
            if not leads:
                flash("No data found for the selected criteria")
                return redirect(url_for("export_data"))
            
            # Create Excel file
            output = BytesIO()
            
            # Prepare data for Excel
            data = []
            for lead in leads:
                data.append({
                    'Lead ID': lead.lead_id,
                    'Architect Name': lead.architect_name or '',
                    'Firm Name': lead.firm_name or '',
                    'Grade': lead.grade or '',
                    'Client Type': lead.client_type or '',
                    'BD Name': lead.bd_name or '',
                    'Meeting Date': lead.meeting_date.strftime("%Y-%m-%d") if lead.meeting_date else '',
                    'Meeting Time': lead.meeting_time.strftime("%H:%M") if lead.meeting_time else '',
                    'Remark': lead.remark or '',
                    'Assigned To': lead.assigned_to or '',
                    'Reschedule Date': lead.reschedule_date.strftime("%Y-%m-%d") if lead.reschedule_date else '',
                    'Reschedule Time': lead.reschedule_time.strftime("%H:%M") if lead.reschedule_time else '',
                    'Reschedule Remark': lead.reschedule_remark or '',
                    'Not Interested': 'Yes' if lead.not_interested else 'No',
                    'Require Letter': 'Yes' if lead.require_letter else 'No',
                    'Email Catalogue': 'Yes' if lead.email_catalogue else 'No',
                    'Quotation Sent': 'Yes' if lead.quotation_sent else 'No',
                    'Created At': lead.created_at.strftime("%Y-%m-%d %H:%M:%S") if lead.created_at else '',
                    'Updated At': lead.updated_at.strftime("%Y-%m-%d %H:%M:%S") if lead.updated_at else ''
                })
            
            # Create DataFrame and Excel file
            df = pd.DataFrame(data)
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Leads Data', index=False)
                
                # Get workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Leads Data']
                
                # Add formatting
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Write headers with formatting
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Auto-adjust column widths
                for i, col in enumerate(df.columns):
                    max_length = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(i, i, min(max_length, 50))
            
            output.seek(0)
            
            # Generate filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if export_type == "date_range":
                filename = f"leads_data_{start_date}_to_{end_date}_{timestamp}.xlsx"
            elif export_type == "created_range":
                filename = f"leads_created_{start_date}_to_{end_date}_{timestamp}.xlsx"
            else:
                filename = f"all_leads_data_{timestamp}.xlsx"
            
            response = make_response(output.read())
            response.headers['Content-Disposition'] = f'attachment; filename={filename}'
            response.headers['Content-type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            
            return response
            
        except Exception as e:
            flash(f"Error exporting data: {str(e)}")
    
    # Get statistics for display
    total_leads = Lead.query.count()
    today = date.today()
    this_month_leads = Lead.query.filter(
        Lead.created_at >= datetime(today.year, today.month, 1)
    ).count()
    
    content = f"""
    <div class="card">
        <h2>Export Data to Excel</h2>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 20px 0;">
            <div style="background: #3498db; color: white; padding: 20px; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0;">{total_leads}</h3>
                <p style="margin: 5px 0 0 0;">Total Leads</p>
            </div>
            <div style="background: #27ae60; color: white; padding: 20px; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0;">{this_month_leads}</h3>
                <p style="margin: 5px 0 0 0;">This Month</p>
            </div>
        </div>
        
        <form method="post">
            <div class="form-group">
                <label for="export_type">Export Type:</label>
                <select id="export_type" name="export_type" onchange="toggleDateFields()" required>
                    <option value="all">All Leads</option>
                    <option value="date_range">By Meeting Date Range</option>
                    <option value="created_range">By Creation Date Range</option>
                </select>
            </div>
            
            <div id="date_fields" style="display: none;">
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                    <div class="form-group">
                        <label for="start_date">Start Date:</label>
                        <input type="date" id="start_date" name="start_date">
                    </div>
                    <div class="form-group">
                        <label for="end_date">End Date:</label>
                        <input type="date" id="end_date" name="end_date">
                    </div>
                </div>
            </div>
            
            <button type="submit" style="background: #27ae60;">
                📤 Export to Excel
            </button>
        </form>
        
        <div style="margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 5px; border-left: 4px solid #3498db;">
            <h4 style="margin: 0 0 10px 0;">Export Information:</h4>
            <ul style="margin: 0; padding-left: 20px;">
                <li><strong>All Leads:</strong> Exports all leads data</li>
                <li><strong>By Meeting Date Range:</strong> Exports leads with meeting dates in the specified range</li>
                <li><strong>By Creation Date Range:</strong> Exports leads created in the specified date range</li>
            </ul>
            <p style="margin: 10px 0 0 0;"><strong>Note:</strong> The Excel file will include all lead details, meeting information, and status updates.</p>
        </div>
    </div>
    
    <script>
        function toggleDateFields() {{
            const exportType = document.getElementById('export_type').value;
            const dateFields = document.getElementById('date_fields');
            const startDate = document.getElementById('start_date');
            const endDate = document.getElementById('end_date');
            
            if (exportType === 'all') {{
                dateFields.style.display = 'none';
                startDate.required = false;
                endDate.required = false;
            }} else {{
                dateFields.style.display = 'block';
                startDate.required = true;
                endDate.required = true;
            }}
        }}
        
        // Set today's date as default
        const today = new Date().toISOString().split('T')[0];
        document.getElementById('end_date').value = today;
        
        // Set start date to 30 days ago
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        document.getElementById('start_date').value = thirtyDaysAgo.toISOString().split('T')[0];
    </script>
    """
    
    return render_page(content, "Export Data - CRM")

# -------------------------------
# MANAGE USERS
# -------------------------------
@app.route("/manage_users", methods=["GET", "POST"])
@login_required
@admin_required
def manage_users():
    users = User.query.all()
    
    if request.method == "POST":
        try:
            username = request.form.get("username", "").strip()
            password = request.form.get("password")
            role = request.form.get("role")
            
            if not username or not password:
                flash("Username and password are required")
            elif User.query.filter_by(username=username).first():
                flash("Username already exists")
            else:
                user = User(username=username, role=role)
                user.set_password(password)
                db.session.add(user)
                db.session.commit()
                flash(f"User {username} added successfully!", "success")
                return redirect(url_for("manage_users"))
        except Exception as e:
            db.session.rollback()
            flash(f"Error adding user: {str(e)}")
    
    content = """
    <div class="card">
        <h2>Manage Users</h2>
        <form method="post">
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px;">
                <div class="form-group">
                    <label for="username">Username *:</label>
                    <input type="text" id="username" name="username" required>
                </div>
                <div class="form-group">
                    <label for="password">Password *:</label>
                    <input type="password" id="password" name="password" required>
                </div>
                <div class="form-group">
                    <label for="role">Role:</label>
                    <select id="role" name="role">
                        <option value="admin">Admin</option>
                        <option value="bd">BD</option>
                        <option value="user">User</option>
                    </select>
                </div>
            </div>
            <button type="submit">Add User</button>
        </form>
    </div>
    
    <div class="card">
        <h3>Existing Users</h3>
        <table>
            <thead>
                <tr>
                    <th>Username</th>
                    <th>Role</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for user in users:
        content += f"""
                <tr>
                    <td>{user.username}</td>
                    <td>{user.role.upper()}</td>
                    <td>
                        {"Protected Admin" if user.username == "admin" else "Active"}
                    </td>
                </tr>
        """
    
    content += """
            </tbody>
        </table>
    </div>
    """
    
    return render_page(content, "Manage Users - CRM")

# -------------------------------
# ERROR HANDLERS
# -------------------------------
@app.errorhandler(404)
def not_found(error):
    content = """
    <div class="card" style="text-align: center;">
        <h2>Page Not Found</h2>
        <p>The page you're looking for doesn't exist.</p>
        <a href="/dashboard" style="color: #3498db;">Go to Dashboard</a>
    </div>
    """
    return render_page(content, "404 - Page Not Found"), 404

@app.errorhandler(500)
def internal_error(error):
    db.session.rollback()
    content = """
    <div class="card" style="text-align: center;">
        <h2>Internal Server Error</h2>
        <p>Something went wrong on our end.</p>
        <a href="/dashboard" style="color: #3498db;">Go to Dashboard</a>
    </div>
    """
    return render_page(content, "500 - Internal Server Error"), 500

# -------------------------------
# RUN APP
# -------------------------------
if __name__ == "__main__":
    app.run(debug=True, port=5001, use_reloader=False)
