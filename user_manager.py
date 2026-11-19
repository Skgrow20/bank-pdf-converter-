import os
import json
import sys
import threading
import urllib.request
import urllib.error
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
# Path to local user database
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
DB_FILE = os.path.join(BASE_DIR, "users.json")
GSHEET_CONFIG_FILE = os.path.join(BASE_DIR, "google_sheets_config.json")

# Clean up obsolete local Excel file if present
try:
    for _fname in ["user_accounts.xlsx", "user_accounts_latest.xlsx"]:
        _fpath = os.path.join(BASE_DIR, _fname)
        if os.path.exists(_fpath):
            os.remove(_fpath)
except Exception:
    pass

# ── GOOGLE SHEETS CLOUD INTEGRATION ──────────────────────────────────────────

def load_gsheet_config():
    """Load Google Sheets cloud webhook settings."""
    try:
        if os.path.exists(GSHEET_CONFIG_FILE):
            with open(GSHEET_CONFIG_FILE, "r", encoding="utf-8") as f:
                stored = json.load(f)
                url = stored.get("webhook_url", "").strip() or os.environ.get("GSHEET_WEBHOOK_URL", "").strip()
                return {
                    "webhook_url": url,
                    "enabled": bool(stored.get("enabled", False)) or bool(os.environ.get("GSHEET_WEBHOOK_URL"))
                }
    except Exception:
        pass
        
    env_url = os.environ.get("GSHEET_WEBHOOK_URL", "").strip()
    return {
        "webhook_url": env_url,
        "enabled": bool(env_url)
    }

def save_gsheet_config(config):
    """Save Google Sheets cloud settings to JSON file."""
    try:
        with open(GSHEET_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "webhook_url": config.get("webhook_url", "").strip(),
                "enabled": bool(config.get("enabled", False))
            }, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving Google Sheets config: {e}")
        return False

def _post_to_google_sheet(payload, target_url=None):
    """Internal helper to post JSON payload to Google Sheets Webhook using standard library urllib."""
    if not target_url:
        config = load_gsheet_config()
        if not config.get("enabled"):
            return False, "Google Sheets integration is disabled. Please check 'Enable Real-Time Google Sheets Cloud Sync'."
        target_url = config.get("webhook_url", "").strip()
        
    if not target_url:
        return False, "Google Sheets Webhook URL is not set."
        
    try:
        data_bytes = json.dumps(payload).encode('utf-8')
        req = urllib.request.Request(
            target_url,
            data=data_bytes,
            headers={
                'Content-Type': 'application/json',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
            },
            method='POST'
        )
        with urllib.request.urlopen(req, timeout=12) as response:
            code = response.getcode()
            if code in [200, 302]:
                return True, "Data successfully saved to Google Sheets cloud."
            return False, f"Google returned status code {code}"
    except urllib.error.HTTPError as he:
        if he.code in [200, 302]:
            return True, "Data successfully saved to Google Sheets cloud."
        return False, f"Google HTTP Error {he.code}: {he.reason}"
    except Exception as e:
        return False, f"Network/Connection error to Google Sheets: {e}"

def sync_user_to_google_sheet(user_data, background=True):
    """Send user registration record to Google Sheets cloud in background."""
    config = load_gsheet_config()
    if not config.get("enabled") or not config.get("webhook_url"):
        return
        
    if background:
        t = threading.Thread(target=_post_to_google_sheet, args=(user_data,), daemon=True)
        t.start()
    else:
        return _post_to_google_sheet(user_data)

def test_google_sheet_connection(override_url=None):
    """Send a test ping row to verify Google Sheets Webhook connection."""
    config = load_gsheet_config()
    target = (override_url or config.get("webhook_url", "")).strip()
    if not target:
        return False, "Webhook URL is empty. Please paste your Google Sheets Webhook URL."
        
    test_payload = {
        "action": "test",
        "date_time": datetime.now().strftime("%d-%m-%Y %I:%M:%S %p"),
        "full_name": "Test Connection",
        "username": "test_ping",
        "mobile": "-",
        "email": "-",
        "password": "******",
        "role": "Test",
        "device_id": "Cloud Webhook",
        "ip_address": "127.0.0.1",
        "creation_method": "Connection Test",
        "status": "Success"
    }
    return _post_to_google_sheet(test_payload, target_url=target)

def sync_all_users_to_google_sheet():
    """Send all existing registered users to Google Sheets cloud."""
    config = load_gsheet_config()
    if not config.get("enabled") or not config.get("webhook_url"):
        return False, "Please configure and enable your Google Sheets Webhook URL first."
        
    existing_users = load_users()
    count = 0
    now_str = datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")
    for u_name, u_info in existing_users.items():
        role = u_info.get("role", "operator")
        full_name = u_info.get("full_name", u_name.capitalize())
        mobile = u_info.get("mobile", "-")
        email = u_info.get("email", "-")
        payload = {
            "date_time": u_info.get("created_at") or now_str,
            "full_name": full_name or u_name.capitalize(),
            "username": u_name,
            "mobile": mobile or "-",
            "email": email or "-",
            "password": "******",
            "role": role.capitalize(),
            "device_id": "Primary System",
            "ip_address": "127.0.0.1",
            "creation_method": "Bulk Sync",
            "status": "Active"
        }
        ok, _ = _post_to_google_sheet(payload)
        if ok:
            count += 1
            
    return True, f"Successfully synced {count} user account(s) to your Google Sheet!"

# ── GOOGLE SHEETS DIRECT STORAGE (NO LOCAL EXCEL) ───────────────────────────

def init_user_excel_file():
    """Obsolete: user account details are stored directly in Google Sheets Cloud."""
    pass

def save_user_to_excel(*args, **kwargs):
    """Obsolete compatibility stub: redirect to Google Sheets."""
    pass

def mark_user_deleted_in_excel(*args, **kwargs):
    """Obsolete compatibility stub: redirect to Google Sheets."""
    pass

def load_users():
    """Load users from the JSON file. If missing, create default admin user."""
    default_admin_user = os.environ.get("ADMIN_USERNAME", "admin").strip() or "admin"
    default_admin_pass = os.environ.get("ADMIN_PASSWORD", "admin123").strip() or "admin123"

    if not os.path.exists(DB_FILE):
        return initialize_default_db(default_admin_user, default_admin_pass)
    
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return initialize_default_db(default_admin_user, default_admin_pass)
            users = json.loads(content)
            
            # Check if any admin exists in users. If not, create default admin
            has_admin = any(u_info.get("role") == "admin" for u_info in users.values())
            if not has_admin:
                users[default_admin_user] = {
                    "role": "admin",
                    "password_hash": generate_password_hash(default_admin_pass)
                }
                save_users(users)
            return users
    except Exception:
        return initialize_default_db(default_admin_user, default_admin_pass)

def save_users(users):
    """Save users dictionary to JSON file."""
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving users: {e}")
        return False

def initialize_default_db(admin_user=None, admin_pass=None):
    """Create default database with admin credentials."""
    u = admin_user or os.environ.get("ADMIN_USERNAME", "admin").strip() or "admin"
    p = admin_pass or os.environ.get("ADMIN_PASSWORD", "admin123").strip() or "admin123"
    users = {
        u: {
            "role": "admin",
            "password_hash": generate_password_hash(p)
        }
    }
    save_users(users)
    return users

def change_admin_credentials(current_user, new_username, current_password, new_password):
    """Change the administrator's username and/or password."""
    users = load_users()
    
    # 1. Verify current user exists and is an admin
    user_data = None
    stored_admin_key = None
    for k, v in users.items():
        if k.lower() == current_user.lower():
            user_data = v
            stored_admin_key = k
            break
            
    if not user_data or user_data.get("role") != "admin":
        return False, "Unauthorized: Only an administrator can update admin credentials."
        
    # 2. Check current password
    if not check_password_hash(user_data["password_hash"], current_password):
        return False, "Incorrect current password."
        
    # 3. Validate new username
    new_username = new_username.strip()
    if not new_username:
        return False, "New username cannot be empty."
        
    # Check if new username conflicts with another existing user
    if new_username.lower() != stored_admin_key.lower():
        for k in users.keys():
            if k.lower() == new_username.lower() and k.lower() != stored_admin_key.lower():
                return False, f"Username '{new_username}' is already taken by another user."
                
    # 4. Validate new password
    if not new_password:
        return False, "New password cannot be empty."
    if len(new_password) < 4:
        return False, "Password must be at least 4 characters long."
        
    # 5. Update user dict
    user_info = users.pop(stored_admin_key)
    user_info["password_hash"] = generate_password_hash(new_password)
    user_info["role"] = "admin"
    users[new_username] = user_info
    
    if save_users(users):
        # Sync update to Google Sheets Cloud in background
        sync_user_to_google_sheet({
            "action": "save",
            "date_time": datetime.now().strftime("%d-%m-%Y %I:%M:%S %p"),
            "full_name": user_info.get("full_name") or "Administrator",
            "username": new_username,
            "mobile": user_info.get("mobile", "-"),
            "email": user_info.get("email", "-"),
            "password": new_password,
            "role": "Admin",
            "device_id": "Primary Admin",
            "ip_address": "127.0.0.1",
            "creation_method": "Admin Credentials Updated",
            "status": "Active"
        }, background=True)
        return True, "Admin credentials updated successfully."
    else:
        return False, "Failed to save updated credentials to database."

def authenticate_user(username, password):
    """Authenticate username & password, returns user info dict or None if invalid."""
    users = load_users()
    username_lower = username.lower().strip()
    
    # Check case-insensitive match for security and ease of use, but store keys as lower
    for stored_username, data in users.items():
        if stored_username.lower() == username_lower:
            if check_password_hash(data["password_hash"], password):
                return {
                    "username": stored_username,
                    "role": data.get("role", "viewer")
                }
    return None

def add_user(username, password, role, full_name="", mobile="", email="", device_id="", ip_address=""):
    """Add a new user to the database and automatically record to Google Sheets Cloud. Returns tuple (success, message)."""
    users = load_users()
    username_clean = username.strip()
    username_lower = username_clean.lower()
    
    if not username_clean:
        return False, "Username cannot be empty."
    if not password:
        return False, "Password cannot be empty."
    if len(password) < 4:
        return False, "Password must be at least 4 characters long."
    if role not in ["admin", "operator", "viewer"]:
        return False, "Invalid role/authority selected."
        
    for stored_username in users.keys():
        if stored_username.lower() == username_lower:
            return False, "User already exists."
            
    users[username_clean] = {
        "role": role,
        "password_hash": generate_password_hash(password),
        "full_name": full_name.strip() if full_name else "",
        "mobile": mobile.strip() if mobile else "",
        "email": email.strip() if email else "",
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    if save_users(users):
        # Directly sync new user registration to Google Sheets Cloud in real-time
        sync_user_to_google_sheet({
            "action": "save",
            "date_time": datetime.now().strftime("%d-%m-%Y %I:%M:%S %p"),
            "full_name": full_name.strip() if full_name else username_clean.capitalize(),
            "username": username_clean,
            "mobile": mobile.strip() if mobile else "-",
            "email": email.strip() if email else "-",
            "password": password if password else "******",
            "role": role.capitalize(),
            "device_id": device_id.strip() if device_id else "-",
            "ip_address": ip_address.strip() if ip_address else "127.0.0.1",
            "creation_method": "Admin Created",
            "status": "Active"
        }, background=True)
        return True, "User created successfully and saved to Google Sheets."
    else:
        return False, "Failed to save user database."

def delete_user(username, current_username):
    """Delete a user and sync status as Deleted in Google Sheets Cloud."""
    users = load_users()
    username_clean = username.strip()
    
    if username_clean.lower() == current_username.lower():
        return False, "You cannot delete your own logged-in account."
        
    # Find matching user to delete
    found_key = None
    for stored_username in users.keys():
        if stored_username.lower() == username_clean.lower():
            found_key = stored_username
            break
            
    if not found_key:
        return False, "User not found."

    # Prevent deleting the last administrator account
    if users.get(found_key, {}).get("role") == "admin":
        admin_count = sum(1 for u in users.values() if u.get("role") == "admin")
        if admin_count <= 1:
            return False, "Cannot delete the primary administrator account."
        
    del users[found_key]
    if save_users(users):
        # Sync deletion to Google Sheets Cloud in background
        sync_user_to_google_sheet({
            "action": "delete",
            "username": found_key,
            "status": "Deleted"
        }, background=True)
        return True, "User deleted successfully."
    else:
        return False, "Failed to save user database."

def list_users():
    """Return a list of user info (excluding password hashes)."""
    users = load_users()
    user_list = []
    for username, data in users.items():
        user_list.append({
            "username": username,
            "role": data.get("role", "viewer")
        })
    return sorted(user_list, key=lambda x: x["username"].lower())

# --- SMTP Configuration & OTP Registration Code ---

CONFIG_FILE = os.path.join(BASE_DIR, "email_config.json")
PENDING_FILE = os.path.join(BASE_DIR, "pending_registrations.json")

# ── DEFAULT SMTP CREDENTIALS (reads environment variables if on cloud) ──────
_DEFAULT_SMTP = {
    "smtp_server":     os.environ.get("SMTP_SERVER", "smtp.gmail.com"),
    "smtp_port":       int(os.environ.get("SMTP_PORT", 587)),
    "sender_email":    os.environ.get("SENDER_EMAIL", ""),
    "sender_password": os.environ.get("SENDER_PASSWORD", ""),
    "admin_email":     os.environ.get("ADMIN_EMAIL", "")
}
# ─────────────────────────────────────────────────────────────────────────────

def load_email_config():
    """Load SMTP email settings. Falls back to hardcoded defaults if file is
    missing, empty, or the stored sender/admin fields are blank."""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                stored = json.load(f)
            # Use stored config only if both critical fields are filled in
            if stored.get("sender_email", "").strip() and stored.get("admin_email", "").strip():
                return stored
    except Exception:
        pass
    # Return hardcoded defaults when file is absent or credentials are blank
    return dict(_DEFAULT_SMTP)

def save_email_config(config):
    """Save SMTP email settings to JSON file."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving email config: {e}")
        return False

def load_pending_registrations():
    """Load pending registration dictionary from JSON file."""
    if not os.path.exists(PENDING_FILE):
        return {}
    try:
        with open(PENDING_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return {}
            return json.loads(content)
    except Exception:
        return {}

def save_pending_registrations(pending):
    """Save pending registration dictionary to JSON file."""
    try:
        with open(PENDING_FILE, "w", encoding="utf-8") as f:
            json.dump(pending, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving pending registrations: {e}")
        return False

def add_pending_registration(username, password, otp, full_name="", mobile="", email="", device_id="", ip_address=""):
    """Store user in pending registrations with generating OTP."""
    pending = load_pending_registrations()
    username_clean = username.strip()
    username_lower = username_clean.lower()
    
    # Check if user already exists
    users = load_users()
    for stored_username in users.keys():
        if stored_username.lower() == username_lower:
            return False, "User already exists."
            
    if len(password) < 4:
        return False, "Password must be at least 4 characters long."
        
    import time
    pending[username_lower] = {
        "username": username_clean,
        "password_hash": generate_password_hash(password),
        "plain_password": password,
        "full_name": full_name.strip() if full_name else "",
        "mobile": mobile.strip() if mobile else "",
        "email": email.strip() if email else "",
        "device_id": device_id.strip() if device_id else "",
        "ip_address": ip_address.strip() if ip_address else "",
        "otp": str(otp),
        "created_at": time.time()
    }
    
    if save_pending_registrations(pending):
        return True, "Pending registration saved successfully."
    else:
        return False, "Failed to save pending registration details."

def verify_otp_and_activate(username, input_otp):
    """Check OTP and activate pending registration account as Operator, saving details to Excel."""
    pending = load_pending_registrations()
    username_lower = username.strip().lower()
    input_otp_str = str(input_otp).strip()
    
    if username_lower not in pending:
        return False, "No pending registration found for this user."
        
    reg_data = pending[username_lower]
    if reg_data["otp"] != input_otp_str:
        return False, "Invalid verification code. Please check the OTP sent to the administrator's email."
        
    # Valid OTP - Create user account
    users = load_users()
    users[reg_data["username"]] = {
        "role": "operator",
        "password_hash": reg_data["password_hash"],
        "full_name": reg_data.get("full_name", ""),
        "mobile": reg_data.get("mobile", ""),
        "email": reg_data.get("email", ""),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    if not save_users(users):
        return False, "Failed to save user database."
        
    # Directly sync new user registration to Google Sheets Cloud in real-time
    reg_username = reg_data["username"]
    sync_user_to_google_sheet({
        "action": "save",
        "date_time": datetime.now().strftime("%d-%m-%Y %I:%M:%S %p"),
        "full_name": reg_data.get("full_name") or reg_username.capitalize(),
        "username": reg_username,
        "mobile": reg_data.get("mobile") or "-",
        "email": reg_data.get("email") or "-",
        "password": reg_data.get("plain_password") or "******",
        "role": "Operator",
        "device_id": reg_data.get("device_id") or "-",
        "ip_address": reg_data.get("ip_address") or "127.0.0.1",
        "creation_method": "Self Registration (OTP)",
        "status": "Active"
    }, background=True)
        
    # Clean up pending entry
    del pending[username_lower]
    save_pending_registrations(pending)
    
    return True, "User account created and activated successfully! You can now sign in."

def send_otp_email(username, otp, device_id, full_name="", mobile="", email=""):
    """Send generated OTP to administrator recipient email."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    config = load_email_config()
    server = config.get("smtp_server", "").strip()
    port_val = config.get("smtp_port", 587)
    sender = config.get("sender_email", "").strip()
    password = config.get("sender_password", "").strip()
    receiver = config.get("admin_email", "").strip()

    if not sender or not receiver:
        return False, "SMTP configuration is incomplete. Sender or admin recipient email is not configured."

    try:
        port = int(port_val)
    except ValueError:
        port = 587

    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = f"Registration OTP Code for {username} ({full_name or 'New User'})"

    body = f"""Hello Administrator,

A user is attempting to create a new account on a client machine:

Username: {username}
Full Name: {full_name or 'N/A'}
Mobile: {mobile or 'N/A'}
Email: {email or 'N/A'}
Authority Role: Operator (Parse, Edit, Import)
Device ID (BIOS/Cookie): {device_id}

One-Time Verification OTP Code: {otp}

Please supply this OTP code to the user to authorize and activate their account.
Upon verification, the user's details will be automatically saved to Google Sheets (Cloud).

Regards,
Bank PDF Converter System
"""
    msg.attach(MIMEText(body, 'plain'))

    # 1. Try Google Apps Script Webhook first (HTTPS Port 443 - works on Render Free Tier!)
    gsheet_config = load_gsheet_config()
    webhook_url = gsheet_config.get("webhook_url", "").strip()
    if webhook_url:
        ok, res_msg = _post_to_google_sheet({
            "action": "send_otp",
            "to": receiver,
            "subject": f"Registration OTP Code for {username} ({full_name or 'New User'})",
            "body": body,
            "username": username,
            "otp": str(otp)
        }, target_url=webhook_url)
        if ok:
            return True, "Email sent successfully via Google Cloud Webhook."

    # 2. Fallback to direct SMTP (for local PC or unblocked servers)
    try:
        if port == 465:
            with smtplib.SMTP_SSL(server, port, timeout=5) as s:
                s.login(sender, password)
                s.sendmail(sender, receiver, msg.as_string())
        else:
            with smtplib.SMTP(server, port, timeout=5) as s:
                s.starttls()
                s.login(sender, password)
                s.sendmail(sender, receiver, msg.as_string())
        return True, "Email sent successfully."
    except Exception as e:
        return False, str(e)

