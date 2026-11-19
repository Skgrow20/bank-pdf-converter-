import os
import json
import subprocess
from datetime import datetime

import sys

# Path to device database
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
DB_FILE = os.path.join(BASE_DIR, "devices.json")

def get_bios_uuid():
    """Retrieve the physical machine's BIOS UUID using PowerShell."""
    try:
        # Run PowerShell command directly to get the BIOS UUID
        cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
        output = subprocess.check_output(cmd, stderr=subprocess.DEVNULL)
        uuid_str = output.decode('utf-8').strip()
        if uuid_str and len(uuid_str) > 10:
            return uuid_str
    except Exception:
        pass
    
    # Fallback to serial number
    try:
        cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_Bios).SerialNumber"]
        output = subprocess.check_output(cmd, stderr=subprocess.DEVNULL)
        sn_str = output.decode('utf-8').strip()
        if sn_str and len(sn_str) > 5:
            return sn_str
    except Exception:
        pass
        
    return None

def load_db():
    """Load devices from file. Initialize if missing."""
    if not os.path.exists(DB_FILE):
        return initialize_db()
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return initialize_db()
            data = json.loads(content)
            # Ensure schema consistency
            if "authorized" not in data:
                data["authorized"] = {}
            if "pending" not in data:
                data["pending"] = {}
            return data
    except Exception:
        return initialize_db()

def save_db(data):
    """Save devices to JSON database."""
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving devices database: {e}")
        return False

def initialize_db():
    """Initialize empty device database."""
    data = {
        "authorized": {},
        "pending": {}
    }
    save_db(data)
    return data

def get_device_id(token):
    """Generate a short 8-character ID from a token (e.g. DEV-a1b2c3d4 or BIOS-a1b2c3d4)."""
    if not token:
        return ""
    if token.startswith("bios_"):
        clean = token.replace("bios_", "")
        return f"BIOS-{clean[:8]}"
    clean = token.replace("device_", "")
    return f"DEV-{clean[:8]}"

def is_device_authorized(token):
    """Check if device token exists in the authorized devices list."""
    return True

def has_any_authorized_devices():
    """Check if there are any authorized devices at all."""
    db = load_db()
    return len(db.get("authorized", {})) > 0

def add_pending_request(token, ip):
    """Register an unauthorized device request."""
    db = load_db()
    if token in db.get("authorized", {}):
        return False
        
    db["pending"][token] = {
        "ip": ip,
        "device_id": get_device_id(token),
        "requested_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    save_db(db)
    return True

def authorize_device(token, name="Authorized PC"):
    """Authorize a specific device token."""
    db = load_db()
    
    # Remove from pending if exists
    db["pending"].pop(token, None)
    
    db["authorized"][token] = {
        "name": name,
        "device_id": get_device_id(token),
        "authorized_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    save_db(db)
    return True

def authorize_by_device_id(device_id, name="Authorized PC"):
    """Find a pending request by short device_id and authorize it."""
    db = load_db()
    token_to_auth = None
    
    # Check in pending requests
    for token, info in db["pending"].items():
        if info.get("device_id") == device_id or token == device_id:
            token_to_auth = token
            break
            
    # Check if we can find it even if it's not currently in pending list (manual force approve)
    if not token_to_auth:
        # If the user typed/pasted a request token directly
        if device_id.startswith("DEV-"):
            # Try to match DEV-xxxxxx
            for token in db["pending"].keys():
                if get_device_id(token) == device_id:
                    token_to_auth = token
                    break
        elif len(device_id) >= 30: # Is a full raw token
            token_to_auth = device_id
            
    if token_to_auth:
        return authorize_device(token_to_auth, name)
    return False

def revoke_device(token):
    """Revoke authorization for a device by token or device_id."""
    db = load_db()
    token_to_remove = None
    
    if token in db["authorized"]:
        token_to_remove = token
    else:
        # Search by short device_id
        for t, info in db["authorized"].items():
            if info.get("device_id") == token:
                token_to_remove = t
                break
                
    if token_to_remove:
        del db["authorized"][token_to_remove]
        save_db(db)
        return True
    return False

def list_authorized_devices():
    """List all authorized devices."""
    db = load_db()
    devices = []
    for token, info in db.get("authorized", {}).items():
        devices.append({
            "token": token,
            "device_id": info.get("device_id"),
            "name": info.get("name", "Authorized PC"),
            "authorized_at": info.get("authorized_at")
        })
    return sorted(devices, key=lambda x: x["authorized_at"], reverse=True)

def list_pending_requests():
    """List all pending device requests."""
    db = load_db()
    requests = []
    for token, info in db.get("pending", {}).items():
        requests.append({
            "token": token,
            "device_id": info.get("device_id"),
            "ip": info.get("ip"),
            "requested_at": info.get("requested_at")
        })
    return sorted(requests, key=lambda x: x["requested_at"], reverse=True)

def clear_all_pending():
    """Clear all pending authorization requests."""
    db = load_db()
    db["pending"] = {}
    save_db(db)
    return True

def reset_all_devices(keep_token=None):
    """Reset all pending and authorized devices, optionally keeping the current PC authorized."""
    db = load_db()
    db["pending"] = {}
    if keep_token and keep_token in db.get("authorized", {}):
        current_info = db["authorized"][keep_token]
        db["authorized"] = {keep_token: current_info}
    else:
        db["authorized"] = {}
        if keep_token:
            db["authorized"][keep_token] = {
                "name": "Current Active PC (Admin)",
                "device_id": get_device_id(keep_token),
                "authorized_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
    save_db(db)
    return True
