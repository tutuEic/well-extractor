"""
Machine-code license system.
Generates machine code from hardware, validates license keys.
"""
import hashlib
import uuid
import subprocess
import os
import sys
import json

LICENSE_FILE = os.path.join(os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__), '.license')

# Secret salt for key generation (change this!)
SALT = "HYDRA-WELL-2026-X8K3M"

def get_machine_code():
    """Generate unique machine code from hardware."""
    identifiers = []
    
    # 1. MAC address
    try:
        mac = uuid.getnode()
        identifiers.append(str(mac))
    except:
        pass
    
    # 2. Disk serial (Windows)
    try:
        result = subprocess.run(
            ['wmic', 'diskdrive', 'get', 'serialnumber'],
            capture_output=True, text=True, timeout=5
        )
        for line in result.stdout.split('\n'):
            line = line.strip()
            if line and 'SerialNumber' not in line:
                identifiers.append(line)
                break
    except:
        pass
    
    # 3. Motherboard serial
    try:
        result = subprocess.run(
            ['wmic', 'baseboard', 'get', 'serialnumber'],
            capture_output=True, text=True, timeout=5
        )
        for line in result.stdout.split('\n'):
            line = line.strip()
            if line and 'SerialNumber' not in line:
                identifiers.append(line)
                break
    except:
        pass
    
    # Generate machine code
    combined = '|'.join(identifiers)
    code = hashlib.sha256(combined.encode()).hexdigest()[:16].upper()
    # Format: XXXX-XXXX-XXXX-XXXX
    return '-'.join([code[i:i+4] for i in range(0, 16, 4)])


def generate_key(machine_code):
    """Generate a license key from machine code."""
    clean = machine_code.replace('-', '')
    h = hashlib.sha256((clean + SALT).encode()).hexdigest()[:16].upper()
    return '-'.join([h[i:i+4] for i in range(0, 16, 4)])


def check_license(key, machine_code):
    """Verify license key against machine code."""
    expected = generate_key(machine_code)
    return key.upper().replace('-', '') == expected.replace('-', '')


def save_license(key):
    """Save validated license to file."""
    try:
        with open(LICENSE_FILE, 'w') as f:
            json.dump({'key': key}, f)
        return True
    except:
        return False


def load_license():
    """Load saved license."""
    try:
        if os.path.exists(LICENSE_FILE):
            with open(LICENSE_FILE, 'r') as f:
                data = json.load(f)
                return data.get('key', '')
    except:
        pass
    return ''


def is_licensed():
    """Check if current machine has a valid license."""
    saved_key = load_license()
    if not saved_key:
        return False
    machine_code = get_machine_code()
    return check_license(saved_key, machine_code)


# ── CLI tool ────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        mc = get_machine_code()
        print(f'Machine Code: {mc}')
        print()
        print('Usage:')
        print(f'  python license.py          — show machine code')
        print(f'  python license.py keygen   — generate key from machine code')
        print(f'  python license.py check KEY — verify a key')
    elif sys.argv[1] == 'keygen':
        mc = get_machine_code()
        key = generate_key(mc)
        print(f'Machine Code: {mc}')
        print(f'License Key:  {key}')
    elif sys.argv[1] == 'check' and len(sys.argv) > 2:
        key = sys.argv[2]
        mc = get_machine_code()
        valid = check_license(key, mc)
        print(f'Machine Code: {mc}')
        print(f'Key:          {key}')
        print(f'Valid:        {valid}')
