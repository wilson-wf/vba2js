import os
import uuid
import hashlib
from datetime import datetime

def generate_file_id():
    return str(uuid.uuid4())

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in {'xlsm', 'xlsx'}

def get_file_hash(file_path):
    sha256_hash = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b""):
            sha256_hash.update(chunk)
    return sha256_hash.hexdigest()

def safe_filename(filename):
    return "".join([c for c in filename if c.isalpha() or c.isdigit() or c in (' ', '.', '_')]).rstrip()

def get_timestamp():
    return datetime.now().isoformat()

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def log_message(log_folder, message):
    log_file = os.path.join(log_folder, f"{datetime.now().strftime('%Y%m%d')}.log")
    with open(log_file, 'a') as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")