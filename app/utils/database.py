import sqlite3
import os
from datetime import datetime
from app.utils.helpers import ensure_dir

class DatabaseManager:
    def __init__(self, db_path):
        ensure_dir(os.path.dirname(db_path))
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id TEXT PRIMARY KEY,
                original_name TEXT NOT NULL,
                upload_time TEXT NOT NULL,
                status TEXT NOT NULL,
                converted_path TEXT,
                error_message TEXT,
                progress INTEGER DEFAULT 0,
                stage TEXT DEFAULT 'pending'
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS convert_tasks (
                id TEXT PRIMARY KEY,
                file_id TEXT NOT NULL,
                vba_modules TEXT,
                js_modules TEXT,
                stage TEXT NOT NULL,
                errors TEXT,
                FOREIGN KEY (file_id) REFERENCES files(id)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS llm_configs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                api_key TEXT,
                endpoint TEXT NOT NULL,
                model TEXT NOT NULL,
                created_at TEXT NOT NULL,
                is_active INTEGER DEFAULT 1
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def add_file(self, file_id, original_name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO files (id, original_name, upload_time, status)
            VALUES (?, ?, ?, ?)
        ''', (file_id, original_name, datetime.now().isoformat(), 'pending'))
        conn.commit()
        conn.close()
    
    def update_file_status(self, file_id, status, progress=0, stage='pending', error_message=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE files SET status = ?, progress = ?, stage = ?, error_message = ?
            WHERE id = ?
        ''', (status, progress, stage, error_message, file_id))
        conn.commit()
        conn.close()
    
    def update_file_converted_path(self, file_id, converted_path):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE files SET converted_path = ?, status = 'completed', progress = 100, stage = 'completed'
            WHERE id = ?
        ''', (converted_path, file_id))
        conn.commit()
        conn.close()
    
    def get_file(self, file_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM files WHERE id = ?', (file_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return {
                'id': row[0],
                'original_name': row[1],
                'upload_time': row[2],
                'status': row[3],
                'converted_path': row[4],
                'error_message': row[5],
                'progress': row[6],
                'stage': row[7]
            }
        return None
    
    def add_convert_task(self, task_id, file_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO convert_tasks (id, file_id, stage)
            VALUES (?, ?, ?)
        ''', (task_id, file_id, 'parsing'))
        conn.commit()
        conn.close()
    
    def update_convert_task(self, task_id, stage, vba_modules=None, js_modules=None, errors=None):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        if vba_modules is not None:
            cursor.execute('''
                UPDATE convert_tasks SET stage = ?, vba_modules = ?
                WHERE id = ?
            ''', (stage, vba_modules, task_id))
        elif js_modules is not None:
            cursor.execute('''
                UPDATE convert_tasks SET stage = ?, js_modules = ?
                WHERE id = ?
            ''', (stage, js_modules, task_id))
        elif errors is not None:
            cursor.execute('''
                UPDATE convert_tasks SET stage = ?, errors = ?
                WHERE id = ?
            ''', (stage, errors, task_id))
        else:
            cursor.execute('''
                UPDATE convert_tasks SET stage = ? WHERE id = ?
            ''', (stage, task_id))
        conn.commit()
        conn.close()
    
    def get_convert_task(self, task_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM convert_tasks WHERE id = ?', (task_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return {
                'id': row[0],
                'file_id': row[1],
                'vba_modules': row[2],
                'js_modules': row[3],
                'stage': row[4],
                'errors': row[5]
            }
        return None
    
    def save_llm_config(self, api_key, endpoint, model):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('UPDATE llm_configs SET is_active = 0 WHERE is_active = 1')
        cursor.execute('''
            INSERT INTO llm_configs (api_key, endpoint, model, created_at, is_active)
            VALUES (?, ?, ?, ?, 1)
        ''', (api_key, endpoint, model, datetime.now().isoformat()))
        conn.commit()
        conn.close()
    
    def get_active_llm_config(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT api_key, endpoint, model FROM llm_configs WHERE is_active = 1 LIMIT 1')
        row = cursor.fetchone()
        conn.close()
        if row:
            return {
                'api_key': row[0],
                'endpoint': row[1],
                'model': row[2]
            }
        return None