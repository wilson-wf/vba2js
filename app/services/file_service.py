import os
from app.utils.helpers import generate_file_id, allowed_file, safe_filename, ensure_dir
from app.utils.database import DatabaseManager

class FileService:
    def __init__(self, config):
        self.config = config
        ensure_dir(config.UPLOAD_FOLDER)
        ensure_dir(config.CONVERTED_FOLDER)
        self.db = DatabaseManager(os.path.join(config.BASE_DIR, 'data', 'converter.db'))
    
    def upload_file(self, file):
        if not file or not allowed_file(file.filename):
            return None, 'Invalid file type. Only .xlsm files are allowed.'
        
        file_id = generate_file_id()
        original_name = safe_filename(file.filename)
        file_path = os.path.join(self.config.UPLOAD_FOLDER, f"{file_id}.xlsm")
        
        try:
            file.save(file_path)
            self.db.add_file(file_id, original_name)
            return file_id, None
        except Exception as e:
            return None, str(e)
    
    def get_file_info(self, file_id):
        return self.db.get_file(file_id)
    
    def update_file_status(self, file_id, status, progress=0, stage='pending', error_message=None):
        self.db.update_file_status(file_id, status, progress, stage, error_message)
    
    def update_converted_path(self, file_id, converted_path):
        self.db.update_file_converted_path(file_id, converted_path)
    
    def get_upload_path(self, file_id):
        return os.path.join(self.config.UPLOAD_FOLDER, f"{file_id}.xlsm")
    
    def get_converted_path(self, file_id):
        return os.path.join(self.config.CONVERTED_FOLDER, f"{file_id}_converted.xlsx")