import zipfile
import olefile
import os
import re
from io import BytesIO

class VBAParser:
    def __init__(self):
        self.vba_project_pattern = re.compile(r'vbaProject\.bin', re.IGNORECASE)
    
    def extract_vba_modules(self, file_path):
        modules = []
        
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                vba_bin = None
                for name in zf.namelist():
                    if self.vba_project_pattern.search(name):
                        vba_bin = zf.read(name)
                        break
                
                if vba_bin is None:
                    return modules
                
                ole = olefile.OleFileIO(BytesIO(vba_bin))
                
                for stream in ole.listdir():
                    if len(stream) >= 2 and stream[0] == 'VBA':
                        module_name = stream[1]
                        try:
                            content = ole.openstream(stream).read()
                            decoded = self._decode_vba_content(content)
                            if decoded.strip():
                                modules.append({
                                    'name': module_name,
                                    'code': decoded
                                })
                        except Exception as e:
                            continue
                
                ole.close()
        except Exception as e:
            pass
        
        return modules
    
    def _decode_vba_content(self, content):
        try:
            return content.decode('utf-16-le', errors='replace')
        except:
            try:
                return content.decode('latin-1', errors='replace')
            except:
                return str(content)
    
    def has_vba_macros(self, file_path):
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                for name in zf.namelist():
                    if self.vba_project_pattern.search(name):
                        return True
            return False
        except:
            return False