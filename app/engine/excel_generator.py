import zipfile
import os
import shutil
from io import BytesIO

class ExcelGenerator:
    def __init__(self):
        self.content_types = {
            '.rels': 'application/vnd.openxmlformats-package.relationships+xml',
            '.xml': 'application/xml',
            '.bin': 'application/octet-stream'
        }
    
    def create_js_macro_workbook(self, source_path, output_path, js_modules):
        with zipfile.ZipFile(source_path, 'r') as src_zip:
            temp_dir = output_path + '_temp'
            os.makedirs(temp_dir, exist_ok=True)
            
            src_zip.extractall(temp_dir)
            
            self._create_macro_files(temp_dir, js_modules)
            self._update_content_types(temp_dir)
            self._update_relationships(temp_dir)
            
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        dst_zip.write(file_path, arcname)
            
            shutil.rmtree(temp_dir)
    
    def _create_macro_files(self, temp_dir, js_modules):
        xl_dir = os.path.join(temp_dir, 'xl')
        os.makedirs(xl_dir, exist_ok=True)
        
        js_dir = os.path.join(xl_dir, 'js')
        os.makedirs(js_dir, exist_ok=True)
        
        for module in js_modules:
            module_name = module['name'].replace('.bas', '.js').replace('.cls', '.js').replace('.frm', '.js')
            js_path = os.path.join(js_dir, module_name)
            with open(js_path, 'w', encoding='utf-8') as f:
                f.write(module['code'])
        
        workbook_json = os.path.join(xl_dir, 'workbook.json')
        with open(workbook_json, 'w', encoding='utf-8') as f:
            f.write(self._generate_workbook_json(js_modules))
    
    def _generate_workbook_json(self, js_modules):
        modules = []
        for module in js_modules:
            module_name = module['name'].replace('.bas', '.js').replace('.cls', '.js').replace('.frm', '.js')
            modules.append({
                'name': module['name'],
                'path': f'xl/js/{module_name}'
            })
        
        return f'''{{
    "version": "1.0",
    "javascript": {{
        "modules": {modules}
    }}
}}'''
    
    def _update_content_types(self, temp_dir):
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
        
        if os.path.exists(content_types_path):
            with open(content_types_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            if '<Default Extension="json" ContentType="application/json"/>' not in content:
                content = content.replace(
                    '</Types>',
                    '    <Default Extension="json" ContentType="application/json"/>\n</Types>'
                )
            
            with open(content_types_path, 'w', encoding='utf-8') as f:
                f.write(content)
    
    def _update_relationships(self, temp_dir):
        xl_rels_dir = os.path.join(temp_dir, 'xl', '_rels')
        os.makedirs(xl_rels_dir, exist_ok=True)
        
        workbook_rels_path = os.path.join(xl_rels_dir, 'workbook.xml.rels')
        
        if os.path.exists(workbook_rels_path):
            with open(workbook_rels_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            if 'workbook.json' not in content:
                content = content.replace(
                    '</Relationships>',
                    '    <Relationship Id="rIdJs" Type="http://schemas.microsoft.com/office/2006/relationships/officeDocument/javascript" Target="workbook.json"/>\n</Relationships>'
                )
            
            with open(workbook_rels_path, 'w', encoding='utf-8') as f:
                f.write(content)
        else:
            with open(workbook_rels_path, 'w', encoding='utf-8') as f:
                f.write('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rIdJs" Type="http://schemas.microsoft.com/office/2006/relationships/officeDocument/javascript" Target="workbook.json"/>
</Relationships>''')