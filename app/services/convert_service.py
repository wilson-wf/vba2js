import os
import json
from app.engine.vba_parser import VBAParser
from app.engine.syntax_mapper import SyntaxMapper
from app.engine.llm_agent import LLMAgent
from app.engine.excel_generator import ExcelGenerator

class ConvertService:
    def __init__(self, config):
        self.config = config
        self.vba_parser = VBAParser()
        self.syntax_mapper = SyntaxMapper()
        self.excel_generator = ExcelGenerator()
        
        llm_config = self._get_llm_config()
        self.llm_agent = LLMAgent(
            api_key=llm_config.get('api_key'),
            endpoint=llm_config.get('endpoint'),
            model=llm_config.get('model')
        )
    
    def _get_llm_config(self):
        from app.utils.database import DatabaseManager
        db = DatabaseManager(os.path.join(self.config.BASE_DIR, 'data', 'converter.db'))
        config = db.get_active_llm_config()
        if config:
            return config
        
        return {
            'api_key': self.config.LLM_API_KEY,
            'endpoint': self.config.LLM_ENDPOINT,
            'model': self.config.LLM_MODEL
        }
    
    def convert(self, file_id, source_path, use_llm=False, progress_callback=None):
        stages = [
            ('parsing', '解析VBA代码', 15),
            ('converting', '转换语法', 40),
            ('testing', '测试验证', 25),
            ('generating', '生成文件', 20)
        ]
        
        try:
            if progress_callback:
                progress_callback(0, 'starting', '开始转换')
            
            vba_modules = self.vba_parser.extract_vba_modules(source_path)
            if not vba_modules:
                return None, '未找到VBA宏代码'
            
            if progress_callback:
                progress_callback(15, 'parsing', 'VBA代码解析完成')
            
            js_modules = []
            for module in vba_modules:
                if use_llm and self.llm_agent.api_key:
                    js_code = self.llm_agent.convert_vba_to_js(module['code'], module['name'])
                    if js_code:
                        js_modules.append({'name': module['name'], 'code': js_code})
                    else:
                        js_code = self.syntax_mapper.convert(module['code'])
                        js_modules.append({'name': module['name'], 'code': js_code})
                else:
                    js_code = self.syntax_mapper.convert(module['code'])
                    js_modules.append({'name': module['name'], 'code': js_code})
            
            if progress_callback:
                progress_callback(55, 'converting', '语法转换完成')
            
            all_js_code = '\n\n'.join([m['code'] for m in js_modules])
            validation = self.llm_agent.validate_js_code(all_js_code)
            
            if progress_callback:
                progress_callback(80, 'testing', '代码验证完成')
            
            output_path = os.path.join(
                self.config.CONVERTED_FOLDER, 
                f"{file_id}_converted.xlsx"
            )
            
            self.excel_generator.create_js_macro_workbook(source_path, output_path, js_modules)
            
            if progress_callback:
                progress_callback(100, 'generating', '文件生成完成')
            
            return output_path, None
            
        except Exception as e:
            return None, str(e)
    
    def validate_js_code(self, js_code):
        return self.llm_agent.validate_js_code(js_code)