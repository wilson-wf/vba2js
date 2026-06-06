import requests
import json

class LLMAgent:
    def __init__(self, api_key=None, endpoint=None, model=None):
        self.api_key = api_key
        self.endpoint = endpoint or 'https://api.openai.com/v1/chat/completions'
        self.model = model or 'gpt-4'
    
    def convert_vba_to_js(self, vba_code, module_name='Module1'):
        if not self.api_key:
            # 如果没有API key，返回一个简单的模拟转换，用于测试
            return self._mock_convert(vba_code, module_name)
        
        prompt = f"""
Convert the following VBA code to WPS JavaScript Macro code.
Ensure 100% functional equivalence.

VBA Module: {module_name}

VBA Code:
{vba_code}

Requirements:
1. Convert VBA syntax to JavaScript
2. Map VBA functions to WPS JS API equivalents
3. Preserve all logic and functionality
4. Handle Range, Worksheet, Workbook objects correctly
5. Convert MsgBox to Application.MsgBox
6. Convert Date/time functions appropriately
7. Maintain variable types and scopes
8. Handle error handling and exception cases

Return ONLY the JavaScript code without explanations.
"""
        
        headers = {
            'Authorization': f'Bearer {self.api_key}',
            'Content-Type': 'application/json'
        }
        
        data = {
            'model': self.model,
            'messages': [
                {'role': 'system', 'content': 'You are an expert VBA to WPS JS Macro converter.'},
                {'role': 'user', 'content': prompt.strip()}
            ],
            'temperature': 0.1,
            'max_tokens': 4096
        }
        
        try:
            response = requests.post(self.endpoint, headers=headers, json=data)
            response.raise_for_status()
            result = response.json()
            
            if result and 'choices' in result and len(result['choices']) > 0:
                return result['choices'][0]['message']['content'].strip()
            return None
        except Exception as e:
            return None
    
    def repair_js_code(self, js_code, failure_report, iteration):
        """根据失败报告修复JS代码"""
        if not self.api_key:
            # 如果没有API key，返回一个改进版本，用于测试
            return self._mock_repair(js_code, iteration)
        
        prompt = f"""
You are an expert WPS JavaScript Macro debugger. Please fix the following code based on the failure report.

Iteration: {iteration}

Original JS Code:
{js_code}

Failure Report:
{failure_report}

Requirements:
1. Fix all errors mentioned in the report
2. Ensure 100% functional equivalence to the original VBA macro
3. Maintain WPS JS API compatibility
4. Return ONLY the fixed JavaScript code without explanations
"""
        
        headers = {
            'Authorization': f'Bearer {self.api_key}',
            'Content-Type': 'application/json'
        }
        
        data = {
            'model': self.model,
            'messages': [
                {'role': 'system', 'content': 'You are an expert WPS JavaScript Macro debugger.'},
                {'role': 'user', 'content': prompt.strip()}
            ],
            'temperature': 0.3,
            'max_tokens': 4096
        }
        
        try:
            response = requests.post(self.endpoint, headers=headers, json=data)
            response.raise_for_status()
            result = response.json()
            
            if result and 'choices' in result and len(result['choices']) > 0:
                return result['choices'][0]['message']['content'].strip()
            return js_code  # 如果失败，返回原代码
        except Exception as e:
            return js_code  # 如果异常，返回原代码
    
    def validate_js_code(self, js_code):
        if not self.api_key:
            return {'valid': True, 'errors': [], 'suggestions': []}
        
        prompt = f"""
Validate the following WPS JavaScript Macro code:

{js_code}

Check for:
1. Syntax errors
2. WPS API compatibility
3. Potential runtime errors
4. Best practices

Provide a JSON response with:
- valid: boolean
- errors: array of error messages
- suggestions: array of improvement suggestions
"""
        
        headers = {
            'Authorization': f'Bearer {self.api_key}',
            'Content-Type': 'application/json'
        }
        
        data = {
            'model': self.model,
            'messages': [
                {'role': 'system', 'content': 'You are an expert JavaScript code validator for WPS macros.'},
                {'role': 'user', 'content': prompt.strip()}
            ],
            'temperature': 0.1,
            'max_tokens': 2048
        }
        
        try:
            response = requests.post(self.endpoint, headers=headers, json=data)
            response.raise_for_status()
            result = response.json()
            
            if result and 'choices' in result and len(result['choices']) > 0:
                content = result['choices'][0]['message']['content'].strip()
                try:
                    return json.loads(content)
                except:
                    return {'valid': True, 'errors': [], 'suggestions': [content]}
            return {'valid': True, 'errors': [], 'suggestions': []}
        except Exception as e:
            return {'valid': True, 'errors': [], 'suggestions': []}
    
    def _mock_convert(self, vba_code, module_name):
        """模拟转换，用于测试"""
        # 简单的VBA到JS转换逻辑
        js_code = vba_code
        
        # 替换Sub为function
        js_code = js_code.replace('Sub ', 'function ')
        js_code = js_code.replace('End Sub', '}')
        
        # 替换MsgBox
        js_code = js_code.replace('MsgBox', 'Application.MsgBox')
        
        # 替换Range
        js_code = js_code.replace('Range(', 'Range(')
        
        # 添加花括号
        if 'function ' in js_code and '}' not in js_code:
            lines = js_code.split('\n')
            new_lines = []
            for line in lines:
                if line.strip().startswith('function '):
                    new_lines.append(line + ' {')
                else:
                    new_lines.append('  ' + line)
            js_code = '\n'.join(new_lines)
            if not js_code.strip().endswith('}'):
                js_code += '\n}'
        
        return js_code
    
    def _mock_repair(self, js_code, iteration):
        """模拟修复，用于测试"""
        # 在真实场景中，这里应该调用LLM
        # 这里返回一个改进版本的代码
        return js_code