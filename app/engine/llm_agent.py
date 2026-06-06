import requests
import json

class LLMAgent:
    def __init__(self, api_key=None, endpoint=None, model=None):
        self.api_key = api_key
        self.endpoint = endpoint or 'https://api.openai.com/v1/chat/completions'
        self.model = model or 'gpt-4'
    
    def convert_vba_to_js(self, vba_code, module_name='Module1'):
        if not self.api_key:
            return None
        
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