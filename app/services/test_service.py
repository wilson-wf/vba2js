import subprocess
import os
import json

class TestService:
    def __init__(self, config):
        self.config = config
    
    def run_syntax_check(self, js_code):
        try:
            test_script = os.path.join(self.config.BASE_DIR, 'tests', 'syntax_check.js')
            
            with open(test_script, 'w') as f:
                f.write(f'''
try {{
    new Function({json.dumps(js_code)});
    console.log(JSON.stringify({{valid: true, errors: []}}));
}} catch (e) {{
    console.log(JSON.stringify({{valid: false, errors: [e.message]}}));
}}
''')
            
            result = subprocess.run(
                ['node', test_script],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            try:
                return json.loads(result.stdout)
            except:
                return {'valid': False, 'errors': ['Syntax check failed']}
        except Exception as e:
            return {'valid': False, 'errors': [str(e)]}
    
    def run_api_validation(self, js_code):
        wps_api_patterns = [
            r'\bRange\s*\(',
            r'\bWorksheets\s*\(',
            r'\bWorkbook\s*\(',
            r'\bApplication\s*\.',
            r'\bThisWorkbook\b',
            r'\bActiveSheet\b',
            r'\bActiveCell\b',
            r'\bMsgBox\s*\(',
            r'\bInputBox\s*\(',
            r'\bCells\s*\(',
            r'\bRows\s*\(',
            r'\bColumns\s*\(',
            r'\bSheets\s*\(',
            r'\bCharts\s*\(',
            r'\bNames\s*\(',
            r'\bPivotTables\s*\(',
            r'\bFormulas\s*\(',
            r'\bRange\b\s*\.'
        ]
        
        errors = []
        suggestions = []
        
        for pattern in wps_api_patterns:
            if pattern in js_code:
                suggestions.append(f"Found WPS API usage: {pattern}")
        
        return {
            'valid': True,
            'errors': errors,
            'suggestions': suggestions
        }
    
    def run_full_test(self, js_code):
        syntax_result = self.run_syntax_check(js_code)
        api_result = self.run_api_validation(js_code)
        
        return {
            'syntax_valid': syntax_result['valid'],
            'syntax_errors': syntax_result['errors'],
            'api_valid': api_result['valid'],
            'api_errors': api_result['errors'],
            'api_suggestions': api_result['suggestions'],
            'overall_valid': syntax_result['valid'] and api_result['valid']
        }