import re

class SyntaxMapper:
    def __init__(self):
        self.type_mappings = {
            'Integer': 'number',
            'Long': 'number',
            'Double': 'number',
            'String': 'string',
            'Boolean': 'boolean',
            'Object': 'object',
            'Variant': 'any'
        }
        
        self.keyword_mappings = {
            'Sub': 'function',
            'Function': 'function',
            'End': '',
            'Dim': 'let',
            'As': '',
            'Set': '',
            'Nothing': 'null',
            'True': 'true',
            'False': 'false',
            'If': 'if',
            'Then': '',
            'Else': 'else',
            'ElseIf': 'else if',
            'For': 'for',
            'Next': '',
            'Do': 'do',
            'Loop': '',
            'While': 'while',
            'Wend': '',
            'Select': 'switch',
            'Case': 'case',
            'To': '',
            'Step': '',
            'Exit': 'return',
            'On': '',
            'Error': '',
            'Resume': '',
            'Next': '',
            'With': '',
            'Each': 'of',
            'In': '',
            'New': 'new',
            'Class': 'class',
            'Private': 'const',
            'Public': 'let',
            'Static': 'const',
            'Optional': '',
            'ByVal': '',
            'ByRef': '',
            'Call': '',
            'Me': 'this',
            'Self': 'this',
            'Type': '',
            'Enum': '',
            'Implements': '',
            'Property': '',
            'Get': '',
            'Let': '',
            'Set': '',
            'Event': '',
            'Friend': '',
            'Protected': '',
            'MustInherit': '',
            'NotInheritable': '',
            'Overridable': '',
            'Overrides': '',
            'MustOverride': '',
            'NotOverridable': '',
            'Shadows': '',
            'Shared': 'static',
            'ReadOnly': '',
            'WriteOnly': '',
            'Default': '',
            'ParamArray': '',
            'Optional': '',
            'ParamArray': '',
            'Optional': ''
        }
        
        self.function_mappings = {
            'MsgBox': 'Application.MsgBox',
            'InputBox': 'Application.InputBox',
            'CreateObject': 'new ActiveXObject',
            'GetObject': 'GetObject',
            'IsNull': '=== null',
            'IsEmpty': '=== undefined',
            'IsNumeric': 'typeof === "number"',
            'IsDate': 'Object.prototype.toString.call() === "[object Date]"',
            'TypeName': 'typeof',
            'CStr': 'String',
            'CInt': 'parseInt',
            'CLng': 'parseInt',
            'CDbl': 'parseFloat',
            'CBool': 'Boolean',
            'CDate': 'new Date',
            'UCase': 'toUpperCase',
            'LCase': 'toLowerCase',
            'Left': 'substring(0,',
            'Right': 'substring(length -',
            'Mid': 'substring(',
            'Len': 'length',
            'Trim': 'trim',
            'LTrim': 'trimStart',
            'RTrim': 'trimEnd',
            'InStr': 'indexOf',
            'Replace': 'replace',
            'Split': 'split',
            'Join': 'join',
            'Date': 'new Date()',
            'Now': 'new Date()',
            'Time': 'new Date()',
            'Year': 'getFullYear()',
            'Month': 'getMonth() + 1',
            'Day': 'getDate()',
            'Hour': 'getHours()',
            'Minute': 'getMinutes()',
            'Second': 'getSeconds()',
            'Abs': 'Math.abs',
            'Sqr': 'Math.sqrt',
            'Sin': 'Math.sin',
            'Cos': 'Math.cos',
            'Tan': 'Math.tan',
            'Atn': 'Math.atan',
            'Exp': 'Math.exp',
            'Log': 'Math.log',
            'Int': 'Math.floor',
            'Fix': 'Math.trunc',
            'Rnd': 'Math.random',
            'Round': 'Math.round',
            'Asc': 'charCodeAt(0)',
            'Chr': 'String.fromCharCode',
            'Hex': 'toString(16)',
            'Oct': 'toString(8)',
            'RGB': '',
            'QBColor': '',
            'StrComp': 'localeCompare',
            'DateDiff': '',
            'DateAdd': '',
            'DatePart': '',
            'Format': '',
            'Shell': '',
            'Beep': '',
            'DoEvents': '',
            'Sleep': 'setTimeout(() => {}, ',
            'Dir': '',
            'FileLen': '',
            'FileDateTime': '',
            'GetAttr': '',
            'SetAttr': '',
            'Kill': '',
            'Name': '',
            'MkDir': '',
            'RmDir': '',
            'ChDir': '',
            'ChDrive': '',
            'CurDir': '',
            'Open': '',
            'Close': '',
            'Input': '',
            'Line': '',
            'Print': '',
            'Write': '',
            'Seek': '',
            'Loc': '',
            'LOF': '',
            'EOF': '',
            'FreeFile': '',
            'Reset': '',
            'Error': '',
            'Err': '',
            'Clear': '',
            'Raise': '',
            'On': '',
            'Resume': '',
            'GoTo': '',
            'Stop': '',
            'End': ''
        }
    
    def convert(self, vba_code):
        js_code = vba_code
        
        js_code = self._remove_comments(js_code)
        js_code = self._convert_sub_and_function(js_code)
        js_code = self._convert_variable_declarations(js_code)
        js_code = self._convert_control_flow(js_code)
        js_code = self._convert_functions(js_code)
        js_code = self._convert_objects(js_code)
        js_code = self._convert_comments(js_code)
        js_code = self._cleanup(js_code)
        
        return js_code
    
    def _remove_comments(self, code):
        code = re.sub(r"'[^'\n]*", '', code)
        code = re.sub(r'//[^\n]*', '', code)
        return code
    
    def _convert_sub_and_function(self, code):
        code = re.sub(r'Sub\s+(\w+)\s*\(([^)]*)\)', r'function \1(\2) {', code)
        code = re.sub(r'Function\s+(\w+)\s*\(([^)]*)\)', r'function \1(\2) {', code)
        code = re.sub(r'End\s+Sub\s*', '}', code)
        code = re.sub(r'End\s+Function\s*', '}', code)
        return code
    
    def _convert_variable_declarations(self, code):
        code = re.sub(r'Dim\s+(\w+)\s+As\s+\w+', r'let \1 = null;', code)
        code = re.sub(r'Dim\s+(\w+)', r'let \1 = null;', code)
        code = re.sub(r'Set\s+(\w+)\s*=\s*Nothing', r'\1 = null;', code)
        code = re.sub(r'Set\s+(\w+)\s*=\s*', r'\1 = ', code)
        return code
    
    def _convert_control_flow(self, code):
        code = re.sub(r'If\s+(.+?)\s+Then\s*', r'if (\1) {', code)
        code = re.sub(r'End\s+If\s*', '}', code)
        code = re.sub(r'ElseIf\s+(.+?)\s+Then\s*', r'} else if (\1) {', code)
        code = re.sub(r'Else\s*', '} else {', code)
        
        code = re.sub(r'For\s+(\w+)\s*=\s*(\d+)\s+To\s+(\d+)(?:\s+Step\s+(\d+))?', 
                      lambda m: f'for (let {m.group(1)} = {m.group(2)}; {m.group(1)} <= {m.group(3)}; {m.group(1)} += {m.group(4) or 1}) {{', 
                      code)
        code = re.sub(r'Next\s*(\w*)', '}', code)
        
        code = re.sub(r'Do\s+While\s+(.+)', r'do {', code)
        code = re.sub(r'Loop\s+While\s+(.+)', r'} while (\1);', code)
        code = re.sub(r'Loop\s*', '} while (true);', code)
        
        code = re.sub(r'While\s+(.+)', r'while (\1) {', code)
        code = re.sub(r'Wend\s*', '}', code)
        
        code = re.sub(r'Select\s+Case\s+(.+)', r'switch (\1) {', code)
        code = re.sub(r'Case\s+(.+)', r'case \1:', code)
        code = re.sub(r'Case\s+Else', 'default:', code)
        code = re.sub(r'End\s+Select\s*', '}', code)
        
        code = re.sub(r'Exit\s+Sub\s*', 'return;', code)
        code = re.sub(r'Exit\s+Function\s*', 'return;', code)
        
        return code
    
    def _convert_functions(self, code):
        for vba_func, js_func in self.function_mappings.items():
            pattern = r'\b' + re.escape(vba_func) + r'\s*\('
            code = re.sub(pattern, js_func + '(', code)
        return code
    
    def _convert_objects(self, code):
        code = re.sub(r'(\w+)\.(\w+)\s*=\s*New\s+(\w+)', r'\1.\2 = new \3();', code)
        code = re.sub(r'New\s+(\w+)\s*\(', r'new \1(', code)
        return code
    
    def _convert_comments(self, code):
        lines = code.split('\n')
        result = []
        in_multi_line = False
        
        for line in lines:
            stripped = line.strip()
            if stripped.startswith('/*'):
                in_multi_line = True
                result.append('//' + line[2:])
            elif stripped.endswith('*/'):
                in_multi_line = False
                result.append(line[:-2])
            elif in_multi_line:
                result.append('//' + line)
            else:
                result.append(line)
        
        return '\n'.join(result)
    
    def _cleanup(self, code):
        lines = code.split('\n')
        cleaned = []
        
        for line in lines:
            line = line.rstrip()
            line = re.sub(r'\s+', ' ', line)
            line = line.strip()
            if line:
                cleaned.append(line)
        
        return '\n'.join(cleaned)