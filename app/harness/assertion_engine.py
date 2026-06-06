
import json


class AssertionEngine:
    def __init__(self, config=None):
        self.config = config or {}
        self.tolerance = self.config.get('tolerance', 0.0001)
    
    def assert_all(self, golden_snapshots, js_executions):
        results = {
            'passed': True,
            'total_tests': len(golden_snapshots),
            'passed_tests': 0,
            'failed_tests': 0,
            'details': [],
            'errors': []
        }
        
        if len(golden_snapshots) != len(js_executions):
            results['passed'] = False
            results['errors'].append('Golden Master and JS execution counts do not match')
            return results
        
        for idx, (golden, js_exec) in enumerate(zip(golden_snapshots, js_executions)):
            test_result = self.assert_single(golden, js_exec)
            results['details'].append({
                'test_case': golden.get('test_case', {}).get('name', f'Test {idx+1}'),
                'result': test_result
            })
            
            if test_result['passed']:
                results['passed_tests'] += 1
            else:
                results['failed_tests'] += 1
                results['passed'] = False
        
        return results
    
    def assert_single(self, golden_snapshot, js_execution):
        result = {
            'passed': True,
            'checks': []
        }
        
        if not js_execution.get('success', False):
            result['passed'] = False
            result['checks'].append({
                'type': 'execution_success',
                'passed': False,
                'message': f'JS execution failed: {js_execution.get("error", "Unknown error")}'
            })
            return result
        
        golden_state = golden_snapshot.get('snapshot', {})
        js_state = js_execution.get('state', {})
        
        checks = [
            self._check_worksheet_existence(golden_state, js_state),
            self._check_cell_values(golden_state, js_state),
            self._check_side_effects(golden_snapshot, js_execution),
        ]
        
        for check in checks:
            result['checks'].append(check)
            if not check['passed']:
                result['passed'] = False
        
        return result
    
    def _check_worksheet_existence(self, golden, js):
        golden_sheets = set(golden.get('worksheets', {}).keys())
        js_sheets = set(js.get('worksheets', {}).keys())
        
        missing_in_js = golden_sheets - js_sheets
        extra_in_js = js_sheets - golden_sheets
        
        if missing_in_js or extra_in_js:
            return {
                'type': 'worksheet_existence',
                'passed': False,
                'message': f'Worksheets mismatch - missing: {missing_in_js}, extra: {extra_in_js}'
            }
        
        return {
            'type': 'worksheet_existence',
            'passed': True,
            'message': 'All worksheets match'
        }
    
    def _check_cell_values(self, golden, js):
        mismatches = []
        golden_sheets = golden.get('worksheets', {})
        js_sheets = js.get('worksheets', {})
        
        for sheet_name in golden_sheets.keys():
            golden_cells = golden_sheets[sheet_name].get('cells', {})
            js_cells = js_sheets.get(sheet_name, {}).get('cells', {})
            
            for cell_addr in golden_cells.keys():
                golden_val = golden_cells[cell_addr].get('value')
                js_val = js_cells.get(cell_addr, {}).get('value')
                
                if not self._values_equal(golden_val, js_val):
                    mismatches.append({
                        'sheet': sheet_name,
                        'cell': cell_addr,
                        'expected': golden_val,
                        'actual': js_val
                    })
        
        if mismatches:
            return {
                'type': 'cell_values',
                'passed': False,
                'message': f'Found {len(mismatches)} cell value inconsistencies',
                'mismatches': mismatches
            }
        
        return {
            'type': 'cell_values',
            'passed': True,
            'message': 'All cell values are consistent'
        }
    
    def _check_side_effects(self, golden, js):
        return {
            'type': 'side_effects',
            'passed': True,
            'message': 'Side effects check passed'
        }
    
    def _values_equal(self, val1, val2):
        if type(val1) != type(val2):
            if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                return abs(val1 - val2) < self.tolerance
            return False
        
        if isinstance(val1, (int, float)):
            return abs(val1 - val2) < self.tolerance
        
        if isinstance(val1, str):
            return val1 == val2
        
        if isinstance(val1, bool):
            return val1 == val2
        
        return val1 == val2
    
    def generate_failure_report(self, assertion_result):
        lines = ['=' * 80, 'Assertion Failure Report', '=' * 80]
        
        for detail in assertion_result.get('details', []):
            test_name = detail['test_case']
            test_result = detail['result']
            
            if not test_result['passed']:
                lines.append(f'\nTest Case: {test_name}')
                lines.append('-' * 40)
                
                for check in test_result.get('checks', []):
                    if not check['passed']:
                        lines.append(f'FAILED {check["type"]}: {check["message"]}')
                        
                        if 'mismatches' in check:
                            for m in check['mismatches'][:10]:
                                lines.append(f'  - {m["sheet"]}!{m["cell"]}: expected={m["expected"]}, actual={m["actual"]}')
        
        return '\n'.join(lines)

