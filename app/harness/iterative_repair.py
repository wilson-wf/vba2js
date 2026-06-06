
import time
from .golden_master import GoldenMasterCapturer
from .wps_simulator import WPSSimulator
from .assertion_engine import AssertionEngine


class IterativeRepairEngine:
    def __init__(self, llm_agent, config=None):
        self.llm_agent = llm_agent
        self.config = config or {}
        self.max_iterations = self.config.get('max_iterations', 10)
        self.patience = self.config.get('patience', 3)
        
        self.golden_capturer = GoldenMasterCapturer(config)
        self.simulator = WPSSimulator()
        self.assertion_engine = AssertionEngine(config)
        
        self.history = []
    
    def run_iterative_repair(
        self,
        xlsm_path,
        initial_js_code,
        test_cases,
        macro_name
    ):
        print(f"  Starting iterative repair process, max {self.max_iterations} rounds")
        print(f"  Test cases count: {len(test_cases)}")
        print(f"  Macro name: {macro_name}")
        
        print("\n  Step 1: Capturing Golden Master...")
        golden_snapshots = self.golden_capturer.capture(xlsm_path, test_cases)
        print(f"  Successfully captured {len(golden_snapshots)} Golden Master snapshots")
        
        current_js = initial_js_code
        best_js = initial_js_code
        best_score = 0
        consecutive_failures = 0
        
        for iteration in range(1, self.max_iterations + 1):
            print(f"\n{'='*80}")
            print(f"  Round {iteration}")
            print(f"{'='*80}")
            
            test_result = self._test_current_code(
                current_js, golden_snapshots, test_cases, macro_name, xlsm_path
            )
            
            self.history.append({
                'iteration': iteration,
                'js_code': current_js,
                'test_result': test_result
            })
            
            passed = test_result['assertion']['passed']
            score = test_result['assertion']['passed_tests'] / max(1, test_result['assertion']['total_tests'])
            
            print(f"\n  This round result: {'PASSED' if passed else 'FAILED'}")
            print(f"  Passed tests: {test_result['assertion']['passed_tests']}/{test_result['assertion']['total_tests']}")
            print(f"  Score: {score:.2%}")
            
            if score > best_score:
                best_score = score
                best_js = current_js
                consecutive_failures = 0
                print(f"  New best score!")
            else:
                consecutive_failures += 1
                print(f"  {consecutive_failures} consecutive rounds without improvement")
            
            if passed:
                print(f"\n  Congratulations! All tests passed, iterative repair complete!")
                return {
                    'success': True,
                    'final_js_code': current_js,
                    'iterations': iteration,
                    'history': self.history,
                    'final_assertion': test_result['assertion']
                }
            
            if consecutive_failures >= self.patience:
                print(f"\n  {self.patience} consecutive rounds without improvement, stopping early, using best result")
                return {
                    'success': False,
                    'final_js_code': best_js,
                    'iterations': iteration,
                    'history': self.history,
                    'final_assertion': test_result['assertion'],
                    'reason': 'patience_exceeded'
                }
            
            print(f"\n  Calling LLM to repair code...")
            repair_result = self._request_repair(
                current_js, test_result, iteration
            )
            
            if not repair_result['success']:
                print(f"  Repair request failed, keeping current code")
                consecutive_failures += 1
                continue
            
            current_js = repair_result['repaired_code']
            print(f"  Got repaired code")
        
        print(f"\n  Reached max iterations {self.max_iterations}, using best result")
        return {
            'success': False,
            'final_js_code': best_js,
            'iterations': self.max_iterations,
            'history': self.history,
            'final_assertion': self.history[-1]['test_result']['assertion'],
            'reason': 'max_iterations'
        }
    
    def _test_current_code(
        self,
        js_code,
        golden_snapshots,
        test_cases,
        macro_name,
        xlsm_path
    ):
        js_executions = []
        
        for idx, test_case in enumerate(test_cases):
            print(f"\n  Executing test {idx+1}/{len(test_cases)}: {test_case.get('name', 'unnamed')}")
            
            self.simulator.load_workbook(xlsm_path)
            
            golden = golden_snapshots[idx]
            initial_state = golden.get('snapshot', {}).copy()
            
            setup_data = test_case.get('setup', {})
            for sheet_name, cells in setup_data.items():
                if sheet_name in initial_state.get('worksheets', {}):
                    for cell_addr, value in cells.items():
                        if 'cells' not in initial_state['worksheets'][sheet_name]:
                            initial_state['worksheets'][sheet_name]['cells'] = {}
                        initial_state['worksheets'][sheet_name]['cells'][cell_addr] = {'value': value}
            
            self.simulator.apply_initial_state(initial_state)
            
            execution_result = self.simulator.execute_js_macro(
                js_code, macro_name
            )
            js_executions.append(execution_result)
            
            status = 'PASSED' if execution_result.get('success') else 'FAILED'
            print(f"    {status} execution result: {execution_result.get('result')}")
            if not execution_result.get('success'):
                print(f"    Error: {execution_result.get('error')}")
        
        assertion_result = self.assertion_engine.assert_all(golden_snapshots, js_executions)
        
        return {
            'js_executions': js_executions,
            'assertion': assertion_result
        }
    
    def _request_repair(self, current_js, test_result, iteration):
        try:
            failure_report = self.assertion_engine.generate_failure_report(test_result['assertion'])
            
            repaired_code = self.llm_agent.repair_js_code(
                current_js,
                failure_report,
                iteration
            )
            
            return {
                'success': True,
                'repaired_code': repaired_code
            }
        except Exception as e:
            print(f"  Repair request exception: {e}")
            return {
                'success': False,
                'error': str(e)
            }

