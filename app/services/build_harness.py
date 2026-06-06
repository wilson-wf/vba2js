
import os
import uuid
from ..engine.vba_parser import VBAParser
from ..engine.syntax_mapper import SyntaxMapper
from ..engine.llm_agent import LLMAgent
from ..engine.excel_generator import ExcelGenerator
from ..harness import IterativeRepairEngine


class BuildHarnessService:
    def __init__(self, config=None):
        self.config = config or {}
        
        self.vba_parser = VBAParser()
        self.syntax_mapper = SyntaxMapper()
        self.llm_agent = LLMAgent(
            api_key=self.config.get('llm_api_key'),
            endpoint=self.config.get('llm_endpoint'),
            model=self.config.get('llm_model')
        )
        self.excel_generator = ExcelGenerator()
        self.repair_engine = IterativeRepairEngine(
            self.llm_agent,
            config={
                'max_iterations': self.config.get('max_iterations', 5),
                'patience': self.config.get('patience', 2)
            }
        )
        
        self.temp_dir = 'data/temp'
        os.makedirs(self.temp_dir, exist_ok=True)
    
    def convert_with_harness(
        self,
        xlsm_path,
        test_cases=None,
        use_llm=True,
        use_iterative=True
    ):
        print("  Starting Build Harness conversion process")
        print(f"  File: {xlsm_path}")
        print(f"  Using LLM: {use_llm}")
        print(f"  Using iterative repair: {use_iterative}")
        
        result = {
            'success': False,
            'input_file': xlsm_path,
            'output_file': None,
            'modules': [],
            'harness_result': None,
            'warnings': []
        }
        
        try:
            print("\n  Step 1: Parsing VBA code...")
            vba_modules = self.vba_parser.extract_vba_modules(xlsm_path)
            print(f"  Found {len(vba_modules)} VBA modules")
            
            if not vba_modules:
                result['warnings'].append('No VBA modules found')
                return result
            
            if not test_cases:
                print("\n  Step 2: Generating test cases...")
                test_cases = self._generate_default_test_cases(vba_modules)
                print(f"  Generated {len(test_cases)} test cases")
            
            converted_modules = []
            
            for module in vba_modules:
                print(f"\n  Processing module: {module['name']}")
                print(f"  Code lines: {len(module['code'].splitlines())}")
                
                initial_js = self._initial_convert(module['code'], use_llm)
                
                if use_iterative:
                    print("\n  Starting iterative repair...")
                    repair_result = self.repair_engine.run_iterative_repair(
                        xlsm_path=xlsm_path,
                        initial_js_code=initial_js,
                        test_cases=test_cases,
                        macro_name=self._extract_macro_name(module['code'])
                    )
                    
                    converted_modules.append({
                        'name': module['name'],
                        'original_vba': module['code'],
                        'initial_js': initial_js,
                        'final_js': repair_result['final_js_code'],
                        'harness_result': repair_result
                    })
                    
                    if repair_result['success']:
                        print("  Iterative repair successful, all tests passed!")
                    else:
                        print(f"  Using best result, test pass rate: {repair_result['final_assertion']['passed_tests']}/{repair_result['final_assertion']['total_tests']}")
                else:
                    converted_modules.append({
                        'name': module['name'],
                        'original_vba': module['code'],
                        'initial_js': initial_js,
                        'final_js': initial_js,
                        'harness_result': None
                    })
            
            print("\n  Step 4: Generating final Excel file...")
            output_path = os.path.join(self.temp_dir, f"{uuid.uuid4().hex}.xlsm")
            
            js_codes = {m['name']: m['final_js'] for m in converted_modules}
            self.excel_generator.generate_excel_with_js(xlsm_path, js_codes, output_path)
            
            result['success'] = True
            result['output_file'] = output_path
            result['modules'] = converted_modules
            
            print("\n  Build Harness process complete!")
            print(f"  Output file: {output_path}")
            
            return result
            
        except Exception as e:
            print(f"\n  Build Harness process failed: {e}")
            result['error'] = str(e)
            return result
    
    def _initial_convert(self, vba_code, use_llm):
        if use_llm:
            llm_result = self.llm_agent.convert_vba_to_js(vba_code)
            if llm_result:
                return llm_result
        
        return self.syntax_mapper.convert(vba_code)
    
    def _extract_macro_name(self, vba_code):
        lines = vba_code.split('\n')
        for line in lines:
            if 'Sub ' in line and '()' in line:
                parts = line.strip().split()
                if len(parts) >= 2:
                    name = parts[1].split('(')[0]
                    return name
        return 'Main'
    
    def _generate_default_test_cases(self, vba_modules):
        return [
            {
                'name': 'Basic Test',
                'setup': {
                    'Sheet1': {
                        'A1': 10,
                        'B1': 20
                    }
                }
            }
        ]

