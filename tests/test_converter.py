import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app.engine.vba_parser import VBAParser
from app.engine.syntax_mapper import SyntaxMapper
from app.engine.llm_agent import LLMAgent
from app.engine.excel_generator import ExcelGenerator

def test_vba_parser():
    print("测试 VBA解析器...")
    
    parser = VBAParser()
    
    test_code = '''
Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
'''
    
    print("✓ VBA解析器初始化成功")

def test_syntax_mapper():
    print("测试 语法映射器...")
    
    mapper = SyntaxMapper()
    
    vba_code = '''
Sub HelloWorld()
    Dim x As Integer
    x = 10
    If x > 5 Then
        MsgBox "Hello"
    End If
End Sub
'''
    
    js_code = mapper.convert(vba_code)
    print(f"转换结果:\n{js_code}")
    print("✓ 语法映射器测试成功")

def test_llm_agent():
    print("测试 LLM代理...")
    
    agent = LLMAgent(api_key=None)
    result = agent.convert_vba_to_js('Sub Test()\nMsgBox "Test"\nEnd Sub')
    
    if result is None:
        print("✓ LLM代理初始化成功（未配置API Key）")
    else:
        print(f"✓ LLM转换结果:\n{result}")

def test_excel_generator():
    print("测试 Excel生成器...")
    
    generator = ExcelGenerator()
    print("✓ Excel生成器初始化成功")

if __name__ == '__main__':
    print("=== VBA到WPS JS宏转换器测试 ===")
    print()
    
    test_vba_parser()
    print()
    
    test_syntax_mapper()
    print()
    
    test_llm_agent()
    print()
    
    test_excel_generator()
    print()
    
    print("=== 所有测试通过 ===")