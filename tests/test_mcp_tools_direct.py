"""
直接测试MCP工具函数（模拟Dify调用）
通过类方法调用，而非MCP工具装饰器
"""
import json
import os
import sys

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_read_document_tool():
    """测试read_document功能（通过类方法）"""
    print('\n' + '='*60)
    print('[测试] read_document功能')
    print('='*60)

    from unified_mcp_server import DocumentReader

    reader = DocumentReader()

    # 测试1: 正常读取
    result = reader.read_document('test_sample.docx')

    print(f'  调用成功: {result["success"]}')
    print(f'  段落数: {result["document_metadata"]["paragraph_count"]}')

    assert result['success'], "读取应成功"
    assert 'document_content' in result, "应包含document_content"
    assert 'document_metadata' in result, "应包含document_metadata"

    print('  [PASS] read_document功能测试通过')
    return True

def test_revise_document_tool():
    """测试revise_document功能（通过类方法）"""
    print('\n' + '='*60)
    print('[测试] revise_document功能')
    print('='*60)

    from unified_mcp_server import DocumentReviser

    reviser = DocumentReviser()

    # 测试1: 验证建议JSON
    suggestions = [
        {
            'id': 'ref_001',
            'rule_id': 'TEST_001',
            'original_text': '测试',
            'suggestion': '检查'
        }
    ]

    # 测试2: 修订文档
    output_path = 'test_mcp_tools_revised.docx'
    result = reviser.revise_document(
        'test_sample.docx',
        json.dumps(suggestions),
        output_path,
        'true'
    )

    print(f'  修订成功: {result["success"]}')
    print(f'  应用修订数: {result.get("applied_revisions", 0)}')
    print(f'  输出路径: {result.get("output_path", "N/A")}')

    assert result['success'], "修订应成功"
    assert os.path.exists(output_path), "输出文件应存在"

    # 清理
    if os.path.exists(output_path):
        os.remove(output_path)
        print('  已清理测试输出文件')

    print('  [PASS] revise_document功能测试通过')
    return True

def test_document_reader_mcp_tools():
    """测试document_reader_mcp的功能"""
    print('\n' + '='*60)
    print('[测试] document_reader_mcp功能')
    print('='*60)

    from document_reader_mcp import DocumentReaderMCPNode

    node = DocumentReaderMCPNode()

    # 测试1: 正常读取
    result = node.read_document('test_sample.docx')

    print(f'  read_document成功: {result["success"]}')
    print(f'  段落数: {result["document_metadata"]["paragraph_count"]}')

    # 测试2: 分段读取
    seg_result = node.read_document_segmented('test_sample.docx', 2000)

    print(f'  read_document_segmented成功: {seg_result["success"]}')
    if 'segment_info' in seg_result:
        print(f'  分段数: {seg_result["segment_info"]["total_segments"]}')

    print('  [PASS] document_reader_mcp功能测试通过')
    return True

def test_edge_cases():
    """测试边界情况"""
    print('\n' + '='*60)
    print('[测试] 边界情况')
    print('='*60)

    from unified_mcp_server import DocumentReader, DocumentReviser

    reader = DocumentReader()
    reviser = DocumentReviser()

    # 测试1: 空文件路径
    try:
        reader.read_document('')
        print('  [FAIL] 空路径应抛出异常')
        return False
    except (ValueError, FileNotFoundError) as e:
        print(f'  [OK] 空路径正确处理: {type(e).__name__}')

    # 测试2: 不存在的文件
    try:
        reader.read_document('不存在的文件.docx')
        print('  [FAIL] 不存在文件应抛出异常')
        return False
    except FileNotFoundError:
        print('  [OK] 不存在文件正确处理: FileNotFoundError')

    # 测试3: 无效JSON
    result = reviser.revise_document('test_sample.docx', 'invalid json')
    print(f'  [OK] 无效JSON处理: success={result["success"]}, error_type={result.get("error_type", "N/A")}')
    assert not result['success'], "无效JSON应失败"

    # 测试4: 空建议列表
    result = reviser.revise_document('test_sample.docx', '[]')
    print(f'  [OK] 空建议列表处理: success={result["success"]}, error_type={result.get("error_type", "N/A")}')

    print('  [PASS] 边界情况测试通过')
    return True

def test_fastmcp_server():
    """测试FastMCP服务器初始化"""
    print('\n' + '='*60)
    print('[测试] FastMCP服务器初始化')
    print('='*60)

    from unified_mcp_server import FASTMCP_AVAILABLE, DOCX_AVAILABLE, mcp

    print(f'  FastMCP可用: {FASTMCP_AVAILABLE}')
    print(f'  python-docx可用: {DOCX_AVAILABLE}')
    print(f'  MCP服务器名称: {mcp.name}')

    # 检查注册的工具
    if hasattr(mcp, '_tools'):
        tools = list(mcp._tools.keys())
        print(f'  注册工具数: {len(tools)}')
        for tool in tools:
            print(f'    - {tool}')

    assert FASTMCP_AVAILABLE, "FastMCP应可用"
    assert DOCX_AVAILABLE, "python-docx应可用"

    print('  [PASS] FastMCP服务器初始化测试通过')
    return True

def main():
    """主测试函数"""
    print('='*60)
    print('MCP功能直接调用测试')
    print('='*60)

    tests = [
        ('read_document功能', test_read_document_tool),
        ('revise_document功能', test_revise_document_tool),
        ('document_reader_mcp功能', test_document_reader_mcp_tools),
        ('边界情况', test_edge_cases),
        ('FastMCP服务器初始化', test_fastmcp_server),
    ]

    passed = 0
    failed = 0

    for name, test_func in tests:
        try:
            test_func()
            passed += 1
        except Exception as e:
            print(f'\n  [FAIL] {name} 测试失败: {e}')
            import traceback
            traceback.print_exc()
            failed += 1

    print('\n' + '='*60)
    print('测试汇总')
    print('='*60)
    print(f'  通过: {passed}/{len(tests)}')
    print(f'  失败: {failed}/{len(tests)}')

    if failed == 0:
        print('\n  [SUCCESS] 所有MCP功能测试通过!')
        return 0
    else:
        print('\n  [FAILED] 部分测试失败')
        return 1

if __name__ == '__main__':
    sys.exit(main())
