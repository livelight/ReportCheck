"""
项目整体测试脚本
运行所有关键测试
"""

import sys
import os

def run_test(name, func):
    """运行单个测试"""
    print(f"\n{'='*60}")
    print(f"测试: {name}")
    print('='*60)
    try:
        func()
        print(f"[PASS] {name}")
        return True
    except Exception as e:
        print(f"[FAIL] {name}: {e}")
        return False

def test_import():
    """测试模块导入"""
    from revision_mcp_node import (
        RevisionMCPNode, CheckSuggestion, DocumentParser,
        TrackChangesReviser, mcp, FASTMCP_AVAILABLE,
        transform_step4_to_step6, prepare_suggestions_for_revision
    )
    assert FASTMCP_AVAILABLE, "FastMCP不可用"
    node = RevisionMCPNode()
    assert node.name == "DocumentRevisionNode"

def test_data_model():
    """测试数据模型"""
    from revision_mcp_node import CheckSuggestion
    suggestion = CheckSuggestion(
        id='test-001',
        rule_id='RULE-001',
        rule_name='测试规则',
        type='hard',
        severity='high',
        section='标题',
        original_text='原文本',
        suggestion='建议文本',
        reason='测试原因'
    )
    assert suggestion.id == 'test-001'
    assert suggestion.severity == 'high'

def test_tool_definition():
    """测试工具定义"""
    from revision_mcp_node import RevisionMCPNode
    node = RevisionMCPNode()
    tool_def = node.get_tool_definition()
    print(f"Tool def keys: {list(tool_def.keys())}")
    print(f"Tools count: {len(tool_def.get('tools', []))}")
    assert 'name' in tool_def
    assert 'tools' in tool_def
    # 工具数量可能为2或4，取决于实现方式
    assert len(tool_def['tools']) >= 2

def test_transform():
    """测试数据转换"""
    import json
    from revision_mcp_node import transform_step4_to_step6
    
    step4_output = {
        'suggestions': [
            {
                'id': 'ref_001',
                'rule_id': 'RULE_001',
                'rule_name': '命名规范',
                'type': 'hard',
                'severity': 'High',
                'section': '标题',
                'original_text': 'test TEXT',
                'suggested_text': 'test text',
                'reason': '命名应小写',
                'accepted': True
            }
        ]
    }
    
    result = transform_step4_to_step6(json.dumps(step4_output))
    result_obj = json.loads(result)
    # transform_step4_to_step6 返回的是建议数组
    assert isinstance(result_obj, list)
    assert len(result_obj) == 1
    assert result_obj[0]['id'] == 'ref_001'

def test_sse_server_import():
    """测试SSE服务器导入"""
    from mcp_sse_server import app, clients
    assert app is not None

def test_files_exist():
    """测试关键文件存在"""
    required_files = [
        'revision_mcp_node.py',
        'mcp_sse_server.py',
        'mcp_server.py',
        'requirements.txt',
        'README.md',
        'prompt_step2_hard_rules.md',
        'prompt_step3_soft_rules_v2.md',
        'prompt_step4_merge_prioritize.md',
        'prompt_step5_user_confirmation.md',
    ]
    
    for f in required_files:
        assert os.path.exists(f), f"缺少文件: {f}"

def main():
    print("="*60)
    print("项目整体测试")
    print("="*60)
    
    tests = [
        ("模块导入", test_import),
        ("数据模型", test_data_model),
        ("工具定义", test_tool_definition),
        ("数据转换", test_transform),
        ("SSE服务器导入", test_sse_server_import),
        ("文件完整性", test_files_exist),
    ]
    
    results = []
    for name, func in tests:
        results.append((name, run_test(name, func)))
    
    # 汇总
    print("\n" + "="*60)
    print("测试结果汇总")
    print("="*60)
    
    passed = sum(1 for _, r in results if r)
    total = len(results)
    
    for name, result in results:
        status = "通过" if result else "失败"
        print(f"  [{status}] {name}")
    
    print("-"*60)
    print(f"总计: {passed}/{total} 通过")
    print("="*60)
    
    return 0 if passed == total else 1

if __name__ == "__main__":
    sys.exit(main())
