"""
统一MCP服务器测试脚本
"""

import json
import os
import sys

def test_import():
    """测试模块导入"""
    print("=" * 60)
    print("测试1: 模块导入")
    print("=" * 60)
    
    try:
        from unified_mcp_server import (
            DocumentReader, DocumentReviser, mcp, 
            FASTMCP_AVAILABLE, DOCX_AVAILABLE
        )
        print("[OK] unified_mcp_server 导入成功")
        print(f"  - FASTMCP_AVAILABLE: {FASTMCP_AVAILABLE}")
        print(f"  - DOCX_AVAILABLE: {DOCX_AVAILABLE}")
        print(f"  - MCP名称: {mcp.name}")
        return True
    except Exception as e:
        print(f"[FAIL] 模块导入失败: {e}")
        return False

def test_read_document():
    """测试文档读取功能"""
    print("\n" + "=" * 60)
    print("测试2: 文档读取功能")
    print("=" * 60)
    
    try:
        from unified_mcp_server import DocumentReader
        
        # 查找测试文件
        test_files = ["test_sample.docx", "sample_document.docx"]
        test_file = None
        for f in test_files:
            if os.path.exists(f):
                test_file = f
                break
        
        if not test_file:
            print("[SKIP] 未找到测试文档，跳过")
            return True
        
        reader = DocumentReader()
        result = reader.read_document(test_file)
        
        if result['success']:
            print(f"[OK] 文档读取成功")
            print(f"  - 文件: {test_file}")
            print(f"  - 段落数: {result['document_metadata']['paragraph_count']}")
            print(f"  - 表格数: {result['document_metadata']['table_count']}")
            print(f"  - 内容长度: {len(result['document_content'])} 字符")
            return True
        else:
            print(f"[FAIL] 文档读取失败: {result.get('error')}")
            return False
    except Exception as e:
        print(f"[FAIL] 文档读取异常: {e}")
        return False

def test_revise_document():
    """测试文档修订功能"""
    print("\n" + "=" * 60)
    print("测试3: 文档修订功能")
    print("=" * 60)
    
    try:
        from unified_mcp_server import DocumentReviser
        
        # 查找测试文件
        test_files = ["test_sample.docx", "sample_document.docx"]
        test_file = None
        for f in test_files:
            if os.path.exists(f):
                test_file = f
                break
        
        if not test_file:
            print("[SKIP] 未找到测试文档，跳过")
            return True
        
        # 模拟检查建议
        suggestions = [
            {
                "id": "test_001",
                "rule_id": "TEST_001",
                "rule_name": "测试修订",
                "type": "content",
                "severity": "High",
                "section": "测试",
                "original_text": "测试",
                "suggestion": "测试文本",
                "reason": "测试修订功能"
            }
        ]
        
        reviser = DocumentReviser()
        output_path = test_file.replace('.docx', '_test_revised.docx')
        
        result = reviser.revise_document(
            file_path=test_file,
            suggestions_json=json.dumps(suggestions),
            output_path=output_path,
            use_track_changes="false"
        )
        
        if result['success']:
            print(f"[OK] 文档修订成功")
            print(f"  - 输出文件: {result['output_path']}")
            print(f"  - 应用修订: {result['applied_revisions']}/{result['total_suggestions']}")
            
            # 清理测试文件
            if os.path.exists(output_path):
                os.remove(output_path)
            
            return True
        else:
            print(f"[FAIL] 文档修订失败: {result.get('error')}")
            return False
    except Exception as e:
        print(f"[FAIL] 文档修订异常: {e}")
        return False

def test_mcp_tools():
    """测试MCP工具"""
    print("\n" + "=" * 60)
    print("测试4: MCP工具")
    print("=" * 60)
    
    try:
        from unified_mcp_server import FASTMCP_AVAILABLE, mcp
        
        if not FASTMCP_AVAILABLE:
            print("[SKIP] FastMCP未安装，跳过")
            return True
        
        print(f"[OK] FastMCP可用")
        print(f"  - MCP名称: {mcp.name}")
        
        # 测试工具列表
        tools = ['read_document', 'get_document_info', 'revise_document', 'validate_suggestions', 'get_tools_info']
        print(f"  - 注册工具数: {len(tools)}")
        for tool in tools:
            print(f"    - {tool}")
        
        return True
    except Exception as e:
        print(f"[FAIL] MCP工具测试失败: {e}")
        return False

def main():
    print("=" * 60)
    print("统一MCP服务器测试")
    print("=" * 60)
    
    tests = [
        ("模块导入", test_import),
        ("文档读取功能", test_read_document),
        ("文档修订功能", test_revise_document),
        ("MCP工具", test_mcp_tools),
    ]
    
    results = []
    for name, func in tests:
        try:
            passed = func()
            results.append((name, passed))
        except Exception as e:
            print(f"\n[ERROR] 测试异常: {e}")
            results.append((name, False))
    
    # 汇总
    print("\n" + "=" * 60)
    print("测试结果汇总")
    print("=" * 60)
    
    passed = sum(1 for _, r in results if r)
    total = len(results)
    
    for name, result in results:
        status = "通过" if result else "失败"
        print(f"  [{status}] {name}")
    
    print("-" * 60)
    print(f"总计: {passed}/{total} 通过")
    print("=" * 60)
    
    return 0 if passed == total else 1

if __name__ == "__main__":
    sys.exit(main())
