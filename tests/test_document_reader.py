"""
文档读取MCP节点测试脚本
"""

import json
import sys
import os

def test_import():
    """测试模块导入"""
    print("=" * 60)
    print("测试1: 模块导入")
    print("=" * 60)
    
    try:
        from document_reader_mcp import (
            DocumentReader, DocumentReaderMCPNode, DocumentMetadata,
            mcp, FASTMCP_AVAILABLE, DOCX_AVAILABLE
        )
        print("[OK] 模块导入成功")
        print(f"  - FASTMCP_AVAILABLE: {FASTMCP_AVAILABLE}")
        print(f"  - DOCX_AVAILABLE: {DOCX_AVAILABLE}")
        return True
    except Exception as e:
        print(f"[FAIL] 模块导入失败: {e}")
        return False


def test_dataclass():
    """测试数据模型"""
    print("\n" + "=" * 60)
    print("测试2: 数据模型创建")
    print("=" * 60)
    
    try:
        from document_reader_mcp import DocumentMetadata
        
        metadata = DocumentMetadata(
            title="测试文档",
            author="测试作者",
            created_date="2024-01-01T00:00:00",
            modified_date="2024-01-02T00:00:00",
            file_type="docx",
            file_size=10240,
            paragraph_count=10,
            table_count=2
        )
        print(f"[OK] DocumentMetadata创建成功")
        print(f"  - 标题: {metadata.title}")
        print(f"  - 作者: {metadata.author}")
        return True
    except Exception as e:
        print(f"[FAIL] 数据模型创建失败: {e}")
        return False


def test_get_document_info():
    """测试获取文档信息功能"""
    print("\n" + "=" * 60)
    print("测试3: 获取文档信息")
    print("=" * 60)
    
    try:
        from document_reader_mcp import DocumentReaderMCPNode
        
        node = DocumentReaderMCPNode()
        
        # 检查是否有测试文档
        test_files = [
            "test_sample.docx",
            "sample_document.docx"
        ]
        
        test_file = None
        for f in test_files:
            if os.path.exists(f):
                test_file = f
                break
        
        if not test_file:
            print("[SKIP] 未找到测试文档，跳过此测试")
            return True
        
        result = node.reader.read_document(test_file)
        
        if result['success']:
            print(f"[OK] 文档信息获取成功")
            print(f"  - 文件: {test_file}")
            print(f"  - 标题: {result['document_metadata']['title']}")
            print(f"  - 段落数: {result['document_metadata']['paragraph_count']}")
            print(f"  - 表格数: {result['document_metadata']['table_count']}")
            print(f"  - 内容长度: {len(result['document_content'])} 字符")
            return True
        else:
            print(f"[FAIL] 文档信息获取失败: {result.get('error')}")
            return False
            
    except Exception as e:
        print(f"[FAIL] 获取文档信息失败: {e}")
        return False


def test_tool_definition():
    """测试工具定义获取"""
    print("\n" + "=" * 60)
    print("测试4: 工具定义获取")
    print("=" * 60)
    
    try:
        from document_reader_mcp import DocumentReaderMCPNode
        
        node = DocumentReaderMCPNode()
        tool_def = node.get_tool_definition()
        
        print(f"[OK] 工具定义获取成功")
        print(f"  - 名称: {tool_def['name']}")
        print(f"  - 版本: {tool_def['version']}")
        print(f"  - 工具数量: {len(tool_def['tools'])}")
        
        for tool in tool_def['tools']:
            print(f"    - {tool['name']}: {tool['description'][:30]}...")
        
        return True
    except Exception as e:
        print(f"[FAIL] 工具定义获取失败: {e}")
        return False


def test_fastmcp_tools():
    """测试FastMCP工具函数"""
    print("\n" + "=" * 60)
    print("测试5: FastMCP工具函数")
    print("=" * 60)
    
    try:
        from document_reader_mcp import FASTMCP_AVAILABLE, mcp
        
        if not FASTMCP_AVAILABLE:
            print("[SKIP] FastMCP未安装，跳过此测试")
            return True
        
        print("[OK] FastMCP可用")
        print(f"  - MCP名称: {mcp.name}")
        return True
    except Exception as e:
        print(f"[FAIL] FastMCP工具测试失败: {e}")
        return False


def main():
    print("=" * 60)
    print("文档读取MCP节点测试")
    print("=" * 60)
    
    tests = [
        ("模块导入", test_import),
        ("数据模型", test_dataclass),
        ("获取文档信息", test_get_document_info),
        ("工具定义获取", test_tool_definition),
        ("FastMCP工具函数", test_fastmcp_tools),
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
