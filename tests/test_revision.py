"""
测试脚本 - 文档修订MCP节点测试
"""

import json
import os
import sys
import tempfile
from pathlib import Path

# 添加当前目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from revision_mcp_node import RevisionMCPNode, CheckSuggestion


def create_sample_docx(output_path: str):
    """创建一个示例Word文档用于测试"""
    from docx import Document
    from docx.shared import Pt, RGBColor

    doc = Document()

    # 标题
    title = doc.add_heading('项目进度报告', level=1)

    # 章节1
    doc.add_heading('一、项目概述', level=2)
    doc.add_paragraph('本项目旨在开发一套企业级文档管理系统。')

    # 章节2
    doc.add_heading('二、当前进展', level=2)
    doc.add_paragraph('目前已完成核心功能的开发和测试工作。')

    # 章节3
    doc.add_heading('三、风险分析', level=2)
    doc.add_paragraph('项目进展顺利，暂无风险。')

    # 章节4
    doc.add_heading('四、下一步计划', level=2)
    doc.add_paragraph('继续优化系统性能，并进行用户验收测试。')

    # 章节5
    doc.add_heading('五、总结', level=2)
    doc.add_paragraph('整体项目进度符合预期。')

    doc.save(output_path)
    print(f"创建测试文档: {output_path}")
    return output_path


def create_sample_suggestions() -> str:
    """创建示例检查建议JSON"""
    suggestions = [
        {
            "id": "ref_001",
            "rule_id": "FORMAT_001",
            "rule_name": "标题格式不规范",
            "type": "format",
            "severity": "High",
            "section": "sec_001",
            "original_text": "一、项目概述",
            "suggestion": "一、项目概述",
            "reason": "标题应简洁，不包含多余标点"
        },
        {
            "id": "ref_002",
            "rule_id": "CONTENT_023",
            "rule_name": "缺少风险分析",
            "type": "content",
            "severity": "High",
            "section": "sec_003",
            "original_text": "项目进展顺利，暂无风险。",
            "suggestion": "项目整体进展顺利，目前主要风险包括：1）人员流动性风险；2）技术架构调整风险。建议加强团队稳定性和代码文档建设。",
            "reason": "风险分析过于简单，应包含具体风险点和应对措施"
        },
        {
            "id": "ref_003",
            "rule_id": "LANG_001",
            "rule_name": "语法错误",
            "type": "language",
            "severity": "Medium",
            "section": "sec_002",
            "original_text": "目前已完成核心功能的开发和测试工作。",
            "suggestion": "目前已完成核心功能的开发和测试工作，整体质量符合预期。",
            "reason": "语句缺少主谓完整性的描述"
        },
        {
            "id": "ref_004",
            "rule_id": "LOGIC_005",
            "rule_name": "逻辑顺序问题",
            "type": "logic",
            "severity": "Low",
            "section": "sec_004",
            "original_text": "继续优化系统性能，并进行用户验收测试。",
            "suggestion": "1. 首先进行用户验收测试；2. 根据测试反馈优化系统性能；3. 部署上线。",
            "reason": "建议先测试后优化，逻辑更合理"
        }
    ]
    return json.dumps(suggestions, ensure_ascii=False, indent=2)


def test_revision_mcp_node():
    """测试文档修订MCP节点"""
    print("=" * 60)
    print("文档修订MCP节点测试")
    print("=" * 60)

    # 创建MCP节点实例
    node = RevisionMCPNode()

    # 打印工具定义
    print("\n[工具定义]")
    print(json.dumps(node.get_tool_definition(), indent=2, ensure_ascii=False))

    # 创建临时测试文档
    temp_dir = tempfile.mkdtemp()
    doc_path = os.path.join(temp_dir, "test_document.docx")
    create_sample_docx(doc_path)

    # 获取示例建议
    suggestions_json = create_sample_suggestions()

    print("\n[测试输入]")
    print(f"文档路径: {doc_path}")
    print(f"文档存在: {os.path.exists(doc_path)}")
    print(f"建议数量: 4条")

    # 执行修订
    print("\n[执行修订...]")
    result = node.revise_document(
        file_path=doc_path,
        suggestions_json=suggestions_json,
        preserve_original=True
    )

    # 输出结果
    print("\n[修订结果]")
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # 验证输出
    if result.get('success'):
        output_path = result.get('output_path', '')
        if os.path.exists(output_path):
            print(f"\n[验证通过] 修订后文档已生成: {output_path}")
            print(f"文件大小: {os.path.getsize(output_path)} bytes")

            # 读取修订后的文档内容
            from docx import Document
            revised_doc = Document(output_path)
            print("\n[修订后文档内容预览]")
            for i, para in enumerate(revised_doc.paragraphs[:10]):
                if para.text.strip():
                    print(f"  {para.text[:80]}")
        else:
            print(f"\n[警告] 输出文件不存在: {output_path}")
    else:
        print(f"\n[错误] 修订失败: {result.get('error')}")

    # 清理
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

    return result


def test_json_parsing():
    """测试JSON解析"""
    print("\n[测试JSON解析]")

    test_cases = [
        # 正常JSON
        '[{"id":"001","rule_id":"R001","rule_name":"测试规则","type":"format","severity":"High","section":"sec1","original_text":"原文","suggestion":"建议","reason":"原因"}]',
        # 空数组
        '[]',
        # 无效JSON
        '{invalid json}'
    ]

    node = RevisionMCPNode()

    for i, test_json in enumerate(test_cases):
        print(f"\n测试用例 {i+1}:")
        print(f"  输入: {test_json[:50]}...")
        try:
            suggestions = json.loads(test_json)
            print(f"  结果: 解析成功，共{len(suggestions)}条建议")
        except json.JSONDecodeError as e:
            print(f"  结果: JSON解析错误 - {e}")


def test_suggestion_model():
    """测试建议数据模型"""
    print("\n[测试建议数据模型]")

    test_data = {
        "id": "ref_001",
        "rule_id": "CONTENT_023",
        "rule_name": "缺少风险分析",
        "type": "content",
        "severity": "High",
        "section": "sec_005",
        "original_text": "项目进展顺利，暂无风险。",
        "suggestion": "项目整体进展顺利，目前主要风险包括：1）人员流动性风险；2）技术架构调整风险。",
        "reason": "风险分析过于简单"
    }

    suggestion = CheckSuggestion.from_dict(test_data)

    print(f"  ID: {suggestion.id}")
    print(f"  规则名称: {suggestion.rule_name}")
    print(f"  类型: {suggestion.type}")
    print(f"  严重程度: {suggestion.severity}")
    print(f"  原始文本: {suggestion.original_text[:30]}...")
    print(f"  修订建议: {suggestion.suggestion[:30]}...")

    print("\n[测试通过]")


if __name__ == "__main__":
    # 运行所有测试
    test_suggestion_model()
    test_json_parsing()
    result = test_revision_mcp_node()
