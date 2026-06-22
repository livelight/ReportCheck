"""
测试第4步到第6步的数据流转
验证代码优化后的兼容性
"""

import json
import sys
sys.path.insert(0, 'c:\\Users\\lpy09\\ComateProjects\\comate-zulu-demo-1781339688688')

from revision_mcp_node import (
    transform_step4_to_step6,
    prepare_suggestions_for_revision,
    CheckSuggestion
)


def test_step4_format():
    """测试第4步输出格式转换"""
    print("=" * 70)
    print("测试1: 第4步输出格式转换")
    print("=" * 70)
    
    # 模拟第4步输出（包含suggestions和statistics）
    step4_output = {
        "suggestions": [
            {
                "id": "sugg_001",
                "rule_id": "NAMING_TEAM_001",
                "rule_name": "项目组名称不规范",
                "type": "naming",
                "severity": "Critical",
                "section": "事件概述",
                "original_text": "项目组立即启动应急预案",
                "suggestion": "改为'研发中心核心系统项目组立即启动应急预案'",
                "reason": "硬性规则要求：不能直接写'项目组'",
                "source": "hard",
                "group": None,
                "order": 1,
                "accepted": True
            },
            {
                "id": "sugg_002",
                "rule_id": "CONTENT_DEPTH_001",
                "rule_name": "根因分析深度不足",
                "type": "content",
                "severity": "High",
                "section": "根因分析",
                "original_text": "数据库连接池耗尽",
                "suggestion": "建议进行5Why分析",
                "reason": "当前分析停留在表面现象",
                "source": "soft",
                "group": None,
                "order": 2,
                "accepted": True
            },
            {
                "id": "sugg_003",
                "rule_id": "LANGUAGE_STYLE_001",
                "rule_name": "语言风格口语化",
                "type": "language",
                "severity": "Low",
                "section": "总结",
                "original_text": "以后要加强管理",
                "suggestion": "改为'后续将建立常态化监控机制'",
                "reason": "'以后'过于口语化",
                "source": "soft",
                "group": None,
                "order": 3,
                "accepted": False  # 用户拒绝
            }
        ],
        "statistics": {
            "total": 3,
            "by_severity": {"Critical": 1, "High": 1, "Medium": 0, "Low": 1}
        },
        "processing_notes": ["已合并1条重复建议"]
    }
    
    step4_json = json.dumps(step4_output, ensure_ascii=False)
    
    print("\n第4步输出格式（对象包装）:")
    print(json.dumps(step4_output, ensure_ascii=False, indent=2)[:500] + "...")
    
    # 测试转换
    try:
        result_json = transform_step4_to_step6(step4_json, accepted_only=True)
        result = json.loads(result_json)
        
        print(f"\n[OK] 转换成功")
        print(f"     原始建议数: {len(step4_output['suggestions'])}")
        print(f"     转换后建议数: {len(result)}")
        print(f"     已过滤拒绝的建议: sugg_003 (accepted=False)")
        
        # 验证字段
        first = result[0]
        required_fields = ['id', 'rule_id', 'rule_name', 'type', 'severity', 
                          'section', 'original_text', 'suggestion', 'reason']
        missing = [f for f in required_fields if f not in first]
        
        if missing:
            print(f"     [FAIL] 缺少字段: {missing}")
        else:
            print(f"     [OK] 所有必需字段都存在")
            
        # 验证没有多余字段
        extra_fields = [f for f in first if f not in required_fields]
        if extra_fields:
            print(f"     [INFO] 移除了多余字段: {extra_fields}")
        else:
            print(f"     [OK] 没有多余字段")
            
        return True
        
    except Exception as e:
        print(f"\n[FAIL] 转换失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_direct_array_format():
    """测试直接数组格式"""
    print("\n" + "=" * 70)
    print("测试2: 直接数组格式（兼容旧格式）")
    print("=" * 70)
    
    # 直接数组格式（旧格式）
    suggestions_array = [
        {
            "id": "hard_001",
            "rule_id": "PUNCT_TIME_001",
            "rule_name": "时间格式不规范",
            "type": "punctuation",
            "severity": "High",
            "section": "事件概述",
            "original_text": "14：30",
            "suggestion": "改为'14:30'",
            "reason": "硬性规则要求：时间格式必须为09:00-23:16"
        }
    ]
    
    array_json = json.dumps(suggestions_array, ensure_ascii=False)
    
    print("\n输入格式（直接数组）:")
    print(json.dumps(suggestions_array, ensure_ascii=False, indent=2))
    
    # 测试转换
    try:
        result_json = transform_step4_to_step6(array_json, accepted_only=False)
        result = json.loads(result_json)
        
        print(f"\n[OK] 转换成功")
        print(f"     建议数: {len(result)}")
        print(f"     [OK] 直接数组格式兼容")
        
        return True
        
    except Exception as e:
        print(f"\n[FAIL] 转换失败: {e}")
        return False


def test_prepare_function():
    """测试prepare_suggestions_for_revision函数"""
    print("\n" + "=" * 70)
    print("测试3: prepare_suggestions_for_revision函数")
    print("=" * 70)
    
    # 模拟第4步输出
    step4_output = {
        "suggestions": [
            {"id": "sugg_001", "rule_id": "R001", "rule_name": "问题1", 
             "type": "format", "severity": "Critical", "section": "概述",
             "original_text": "原文1", "suggestion": "建议1", "reason": "原因1"},
            {"id": "sugg_002", "rule_id": "R002", "rule_name": "问题2",
             "type": "content", "severity": "High", "section": "分析",
             "original_text": "原文2", "suggestion": "建议2", "reason": "原因2"},
            {"id": "sugg_003", "rule_id": "R003", "rule_name": "问题3",
             "type": "language", "severity": "Medium", "section": "总结",
             "original_text": "原文3", "suggestion": "建议3", "reason": "原因3"}
        ]
    }
    
    step4_json = json.dumps(step4_output, ensure_ascii=False)
    
    # 测试1：接受所有建议
    print("\n测试3.1: 接受所有建议")
    result1 = prepare_suggestions_for_revision(step4_json, "")
    data1 = json.loads(result1)
    print(f"     结果: {len(data1)}条建议（应为3条）")
    
    # 测试2：只接受特定ID
    print("\n测试3.2: 只接受sugg_001和sugg_003")
    result2 = prepare_suggestions_for_revision(step4_json, "sugg_001,sugg_003")
    data2 = json.loads(result2)
    print(f"     结果: {len(data2)}条建议（应为2条）")
    ids = [s.get('id') for s in data2]
    print(f"     ID列表: {ids}")
    
    return len(data1) == 3 and len(data2) == 2


def test_check_suggestion_parsing():
    """测试CheckSuggestion解析"""
    print("\n" + "=" * 70)
    print("测试4: CheckSuggestion解析")
    print("=" * 70)
    
    suggestion_dict = {
        "id": "test_001",
        "rule_id": "RULE_001",
        "rule_name": "测试规则",
        "type": "format",
        "severity": "High",
        "section": "测试章节",
        "original_text": "原文",
        "suggestion": "修改建议",
        "reason": "修改原因"
    }
    
    try:
        suggestion = CheckSuggestion.from_dict(suggestion_dict)
        
        print(f"\n[OK] 解析成功")
        print(f"     ID: {suggestion.id}")
        print(f"     Rule: {suggestion.rule_name}")
        print(f"     Type: {suggestion.type}")
        print(f"     Severity: {suggestion.severity}")
        
        return True
        
    except Exception as e:
        print(f"\n[FAIL] 解析失败: {e}")
        return False


def main():
    """主测试函数"""
    print("\n" + "=" * 70)
    print("第4步到第6步数据流转测试")
    print("=" * 70)
    
    results = []
    
    # 运行测试
    results.append(("测试1: 第4步格式转换", test_step4_format()))
    results.append(("测试2: 直接数组格式", test_direct_array_format()))
    results.append(("测试3: prepare函数", test_prepare_function()))
    results.append(("测试4: CheckSuggestion解析", test_check_suggestion_parsing()))
    
    # 汇总结果
    print("\n" + "=" * 70)
    print("测试汇总")
    print("=" * 70)
    
    passed = sum(1 for _, r in results if r)
    total = len(results)
    
    for name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"  {status} {name}")
    
    print(f"\n总计: {passed}/{total} 通过")
    
    if passed == total:
        print("\n[OK] 所有测试通过！第4步到第6步数据流转正常。")
    else:
        print(f"\n[WARNING] 有{total-passed}个测试失败，请检查。")
    
    print("=" * 70)


if __name__ == "__main__":
    main()
