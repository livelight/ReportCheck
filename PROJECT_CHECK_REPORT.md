# 项目检查报告 - 文档检查工作流

**检查日期**: 2024年
**项目**: 银行IT系统生产事件分析报告检查工作流

---

## 一、工作流架构总览

```
┌─────────────┐     ┌──────────────────────┐     ┌──────────────────────┐
│ 1. 开始节点  │────▶│ 2. 硬性规则检查       │────▶│ 3. 软性规则检查       │
│  (读入文档)  │     │  (prompt_step2)       │     │  (prompt_step3_v2)    │
└─────────────┘     └──────────────────────┘     └──────────────────────┘
                              │                             │
                              └──────────────┬──────────────┘
                                             ▼
                              ┌──────────────────────┐
                              │ 4. 合并去重排序       │
                              │  (prompt_step4)       │
                              └──────────────────────┘
                                             │
                              ┌──────────────┴──────────────┐
                              ▼                             ▼
               ┌──────────────────────┐     ┌──────────────────────┐
               │ 5. 用户确认           │────▶│ 6. MCP文档修订节点    │
               │  (选择接受/拒绝)      │     │  (revision_mcp_node)  │
               └──────────────────────┘     └──────────────────────┘
                                                        │
                                                        ▼
                                        ┌──────────────────────┐
                                        │ 7. LLM复核            │
                                        │  (修订质量检查)        │
                                        └──────────────────────┘
```

---

## 二、各步骤输入输出一致性检查

### 第2步：硬性规则检查 ✅ 正常

| 检查项 | 状态 | 说明 |
|--------|------|------|
| **输入** | ✅ | 文档内容 (document_content) |
| **输出格式** | ✅ | JSON数组，包含id, rule_id, rule_name, type, severity, section, original_text, suggestion, reason |
| **type类型** | ✅ | naming/format/punctuation/figure/style/numbering/consistency |
| **severity** | ✅ | Critical/High/Medium/Low |

**与第4步兼容性**: ✅ 输出格式与第4步输入要求一致

---

### 第3步：软性规则检查 ✅ 正常

| 检查项 | 状态 | 说明 |
|--------|------|------|
| **输入** | ✅ | 文档内容 (document_content) |
| **输出格式** | ✅ | JSON数组，包含id, rule_id, rule_name, type, severity, section, original_text, suggestion, reason |
| **type类型** | ✅ | content/logic/language/format |
| **severity** | ✅ | Critical/High/Medium/Low |

**与第4步兼容性**: ✅ 输出格式与第4步输入要求一致

---

### 第4步：合并去重排序 ⚠️ 需要优化

| 检查项 | 状态 | 说明 |
|--------|------|------|
| **输入** | ✅ | hard_rules_results + soft_rules_results (两个JSON数组) |
| **输出格式** | ⚠️ | 包含suggestions数组和statistics对象 |

**问题发现**:

```json
// 第4步当前输出格式
{
  "suggestions": [...],  // 建议数组
  "statistics": {...},   // 统计信息
  "processing_notes": [...]
}
```

```python
// 第6步代码期望的输入 (CheckSuggestion.from_dict)
{
  "id": "xxx",
  "rule_id": "xxx", 
  "rule_name": "xxx",
  "type": "xxx",
  "severity": "xxx",
  "section": "xxx",
  "original_text": "xxx",
  "suggestion": "xxx",
  "reason": "xxx"
}
```

**⚠️ 兼容性问题**: 
第4步输出的是包含`suggestions`的对象，但第6步期望直接接收建议数组。

**建议修改**:
1. 第5步（用户确认）输出时，提取`suggestions`数组传递给第6步
2. 或在第4步提示词中明确说明输出格式适配要求

---

### 第6步：MCP文档修订节点 ⚠️ 需要优化

| 检查项 | 状态 | 说明 |
|--------|------|------|
| **输入参数** | ✅ | file_path, suggestions_json, output_path, use_track_changes |
| **suggestions_json格式** | ⚠️ | 期望数组格式，但需要确认与第4/5步输出兼容 |

**代码检查** (revision_mcp_node.py:567-582):

```python
# 第6步代码解析suggestions_json
suggestions_data = json.loads(suggestions_json)
if not isinstance(suggestions_data, list):
    raise ValueError("suggestions_json必须是数组格式")
    
suggestions = [CheckSuggestion.from_dict(s) for s in suggestions_data]
```

**问题发现**:

1. **第4步输出 vs 第6步输入不匹配**:
   - 第4步输出: `{"suggestions": [...], "statistics": {...}}`
   - 第6步期望: `[...]` (直接是数组)

2. **字段映射问题**:
   - 第4步输出字段包含: id, rule_id, rule_name, type, severity, section, original_text, suggestion, reason, source, group, order
   - 第6步CheckSuggestion只需要: id, rule_id, rule_name, type, severity, section, original_text, suggestion, reason
   - **多余的字段**: source, group, order 不影响，会被忽略

**解决方案**:

```python
# 在第5步到第6步的数据流转中，添加转换逻辑

def transform_step4_to_step6(step4_output):
    """将第4步输出转换为第6步输入格式"""
    suggestions = step4_output.get("suggestions", [])
    # 只保留第6步需要的字段
    filtered = []
    for s in suggestions:
        filtered.append({
            "id": s.get("id"),
            "rule_id": s.get("rule_id"),
            "rule_name": s.get("rule_name"),
            "type": s.get("type"),
            "severity": s.get("severity"),
            "section": s.get("section"),
            "original_text": s.get("original_text"),
            "suggestion": s.get("suggestion"),
            "reason": s.get("reason")
        })
    return filtered
```

---

## 三、关键问题与优化建议

### 问题1: 第4步输出格式与第6步期望不匹配 ❌

**严重程度**: High

**问题描述**:
- 第4步输出包装在`{"suggestions": [...]}`对象中
- 第6步期望直接接收数组`[...]`

**影响**:
- 第6步会解析失败，提示"suggestions_json必须是数组格式"

**解决方案** (选择其一):

**方案A**: 修改第5步输出逻辑 (推荐)
```python
# 在第5步用户确认后，提取suggestions数组
step4_output = json.loads(step4_result)
suggestions_array = step4_output["suggestions"]  # 提取数组
# 传递给第6步
step6_input = json.dumps(suggestions_array, ensure_ascii=False)
```

**方案B**: 修改第6步代码，支持对象格式
```python
# 在第6步代码中添加适配逻辑
data = json.loads(suggestions_json)
if isinstance(data, dict) and "suggestions" in data:
    suggestions_data = data["suggestions"]
elif isinstance(data, list):
    suggestions_data = data
else:
    raise ValueError("格式不正确")
```

---

### 问题2: 第4步建议中的`source`字段未使用 ❓

**严重程度**: Low

**问题描述**:
- 第4步输出包含`source`字段(hard/soft/merged)
- 第6步代码未使用此字段

**建议**:
- 保留该字段用于日志记录或调试
- 或在第6步添加统计信息："来自硬性规则X条，软性规则Y条"

---

### 问题3: 第4步建议中的`group`字段未使用 ❓

**严重程度**: Low

**问题描述**:
- 第4步输出可能包含`group`字段标识关联建议
- 第6步代码未利用此信息进行批量处理

**建议**:
- 可选优化：第6步可按group批量处理相关联的建议
- 当前实现已能正确处理，非必须修改

---

### 问题4: type类型定义不一致 ⚠️

**严重程度**: Medium

**问题描述**:

| 步骤 | type定义 | 问题 |
|------|---------|------|
| 第2步 | naming/format/punctuation/figure/style/numbering/consistency | 7种 |
| 第3步 | content/logic/language/format | 4种 |
| 第6步代码 | IssueType枚举: FORMAT/CONTENT/LANGUAGE/LOGIC | 4种 |

**问题**:
- 第2步的naming/punctuation/figure/style/numbering/consistency在第6步代码中没有对应
- 第6步代码只定义了4种类型，但第2步有7种

**影响**:
- 代码中`IssueType`枚举未实际使用于CheckSuggestion
- CheckSuggestion.type是字符串类型，不会触发枚举校验
- **实际无影响**，但建议保持统一

**建议**:
```python
# 当前代码实现 (无影响)
@dataclass
class CheckSuggestion:
    type: str  # 字符串类型，接受任意值
    
# 如果改为枚举校验，则需要统一
class IssueType(Enum):
    NAMING = "naming"
    FORMAT = "format"
    PUNCTUATION = "punctuation"
    FIGURE = "figure"
    STYLE = "style"
    NUMBERING = "numbering"
    CONSISTENCY = "consistency"
    CONTENT = "content"
    LOGIC = "logic"
    LANGUAGE = "language"
```

---

### 问题5: 缺少第5步（用户确认）提示词 ❌

**严重程度**: High

**问题描述**:
- 工作流设计中包含第5步用户确认
- 项目文件中缺少第5步的提示词设计

**建议**:
创建`prompt_step5_user_confirmation.md`，包含:
1. 展示第4步输出的建议列表
2. 允许用户选择接受/拒绝每条建议
3. 提供批量操作（全选/全不选/按severity筛选）
4. 输出用户确认后的建议数组

---

### 问题6: 缺少第7步（复核）提示词 ❌

**严重程度**: Medium

**问题描述**:
- 工作流设计中包含第7步LLM复核
- 项目文件中缺少第7步的提示词设计

**建议**:
创建`prompt_step7_review.md`，包含:
1. 输入：修订后的文档
2. 检查：修订是否准确、是否遗漏、是否引入新问题
3. 输出：复核结果（通过/需修正）

---

## 四、数据流转验证

### 完整数据流验证

```
第1步 ──▶ 第2步 ──▶ 第4步 ──▶ 第5步 ──▶ 第6步
        ──▶ 第3步 ──▶       

数据格式:
1. docx文件
   ↓
2. JSON数组 [CheckSuggestion, ...]
   ↓
3. JSON数组 [CheckSuggestion, ...]  
   ↓
4. {"suggestions": [...], "statistics": {...}}  ⚠️ 需要转换
   ↓
5. JSON数组 [CheckSuggestion, ...]  (用户筛选后)
   ↓
6. 修订后的docx文件
```

### 关键字段映射检查

| 字段 | 第2步 | 第3步 | 第4步 | 第5步 | 第6步代码 | 一致性 |
|------|-------|-------|-------|-------|----------|--------|
| id | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| rule_id | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| rule_name | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| type | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| severity | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| section | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| original_text | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| suggestion | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| reason | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

**结论**: 核心字段完全一致 ✅

---

## 五、优化建议汇总

### 高优先级 (必须修复)

1. **修复第4步到第6步的数据格式转换**
   - 方案：在第5步输出时提取`suggestions`数组
   - 或修改第6步代码支持对象格式

2. **补充第5步用户确认提示词**
   - 文件：`prompt_step5_user_confirmation.md`
   - 功能：展示建议列表，支持接受/拒绝选择

### 中优先级 (建议优化)

3. **补充第7步复核提示词**
   - 文件：`prompt_step7_review.md`
   - 功能：复核修订质量

4. **统一type类型定义**
   - 将第2步的7种类型纳入统一枚举
   - 便于后续维护和统计

### 低优先级 (可选优化)

5. **利用第4步的source字段**
   - 在第6步添加统计：硬性规则X条，软性规则Y条

6. **利用第4步的group字段**
   - 按关联组批量处理建议

---

## 六、修复代码示例

### 方案A: 第5步输出转换 (推荐)

```python
# 在第5步后置处理中添加
def prepare_for_step6(step4_output_json):
    """准备第6步输入数据"""
    step4_data = json.loads(step4_output_json)
    
    # 提取建议数组
    suggestions = step4_data.get("suggestions", [])
    
    # 用户确认后，过滤掉被拒绝的建议
    # (假设用户界面已完成选择)
    accepted_suggestions = [s for s in suggestions if s.get("accepted", True)]
    
    # 移除第6步不需要的字段
    for s in accepted_suggestions:
        s.pop("source", None)
        s.pop("group", None)
        s.pop("order", None)
        s.pop("accepted", None)
    
    return json.dumps(accepted_suggestions, ensure_ascii=False)
```

### 方案B: 第6步输入适配

```python
# 在revision_mcp_node.py revise_document函数中添加
def revise_document(file_path, suggestions_json, ...):
    # ... 现有代码 ...
    
    # 适配第4步输出格式
    suggestions_data = json.loads(suggestions_json)
    
    if isinstance(suggestions_data, dict):
        # 第4步格式: {"suggestions": [...], "statistics": {...}}
        suggestions_data = suggestions_data.get("suggestions", [])
    elif not isinstance(suggestions_data, list):
        raise ValueError("suggestions_json格式不正确")
    
    # 继续处理...
    suggestions = [CheckSuggestion.from_dict(s) for s in suggestions_data]
```

---

## 七、检查结论

| 检查维度 | 结果 | 说明 |
|---------|------|------|
| 架构设计 | ✅ | 7步工作流设计合理 |
| 第2步提示词 | ✅ | 完整，输出格式正确 |
| 第3步提示词 | ✅ | 完整，输出格式正确 |
| 第4步提示词 | ✅ | 完整，但输出格式需适配 |
| 第5步提示词 | ❌ | 缺失，需要补充 |
| 第6步代码 | ⚠️ | 功能完整，但需适配第4步输出 |
| 第7步提示词 | ❌ | 缺失，需要补充 |
| 字段一致性 | ✅ | 核心字段完全一致 |
| 数据流转 | ⚠️ | 第4→5→6步需要格式转换 |

### 总体评估

**项目完成度**: 75%

**主要问题**:
1. 第4步输出格式与第6步期望不匹配 (影响流程)
2. 缺少第5步用户确认提示词 (影响功能完整性)
3. 缺少第7步复核提示词 (影响功能完整性)

**建议**: 
- 优先修复第4→6步的数据格式问题
- 补充第5、7步提示词
- 项目即可达到生产可用状态
