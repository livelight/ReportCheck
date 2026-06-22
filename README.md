# 文档检查工作流MCP服务

## 功能概述

本项目提供完整的文档检查与修订工作流MCP服务，包括：

| 服务 | 文件 | 功能 |
|------|------|------|
| **文档读取 + 修订** | `unified_mcp_server.py` | 读取文档、根据检查建议自动修订文档 |
| **文档读取** | `document_reader_mcp.py` | 独立文档读取服务，支持智能分段 |
| **时间轴图生成** | `timeline_mcp_server.py` | 根据事件处置数据生成专业时间轴图 |

支持三种修订模式和自动生成修改记录文档，适用于企业文档自动化检查、AI辅助写作、合规性审查等场景。

## 项目背景与目标

### 业务场景
在大模型应用（如Dify工作流）中，需要自动化处理企业文档，包括：
1. **文档检查** - 发现格式、内容、语言、逻辑问题
2. **生成建议** - 基于检查规则给出修改建议
3. **自动修订** - 根据建议自动修改文档
4. **复核确认** - 验证修订结果

### 项目定位
本项目提供完整的文档检查工作流MCP服务：
- **第1步 - 文档读取MCP节点**：读取文档内容，支持智能分段
- **第6步 - 文档修订MCP节点**：根据检查建议自动修订文档

推荐使用**统一MCP服务器**（`unified_mcp_server.py`）同时运行第1步和第6步服务。

## 整体架构（7步工作流）

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│ 1. 开始节点  │────▶│ 2. LLM节点  │────▶│ 3. LLM节点  │
│  (读入文档)  │     │ (硬性规则检查)│     │ (软性规则检查)│
└─────────────┘     └─────────────┘     └─────────────┘
                                                 │
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│ 7. LLM节点  │◀────│ 6. MCP节点  │◀────│ 4. LLM节点  │
│  (复核文档)  │     │  (本文档修订) │     │(建议合并去重) │
└─────────────┘     └─────────────┘     └─────────────┘
                             ▲
                             │
                      ┌─────────────┐
                      │ 5. 用户输入  │
                      │(选择采纳建议)│
                      └─────────────┘
```

## 核心能力

| 功能 | 说明 |
|------|------|
| **文档解析** | 支持DOCX、WPS、WPSX格式 |
| **修订应用** | 根据LLM建议自动替换文本，支持三种修订模式 |
| **修订追踪** | Word Track Changes（w:del/w:ins）、红色括号批注、直接替换 |
| **修改记录明细** | 自动生成修改记录文档，记录每次修改的明细，支持后续追加 |
| **时间轴图生成** | 根据事件处置数据生成专业PNG时间轴图，支持横向递进、蛇形多行、同时刻多事件 |
| **格式兼容** | 生成的文档可被WPS/Word正常打开 |

### 技术特点

1. **FastMCP框架** - 使用`@mcp.tool()`注册工具，可被Dify等MCP客户端调用
2. **三种修订模式** - Track Changes模式（红色删除线+蓝色下划线）、红色括号批注模式、直接替换模式
3. **异常处理** - 所有工具都有try-except，返回结构化错误信息
4. **日志记录** - 详细记录调用时间、输入输出、错误信息

## 支持格式

- Word文档 (.docx)
- WPS文档 (.wps, .wpsx)

## 安装依赖

```bash
pip install -r requirements.txt
```

### 主要依赖

| 依赖包 | 版本要求 | 用途 |
|--------|----------|------|
| `python-docx` | >=0.8.11 | Word文档解析和生成 |
| `lxml` | >=4.9.0 | XML处理（OOXML格式） |
| `fastmcp` | >=0.4.0 | MCP框架支持 |
| `matplotlib` | >=3.5.0 | 时间轴图生成 |
| `pillow` | >=9.0.0 | 图片处理 |
| `flask` | >=2.0.0 | HTTP API服务（可选） |
| `fastapi` | >=0.104.0 | SSE服务器（可选） |
| `uvicorn` | >=0.23.0 | ASGI服务器（可选） |
| `sse-starlette` | >=1.6.0 | SSE流支持（可选） |

### LLM节点 Temperature 设置

在Dify中配置各步骤的LLM节点时，Temperature设置如下：

| 步骤 | 任务类型 | Temperature | 说明 |
|------|----------|-------------|------|
| **第2步 硬性规则检查** | 格式规范检查 | **0.2-0.3** | 需要精确、稳定地按照明确规则执行，输出必须一致 |
| **第3步 软性规则检查** | 内容质量检查 | **0.2-0.4** | 需要一定灵活性评估内容深度和逻辑，但保持稳定性 |
| **第4步 合并去重排序** | 结果整合 | **0.3-0.5** | 需要判断和整合能力，适当灵活性有助于优化建议 |
| **第5步 用户确认** | 交互展示 | **0-0.1** | 主要是格式化和展示，不需要创造性 |
| **第6步 文档修订** | MCP工具执行 | **不适用** | 非LLM步骤，代码直接执行 |
| **第7步 复核** | 质量评估 | **0.3-0.5** | 需要综合判断文档质量 |

#### Temperature配置原则

```
低Temperature (0.0-0.3): 需要精确、一致、按规则执行的任务
中Temperature (0.3-0.5): 需要判断、整合、评估的任务
高Temperature (>0.5): 本项目不适用（不需要创造性生成）
```

#### Dify中的具体配置

**第2步 - 硬性规则检查**
```yaml
Model: GPT-4 / Claude-3.5-Sonnet
Temperature: 0.2
Max Tokens: 8000-12000
```

**第3步 - 软性规则检查**
```yaml
Model: GPT-4 / Claude-3.5-Sonnet
Temperature: 0.3
Max Tokens: 8000-12000
```

**第4步 - 合并去重排序**
```yaml
Model: GPT-4 / Claude-3.5-Sonnet
Temperature: 0.4
Max Tokens: 8000-12000
```

**第7步 - 复核**
```yaml
Model: GPT-4 / Claude-3.5-Sonnet
Temperature: 0.4
Max Tokens: 6000-8000
```

#### 调整建议

- 如果某步输出不稳定（同一输入结果变化大）：**降低 Temperature 0.1**
- 如果需要更多灵活性（如第4步合并建议）：**提高 Temperature 0.1**
- 硬性规则检查（第2步）建议不超过 **0.3**，确保规则执行严格性

## 近期改进（P0/P1/P2）

### P0: 移除无法实现的规则

基于技术可行性分析，从硬性规则中移除了以下依赖Word UI或页面布局信息的规则：

| 规则ID | 规则名称 | 移除原因 |
|--------|----------|----------|
| FORMAT_INDENT_001 | 首行缩进2字符 | python-docx无法访问Word的"标尺"缩进概念 |
| FORMAT_MARGIN_001 | 页边距标准 | 无法读取页面边距设置（OOXML不存储实际边距值） |
| TABLE_PAGE_001 | 表格跨页处理 | 无法检测分页（渲染后信息） |
| FIGURE_BORDER_001 | 图片边框线粗细 | 无法读取边框线粗细属性 |

同时从软性规则中移除了 **SECURITY_SENSITIVE_001**（敏感信息检查），因该检查需要安全策略配置，不适合作为通用规则。

### P1: 文档智能分段

**document_reader_mcp.py v1.1.0** 新增智能分段功能：

```python
# 使用分段读取
read_document_segmented(file_path, max_tokens_per_segment=4000)
```

- 自动估算token数（中文1字≈1token，英文4字符≈1token）
- 保持段落完整性，不切割单个段落
- 优先在章节边界处分割（识别"一、"、"1."、"（一）"等模式）
- 表格作为独立分段处理
- 返回分段统计信息（总分段数、平均每段token数等）

### P2: 三种修订模式支持

**unified_mcp_server.py v1.1.0** 支持三种文档修订模式，通过 `use_track_changes` 参数控制：

| 模式 | 参数值 | 说明 |
|------|--------|------|
| **Word Track Changes** | `"true"` | 使用w:del/w:ins元素，可在Word中查看/接受/拒绝修订 |
| **红色括号批注** | `"comment"` | 保留原文，在段落后插入红色括号批注（默认模式） |
| **直接替换** | `"false"` | 直接替换文本，不保留修订痕迹 |

**Word Track Changes 模式示例**：

```xml
<!-- 实际生成的OOXML结构 -->
<w:del w:id="1000" w:author="DocumentReviser" w:date="2024-01-15T10:30:00Z">
  <w:r><w:t>原文</w:t></w:r>
</w:del>
<w:ins w:id="1001" w:author="DocumentReviser" w:date="2024-01-15T10:30:00Z">
  <w:r><w:t>修改后文本</w:t></w:r>
</w:ins>
```

**红色括号批注模式示例**：

```
原文内容（【建议】修改后内容 【原因】修改原因说明）
```

## 核心文件

| 文件 | 说明 |
|------|------|
| `unified_mcp_server.py` | **统一MCP服务器 v1.1.0（推荐）** - 同时启动第1步（文档读取）和第6步（文档修订）服务，支持Word原生Track Changes，联动启动时间轴图服务 |
| `timeline_mcp_server.py` | **时间轴图MCP服务器 v1.0.0** - 根据事件处置数据生成专业PNG时间轴图，支持横向递进/蛇形多行/同时刻多事件 |
| `document_reader_mcp.py` | 第1步文档读取MCP节点 v1.1.0（独立运行，支持智能分段） |
| `revision_mcp_node.py` | 第6步文档修订MCP节点（独立运行，旧版，已不推荐使用） |
| `prompt_step2_hard_rules.md` | 第2步硬性规则检查提示词 |
| `prompt_step3_soft_rules_v2.md` | 第3步软性规则检查提示词（去重优化版） |
| `prompt_step4_merge_prioritize.md` | 第4步合并去重排序提示词 |
| `prompt_step5_user_confirmation.md` | 第5步用户确认提示词 |
| `mcp_server.py` | HTTP API服务（Flask，同步方式） |
| `mcp_sse_server.py` | SSE服务器（FastAPI，流式方式，推荐） |
| `tests/` | 测试文件夹，包含所有测试脚本 |
| `check_revisions.py` | 修订标记验证工具 |
| `PROJECT_CHECK_REPORT.md` | 项目检查报告 |
| `OPTIMIZATION_SUMMARY.md` | 优化总结报告 |
| `FINAL_CHECK_REPORT.md` | 最终检查报告 |

## 代码结构

### unified_mcp_server.py（推荐）

```
unified_mcp_server.py
├── DocumentMetadata         # 文档元数据模型
├── CheckSuggestion          # 建议数据模型
├── Severity                 # 严重程度枚举
├── DocumentReader           # 文档读取器（第1步）
│   ├── read_document()      # 读取文档主入口
│   ├── _read_docx()         # Word文档解析
│   ├── _read_wps()          # WPS文档解析
│   └── _extract_docx_metadata()  # 提取文档元数据
├── DocumentReviser          # 文档修订器（第6步）
│   ├── revise_document()    # 修订文档主入口
│   ├── _revise_docx()       # Word文档修订
│   ├── _revise_wps()        # WPS文档修订
│   ├── _apply_track_changes_to_paragraph()  # 应用Track Changes（核心）
│   ├── _create_track_changes_element()      # 创建w:del/w:ins元素
│   ├── _add_red_comment_to_paragraph()      # 添加红色括号批注（comment模式）
│   ├── _write_revision_log()                # 生成修改记录明细文档
│   └── _get_revision_log_path()             # 获取修改记录文档路径
└── MCP工具 (5个)
    ├── read_document()          # 第1步 - 读取文档
    ├── get_document_info()      # 第1步 - 获取文档信息
    ├── revise_document()        # 第6步 - 修订文档
    ├── validate_suggestions()   # 第6步 - 验证建议JSON
    └── get_tools_info()         # 获取工具信息
```

### timeline_mcp_server.py

```
timeline_mcp_server.py
├── TimelineChartGenerator     # 时间轴图生成器
│   ├── validate_events_data() # 验证事件数据格式
│   ├── generate_timeline()    # 生成时间轴PNG图（横向递进/蛇形多行）
│   ├── generate_timeline_v2() # 字符串参数版本（MCP工具调用）
│   ├── get_timeline_data_template()  # 获取数据模板
│   └── NODE_STYLES            # 节点样式配置（开始/结束/普通/特殊）
└── MCP工具 (5个)
    ├── read_document()         # 读取文档（复用unified_mcp_server）
    ├── generate_timeline()     # 生成时间轴图
    ├── validate_timeline_data()# 验证时间轴数据
    ├── get_timeline_template() # 获取数据模板
    └── get_tools_info()        # 获取工具信息
```

### document_reader_mcp.py

```
document_reader_mcp.py
├── DocumentMetadata         # 文档元数据模型
├── DocumentSegment          # 文档分段模型（v1.1.0新增）
├── DocumentReader           # 文档读取器
│   ├── read_document()      # 读取文档
│   ├── _read_docx()         # Word解析
│   └── _read_wps()          # WPS解析
├── DocumentSegmenter        # 文档分段器（v1.1.0新增）
│   ├── segment_document()   # 智能分段
│   ├── estimate_tokens()    # 估算token数
│   └── get_segment_info()   # 获取分段统计
└── MCP工具 (4个)
    ├── read_document()           # 读取文档
    ├── read_document_segmented() # 分段读取文档（v1.1.0新增）
    ├── get_document_info()       # 获取文档信息
    └── get_tools_info()          # 工具信息
```

## 快速开始

### 1. 安装依赖并测试

```bash
pip install -r requirements.txt
python test_all.py
```

### 2. 启动服务（推荐SSE方式）

```bash
python mcp_sse_server.py --port 8000
```

### 3. 在Dify中配置MCP节点

```
URL: http://localhost:8000/revise
Method: POST
Headers:
  Content-Type: application/json
Body:
{
    "file_path": "/path/to/document.docx",
    "suggestions_json": "{{ step_5_output.suggestions }}"
}
```

### 4. Python API直接调用

```python
from unified_mcp_server import DocumentReader, DocumentReviser
import json

# 第1步 - 读取文档
reader = DocumentReader()
result = reader.read_document("原文档.docx")
print(f"段落数: {result['document_metadata']['paragraph_count']}")

# 第6步 - 修订文档
reviser = DocumentReviser()
suggestions = [
    {
        "id": "ref_001",
        "original_text": "原文",
        "suggestion": "修改后文本"
    }
]
result = reviser.revise_document(
    file_path="原文档.docx",
    suggestions_json=json.dumps(suggestions),
    output_path="修订后文档.docx",
    use_track_changes="comment"  # 默认comment模式: "true"=Track Changes, "comment"=红色批注, "false"=直接替换
)
print(f"应用修订数: {result['applied_revisions']}")
```

### 5. 时间轴图生成

```python
from timeline_mcp_server import TimelineChartGenerator

data = {
    "title": "事件处置时间轴",
    "incident_id": "INC-001",
    "events": [
        {"time": "14:00", "descriptions": ["告警触发"], "node_type": "start"},
        {"time": "14:30", "descriptions": ["分析定位", "执行操作"], "node_type": "normal"},
        {"time": "16:00", "descriptions": ["关闭事件"], "node_type": "end"},
    ]
}
result = TimelineChartGenerator.generate_timeline(data, output_path="timeline.png")
print(f"时间轴图已生成: {result['output_path']}")
```

**详细运行命令请参考下方 [运行命令参考](#运行命令参考) 章节**

## 数据流转

### 输入（来自第5步用户选择）
```json
[
  {
    "id": "ref_001",
    "rule_id": "CONTENT_023",
    "rule_name": "缺少风险分析",
    "type": "content",
    "severity": "High",
    "section": "sec_005",
    "original_text": "项目进展顺利，暂无风险。",
    "suggestion": "当前主要风险：1）人员风险；2）技术风险。",
    "reason": "风险分析过于简单"
  }
]
```

### 处理过程
1. **解析原文档** - 使用`python-docx`读取段落和文本
2. **查找匹配** - 在文档中定位`original_text`
3. **添加修订标记** - 插入OOXML标准的`w:del`和`w:ins`元素
4. **保存新文档** - 生成带修订标记的文档，可被WPS/Word识别

### 输出
```json
{
  "success": true,
  "output_path": "文档_revised.docx",
  "use_track_changes": true,
  "total_suggestions": 1,
  "applied_revisions": 1,
  "revisions_detail": [...],
  "statistics": {
    "by_severity": {"High": 1},
    "by_type": {"content": 1}
  }
}
```

## 输出格式

```json
{
    "success": true,
    "output_path": "原文档_revised.docx",
    "total_suggestions": 4,
    "applied_revisions": 4,
    "revisions_detail": [
        {
            "id": "ref_001",
            "rule_id": "FORMAT_001",
            "rule_name": "标题格式不规范",
            "type": "format",
            "severity": "High",
            "original": "一、项目概述",
            "new": "一、项目概述",
            "status": "success"
        }
    ],
    "skipped_revisions": 0,
    "statistics": {
        "by_severity": {
            "Critical": 0,
            "High": 2,
            "Medium": 1,
            "Low": 1
        },
        "by_type": {
            "format": 1,
            "content": 1,
            "language": 1,
            "logic": 1
        }
    }
}
```

## 工作流节点说明

| 节点 | 功能 |
|------|------|
| 1. 开始节点 | 读入文档 |
| 2. LLM节点(硬性规则) | 根据强制要求检查并输出建议 |
| 3. LLM节点(软性规则) | 根据建议规则检查并输出建议 |
| 4. LLM节点(合并去重) | 整合建议，标记严重程度和类别，输出JSON |
| 5. 用户输入 | 用户决定采纳哪些修改建议 |
| 6. **MCP节点(本文)** | 根据采纳的建议修订文档 |
| 7. LLM节点(复核) | 对生成的文档进行复核 |

## 修订模式技术实现

支持三种修订模式，适用于不同场景：

### 模式一：Word Track Changes（use_track_changes="true"）

...（中间内容省略，保持不变）...

### 模式三：直接替换（use_track_changes="false"）

直接替换原文，不保留任何修订痕迹：

特点：
- 原文被直接替换为新内容
- 文档中看不到修改历史
- 适合已经确认无误的最终修订

## 修改记录明细文档

每次修订完成后，自动在与原文档同目录下生成修改记录明细文档 **`原文件名_revision_log.docx`**。

文档格式为方便阅读的DOCX格式，包含以下内容：

### 文档结构

```
文件修改记录明细（标题）
被修改文件：xxx.docx

第1次修改
修改时间：2026-06-16 15:30:00
修改概述：共修订N条内容，成功应用M条；严重程度分布：High: 2条、Medium: 1条；类型分布：content: 2条、language: 1条

修改内容明细：
1. [✅] ref_001
   规则：缺少风险分析（CONTENT_023）
   类型：content　严重程度：High
   状态：已应用
   原文：项目进展顺利，暂无风险。
   修改后：当前主要风险：人员风险和技术风险。
   原因：风险分析过于简单

2. [✅] ref_002
   ...（省略）
```

### 功能特点

- **自动生成**：每次 `revise_document` 调用成功后自动生成
- **支持追加**：后续再次修改同一文档时，自动追加新的修改记录（用分隔线区分）
- **修改次数自增**：自动统计已有修改次数，递增记录"第N次修改"
- **状态标记**：每条修改记录标记 ✅（已应用）或 ❌（未找到原文）
- **颜色标识**：修改后内容显示为深绿色，修改原因显示为灰色
- **返回值**：修订结果中包含 `revision_log_path` 字段，指向生成的修改记录文档路径

## 运行命令参考

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 测试命令

```bash
# 整体项目测试（推荐）
python test_all.py

# Step4到Step6数据流转测试
python test_step4_to_step6.py

# 基础功能测试
python test_revision.py

# FastMCP框架测试
python test_fastmcp.py

# 修订模式测试
python test_track_changes.py

# WPS/WPSX文档测试
python test_wps_track_changes.py

# SSE服务器测试（需先启动SSE服务）
python test_sse.py
```

### 3. 启动服务

#### **方式零: 统一MCP服务器（强烈推荐）**

同时启动第1步（文档读取）和第6步（文档修订）的所有MCP服务，**推荐用于生产环境**。

```bash
# stdio模式（默认）
python unified_mcp_server.py

# SSE模式
python unified_mcp_server.py --transport sse --host 0.0.0.0 --port 8000
```

可用工具：
| 工具名 | 步骤 | 功能 |
|--------|------|------|
| `read_document` | 第1步 | 读取文档内容和元数据 |
| `get_document_info` | 第1步 | 获取文档基本信息 |
| `revise_document` | 第6步 | 根据建议修订文档 |
| `validate_suggestions` | 第6步 | 验证建议JSON格式 |
| `get_tools_info` | - | 获取工具信息 |

测试：
```bash
python test_unified_server.py
```

#### 方式一: stdio模式（MCP标准）

```bash
python revision_mcp_node.py
```

#### 方式二: SSE模式（适合Dify集成）

**方法A: FastMCP内置SSE**

```bash
# 默认参数
python revision_mcp_node.py --transport sse

# 自定义参数
python revision_mcp_node.py --transport sse --host 0.0.0.0 --port 8000
```

端点：
- `GET /sse` - SSE连接端点
- `POST /messages` - 发送消息

**方法B: 独立SSE服务器（功能更丰富，推荐）**

```bash
# 默认参数（端口8000）
python mcp_sse_server.py

# 自定义参数
python mcp_sse_server.py --host 0.0.0.0 --port 8000

# 开发模式（热重载）
python mcp_sse_server.py --reload
```

端点：
| 端点 | 方法 | 说明 |
|------|------|------|
| `/sse` | GET | SSE连接（实时接收消息） |
| `/messages` | POST | 发送工具调用请求 |
| `/revise` | POST | 同步修订接口 |
| `/revise/stream` | POST | 流式修订（带进度事件） |
| `/health` | GET | 健康检查 |
| `/tools` | GET | 工具定义 |

#### 方式三: 第1步文档读取服务

```bash
# stdio模式
python document_reader_mcp.py

# SSE模式
python document_reader_mcp.py --transport sse --host 0.0.0.0 --port 8001
```

工具：
- `read_document` - 读取文档内容和元数据
- `read_document_segmented` - 分段读取文档（支持长文档，自动按章节/token数分割）
- `get_document_info` - 获取文档基本信息
- `get_tools_info` - 工具信息

**智能分段功能**（v1.1.0新增）：
- 自动将长文档分割成多个段落块（默认每段约4000 tokens）
- 保持段落完整性，优先在章节边界处分割
- 返回分段统计信息和各段内容
- 适用于超长文档的分布式处理

输出格式：
```json
{
  "success": true,
  "document_content": "文档完整文本内容...",
  "document_metadata": {
    "title": "文档标题",
    "author": "作者",
    "created_date": "创建日期",
    "paragraph_count": 12,
    "table_count": 2
  }
}
```

#### 方式四: HTTP API服务（Flask）

```bash
python mcp_server.py
```

默认运行在 `http://localhost:5000`

端点：
| 端点 | 方法 | 说明 |
|------|------|------|
| `/health` | GET | 健康检查 |
| `/tools` | GET | 工具定义 |
| `/revise` | POST | 修订文档（JSON） |
| `/revise/form` | POST | 修订文档（表单/文件上传） |

## MCP服务调用指南

### 第1步 - 读入文档（DocumentReaderMCPNode）

**所在文件**：`unified_mcp_server.py` 或 `document_reader_mcp.py`

| 工具名 | 功能 | 输入参数 | 输出内容 |
|--------|------|----------|----------|
| `read_document` | 读取完整文档内容 | `file_path`: 文档路径 | `document_content`, `document_metadata`, `structured_content` |
| `read_document_segmented` | **分段读取**（v1.1.0新增，长文档推荐） | `file_path`: 文档路径<br>`max_tokens_per_segment`: 每段最大token数（可选，默认4000） | `document_content`, `segments`, `segment_info` |
| `get_document_info` | 快速获取文档元数据（不读内容） | `file_path`: 文档路径 | 文件大小、创建时间、段落数等 |

**推荐用法**：
- 普通文档（<8000 tokens）：使用 `read_document`
- 长文档（>8000 tokens）：使用 `read_document_segmented`，自动按章节边界分割

**输出示例**：
```json
{
  "success": true,
  "document_content": "文档完整文本...",
  "document_metadata": {
    "title": "文档标题",
    "paragraph_count": 12,
    "table_count": 2
  },
  "structured_content": {
    "paragraphs": [{"index": 0, "text": "...", "style": "Normal"}],
    "tables": []
  }
}
```

---

### 第6步 - 修订文档（DocumentReviser）

**所在文件**：`unified_mcp_server.py`

| 工具名 | 功能 | 输入参数 | 输出内容 |
|--------|------|----------|----------|
| `revise_document` | 根据建议修订文档 | `file_path`: 原文档路径<br>`suggestions_json`: 建议JSON字符串<br>`output_path`: 输出路径（可选）<br>`use_track_changes`: 修订模式（可选，默认"true"） | `output_path`, `applied_revisions`, `revisions_detail` |
| `validate_suggestions` | 验证建议JSON格式 | `suggestions_json`: 建议JSON字符串 | 验证结果、建议数量 |

**关键参数说明**：
- `use_track_changes`: 支持三种修订模式
  - `"true"` - **Word Track Changes模式**：使用w:del/w:ins元素，可在Word中查看/接受/拒绝修订
  - `"comment"` - **红色括号批注模式**：保留原文，在文本后插入红色括号批注（格式：`（【建议】修改后内容 【原因】修改原因）`）
  - `"false"` - **直接替换模式**：直接替换文本，不保留修订痕迹

**输入示例**（suggestions_json）：
```json
[
  {
    "id": "ref_001",
    "rule_id": "CONTENT_023",
    "original_text": "项目进展顺利，暂无风险。",
    "suggestion": "当前主要风险：1）人员风险；2）技术风险。"
  }
]
```

**输出示例**：
```json
{
  "success": true,
  "output_path": "文档_revised.docx",
  "use_track_changes": true,
  "total_suggestions": 1,
  "applied_revisions": 1,
  "revisions_detail": [
    {
      "id": "ref_001",
      "rule_id": "CONTENT_023",
      "original": "项目进展顺利，暂无风险。",
      "new": "当前主要风险：1）人员风险；2）技术风险。",
      "status": "success"
    }
  ]
}
```

---

### 其他辅助MCP服务

| 工具名 | 功能 | 用途 |
|--------|------|------|
| `get_tools_info` | 获取所有工具定义 | Dify/MCP客户端配置时调用，获取工具列表和参数定义 |

---

### 在Dify中的配置示例

**第1步 - 文档读取节点配置**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: read_document 或 read_document_segmented
输入:
  file_path: {{ start.file_path }}
  # 如使用分段读取:
  # max_tokens_per_segment: 4000
输出变量:
  document_content: {{ document_content }}
  document_metadata: {{ document_metadata }}
```

**第6步 - 文档修订节点配置**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: revise_document
输入:
  file_path: {{ start.file_path }}
  suggestions_json: {{ step5.selected_suggestions }}
  use_track_changes: "true"  # 可选: "true"=Track Changes, "comment"=红色批注, "false"=直接替换
输出变量:
  output_path: {{ output_path }}
  applied_revisions: {{ applied_revisions }}
```

---

### 服务启动与调用关系

```
┌─────────────────────────────────────────────────────────────┐
│                    unified_mcp_server.py                     │
│  ┌─────────────────────┐      ┌──────────────────────────┐  │
│  │   第1步服务          │      │        第6步服务         │  │
│  │  - read_document    │      │   - revise_document      │  │
│  │  - read_document_   │─────▶│   - validate_suggestions │  │
│  │    segmented        │      │                          │  │
│  │  - get_document_    │      │                          │  │
│  │    info             │      │                          │  │
│  └─────────────────────┘      └──────────────────────────┘  │
│                           get_tools_info (共用)             │
└─────────────────────────────────────────────────────────────┘
```

**启动命令**：
```bash
# 方式一：统一服务器（推荐，同时支持第1步和第6步）
python unified_mcp_server.py --transport sse --port 8000

# 方式二：独立服务（分开部署）
# 第1步服务（端口8001）
python document_reader_mcp.py --transport sse --port 8001
# 第6步服务（端口8000）
python unified_mcp_server.py --transport sse --port 8000
```

## 项目效果分析与限制说明

### ✅ 实际能达到的效果

#### 1. 文档读取与解析（第1步）
| 能力 | 实现程度 | 说明 |
|------|----------|------|
| DOCX读取 | ✅ 完全支持 | python-docx稳定支持 |
| WPS/WPSX读取 | ⚠️ 基本支持 | 复用DOCX解析，大部分情况正常 |
| 文本提取 | ✅ 完整 | 段落、表格内容均可提取 |
| 智能分段 | ✅ 可用 | 按token估算分段，保持段落完整性 |

#### 2. 规则检查（第2、3步）
| 能力 | 实现程度 | 说明 |
|------|----------|------|
| 错别字/语法检查 | ✅ 较好 | LLM擅长语言类检查 |
| 格式规范检查 | ⚠️ 有限 | 依赖LLM文本分析，非精确格式验证 |
| 内容逻辑检查 | ✅ 较好 | LLM可评估内容质量和逻辑 |
| 一致性检查 | ✅ 较好 | 可检查术语、名称一致性 |

#### 3. 文档修订（第6步）
| 能力 | 实现程度 | 说明 |
|------|----------|------|
| 文本替换 | ✅ 支持 | 简单文本替换准确 |
| Word修订模式 | ⚠️ 基础支持 | 生成w:del/w:ins元素，但样式可能不完美 |
| 格式保留 | ⚠️ 有限 | 复杂格式可能丢失 |

### ⚠️ 技术限制与风险点

#### 1. 硬性规则检查的准确性问题

**问题**：LLM-based格式检查不够精确

```
目标：检查"字体字号一致"、"段落缩进2字符"
实际：LLM只能看到文本内容，无法访问：
  - 实际的字体名称和大小
  - 段落缩进的精确数值
  - 页边距设置
  - 行距、段间距等精确格式
```

**影响**：
- 格式检查依赖LLM的"推测"，可能误判
- 无法验证Word文档的实际样式设置
- 已移除的4条规则（缩进、页边距、表格跨页、图片边框）正是因为技术上无法实现

#### 2. 文档修订的精确性问题

**问题**：文本匹配可能不准确

```python
# 当前实现：简单字符串匹配
if original in para.text:
    # 替换文本
```

**风险场景**：
- 原文多次出现：可能替换错误位置
- 格式复杂的段落：可能破坏原有格式
- 表格单元格中的文本：处理不完善
- 修订痕迹重叠：多次修订可能冲突

#### 3. Word修订模式的局限性

**生成的Track Changes可能不完美**：
- 修订者信息（作者、时间戳）可能不完整
- 缺少格式属性（颜色、删除线样式）
- Word/WPS中可能显示为"无格式修订"
- 不保证在WPS中完全兼容

#### 4. 复杂文档元素的支持限制

| 元素 | 支持程度 | 问题 |
|------|----------|------|
| 纯文本段落 | ✅ 完整 | 无问题 |
| 简单表格 | ⚠️ 基本 | 可读取，但修订时可能破坏结构 |
| 图片/图表 | ❌ 不支持 | 无法读取和修订 |
| 页眉页脚 | ❌ 不支持 | python-docx不支持 |
| 目录/引用 | ❌ 不支持 | 无法处理 |
| 复杂格式 | ⚠️ 有限 | 超链接、特殊样式可能丢失 |

#### 5. 性能与成本问题

**LLM调用成本**：

| 步骤 | 估算成本（每篇文档） |
|------|---------------------|
| 第2步 硬性检查 | 输入全文 + 输出建议 |
| 第3步 软性检查 | 输入全文 + 输出建议 |
| 第4步 合并去重 | 前两步输出 + 合并结果 |
| 第7步 复核 | 修订后全文 + 评估 |

**长文档问题**：
- 超过8k tokens的文档需要分段处理
- 分段后上下文丢失，可能影响检查质量
- 多次LLM调用成本显著增加

### 📊 适用场景评估

#### ✅ 推荐使用（能达到目标）
- 短篇文档（<10页）的格式和内容检查
- 纯文本文档，格式简单的报告
- 辅助人工审核，提供修改建议
- 标准化程度高的文档模板检查
- 语言类问题检查（错别字、语法、表述）

#### ❌ 不推荐使用（可能达不到目标）
- 复杂格式文档（含图表、图片、页眉页脚）
- 精确格式合规检查（字体、字号、页边距）
- 完全无人值守的自动修订
- 长文档（>50页）的一次性处理
- 需要精确页面布局的文档

### 💡 使用建议

1. **降低期望定位**：定位为"辅助审核工具"而非"全自动检查系统"
2. **分场景使用**：
   - 语言类检查：可高度依赖
   - 格式类检查：仅作为参考
3. **增加人工确认**：第6步的修订建议必须人工审核后再应用
4. **分段处理优化**：长文档分段后需要更智能的上下文保留策略

## 注意事项

1. **WPS文档支持**：WPS/WPSX使用与DOCX相同的OOXML格式，可被WPS Office和Microsoft Word同时识别
2. **批量修订**：建议按顺序逐条应用，以确保准确性
3. **错误处理**：对于未找到的原文，会标记为`not_found`状态
4. **备份**：建议在修订前备份原文档
5. **编码问题**：中文环境下注意Windows终端编码设置

## 应用场景

- **企业文档自动化检查** - 合同、报告、规范文档的智能审查
- **AI辅助写作** - 自动修正语法、逻辑、格式问题
- **合规性审查** - 确保文档符合企业规范和标准
- **多轮迭代修订** - 支持接受/拒绝修订的交互流程
- **文档质量管控** - 统一文档格式和内容标准
