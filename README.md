# 文档检查与时间轴生成MCP服务

## 项目简介

本项目基于 MCP (Model Context Protocol) 协议和 FastMCP 框架，提供两个独立的工作流服务：

| 工作流 | 说明 | 核心文件 |
|--------|------|----------|
| **文档修订工作流** | 读取文档 → 获取系统信息 → 规则检查 → 合并建议 → 自动修订 → 复核 | `unified_mcp_server.py` |
| **时间轴图工作流** | 读取处置过程 → LLM分析 → 生成专业时间轴PNG图 | `timeline_mcp_server.py` |

---

# 工作流一：文档修订工作流

## 业务场景

在大模型应用（如Dify工作流）中，自动化处理企业文档的检查与修订：
1. **文档检查** - 发现格式、内容、语言、逻辑问题
2. **系统信息关联** - 从Excel系统信息表获取系统正式名称、开发/运维部门等基础信息
3. **文档与系统匹配** - 分析文档中涉及的IT系统，匹配系统信息表
4. **生成建议** - 基于检查规则给出修改建议
5. **自动修订** - 根据建议自动修改文档，支持三种修订模式
6. **复核确认** - 验证修订结果

## 整体架构（8步工作流）

```
                              ┌─────────────────┐
                              │ 1.5a MCP节点     │
                              │ (Excel读取系统信息)│
                              └────────┬────────┘
                                       │ 系统信息表JSON
                                       ▼
┌─────────────┐              ┌─────────────────┐
│ 1. 开始节点  │              │ 1.5b LLM节点     │
│  (读入文档)  │──文档内容───▶│(分析文档关联系统) │
└─────────────┘              └────────┬────────┘
                                       │ 匹配结果 + 系统信息
                                       ▼
                              ┌─────────────────┐
                              │ 2. LLM节点       │
                              │ (硬性规则检查)    │
                              └────────┬────────┘
                                       │
                              ┌────────▼────────┐
                              │ 3. LLM节点       │
                              │ (软性规则检查)    │
                              └────────┬────────┘
                                       │
┌─────────────┐     ┌─────────────┐    │
│ 7. LLM节点  │◀────│ 6. MCP节点  │◀───┤
│  (复核文档)  │     │  (文档修订)  │    │
└─────────────┘     └─────────────┘    │
                             ▲         │
                             │    ┌────┴────────────┐
                             │    │ 4. LLM节点       │
                             │    │(建议合并去重排序) │
                             │    └────────┬────────┘
                             │             │
                      ┌──────┴──────┐      │
                      │ 5. 用户输入  │◀─────┘
                      │(选择采纳建议)│
                      └─────────────┘
```

### 各节点功能

| 节点 | 功能 |
|------|------|
| 1. MCP节点(文档读取) | 读入文档内容，支持DOCX/WPS/WPSX |
| **1.5a. MCP节点(系统信息读取)** | **从Excel系统信息表中读取系统正式名称、研发/运维牵头部门等** |
| **1.5b. LLM节点(系统关联分析)** | **分析文档中涉及的IT系统，与系统信息表匹配，输出结构化系统关联信息** |
| 2. LLM节点(硬性规则) | 根据强制规则（格式、编号、命名规范等）检查并输出建议 |
| 3. LLM节点(软性规则) | 根据建议规则（内容、逻辑、语言）检查并输出建议 |
| 4. LLM节点(合并去重) | 整合前两步建议，合并同一原文的修改，标记严重程度 |
| 5. 用户输入 | 用户决定采纳哪些建议 |
| 6. **MCP节点(文档修订)** | 根据建议自动修订文档，支持三种模式 |
| 7. LLM节点(复核) | 对修订后的文档整体复核 |

## 核心能力

| 功能 | 说明 |
|------|------|
| **文档解析** | 支持DOCX、WPS、WPSX格式 |
| **系统信息读取** | 从Excel(.xlsx)系统信息表中读取系统正式名称、研发牵头部门、运维牵头部门等，支持表头模糊匹配 |
| **文档-系统关联分析** | 结合LLM从文档中识别涉及的IT系统，与系统信息表进行精确/模糊匹配 |
| **修订应用** | 根据LLM建议自动替换文本，支持三种修订模式 |
| **修订追踪** | Word Track Changes（w:del/w:ins）、红色括号批注、直接替换 |
| **修改记录明细** | 自动生成修改记录DOCX文档，记录每次修改明细，支持追加 |
| **格式兼容** | 生成的文档可被WPS/Word正常打开 |

## 修订模式技术实现

支持三种修订模式：

### 模式一：Word Track Changes（use_track_changes="true"）

使用OOXML标准标记，生成真正的Word修订痕迹：

```xml
<w:del w:id="1" w:author="DocumentReviewer">
  <w:r><w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr><w:t>原文本</w:t></w:r>
</w:del>
<w:ins w:id="1" w:author="DocumentReviewer">
  <w:r><w:rPr><w:u w:val="single"/><w:color w:val="0000FF"/></w:rPr><w:t>新文本</w:t></w:r>
</w:ins>
```

- 删除文本显示为红色删除线，插入文本显示为蓝色下划线
- 可在WPS/Word"审阅"选项卡中逐条接受/拒绝修订

### 模式二：红色括号批注（use_track_changes="comment"）默认模式

保留原文，在文本后插入红色括号批注（保留原文所有字体属性，仅颜色改为红色）：

```
原文内容（【建议】修改后内容 【原因】修改原因说明）
```

- 原文完整保留，不会被删除
- 批注以红色字体显示，字体大小/名称与原文一致
- 适合需要保留原文供人工参考的场景

### 模式三：直接替换（use_track_changes="false"）

直接替换原文，不保留任何修订痕迹。适合已经确认无误的最终修订。

## 修改记录明细文档

每次修订完成后，自动生成 **`原文件名_revision_log.docx`**：

```
文件修改记录明细
被修改文件：xxx.docx

第1次修改
修改时间：2026-06-22 15:30:00
修改概述：共修订2条内容，成功应用2条；严重程度分布：High: 1条...

修改内容明细：
1. [✅] ref_001
   规则：缺少风险分析（CONTENT_023）
   类型：content  严重程度：High
   原文：项目进展顺利，暂无风险。
   修改后：当前主要风险：人员风险和技术风险。
   原因：风险分析过于简单

──────────────────────────────────────────────────
第2次修改                 ← 后续修改自动追加
...
```

**功能特点**：
- 每次修订成功后自动生成，修订结果中包含 `revision_log_path` 字段
- 后续修改同一文档时自动追加，次数递增
- 已应用标记 ✅，未找到原文标记 ❌
- 修改后内容深绿色，修改原因灰色

## 数据流转

### 输入（建议JSON格式）
```json
[
  {
    "id": "ref_001",
    "rule_id": "CONTENT_023",
    "rule_name": "缺少风险分析",
    "type": "content",
    "severity": "High",
    "original_text": "项目进展顺利，暂无风险。",
    "suggestion": "当前主要风险：1）人员风险；2）技术风险。",
    "reason": "风险分析过于简单"
  }
]
```

### 输出
```json
{
  "success": true,
  "output_path": "文档_revised.docx",
  "total_suggestions": 1,
  "applied_revisions": 1,
  "revision_log_path": "文档_revision_log.docx",
  "revisions_detail": [
    {
      "id": "ref_001", "rule_id": "CONTENT_023",
      "original": "项目进展顺利，暂无风险。",
      "new": "当前主要风险：1）人员风险；2）技术风险。",
      "status": "success"
    }
  ]
}
```

## Dify配置示例

**第1步 - 文档读取节点**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: read_document 或 read_document_segmented
输入:
  file_path: {{ start.file_path }}
```

**第1.5步a - 系统信息读取节点（新增）**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: read_system_info
输入:
  file_path: {{ system_info_excel_path }}    # 系统信息Excel文件路径
输出变量:
  system_info_table: {{ system_info_table }}  # 系统信息表JSON
```

**第1.5步b - 文档系统关联分析节点（新增）**：
```yaml
节点类型: LLM
提示词: 使用 docs/prompt_step1_5_analyze_systems.md
输入变量:
  document_content: {{ step1.document_content }}
  system_info_table: {{ step1_5a.system_info_table }}
输出变量:
  system_match_result: {{ system_match_result }}  # 匹配结果JSON，传递给第2步
```

**第6步 - 文档修订节点**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: revise_document
输入:
  file_path: {{ start.file_path }}
  suggestions_json: {{ step5.selected_suggestions }}
  use_track_changes: "comment"  # true=Track Changes, comment=红色批注(默认), false=直接替换
输出变量:
  output_path: {{ output_path }}
  applied_revisions: {{ applied_revisions }}
```

> **注意**：系统信息表Excel文件的格式必须包含以下列（表头支持模糊匹配）：
> - 系统中文名称、系统英文简称、模块中文名称、模块英文简称、研发牵头部门、运维牵头部门

---

# 工作流二：时间轴图生成工作流

## 业务场景

根据事件分析报告中的处置过程描述，自动生成专业美观的时间轴图（PNG格式），适用于：
- **安全事件应急响应** - 展示各处置步骤的时间线
- **运维故障排查** - 记录问题发现到解决的完整过程
- **项目管理** - 展示项目里程碑和关键节点
- **审计合规** - 记录操作行为和决策时间点

## 工作流架构（4步工作流）

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│ 1. MCP节点  │────▶│ 2. LLM节点  │────▶│ 3. LLM节点  │────▶│ 4. MCP节点  │
│  (读入报告)  │     │ (分析处置过程)│     │ (校验时间顺序)│     │(生成时间轴图)│
└─────────────┘     └─────────────┘     └─────────────┘     └─────────────┘
```

### 各节点功能

| 节点 | 功能 | 说明 |
|------|------|------|
| 1. MCP节点(读取报告) | 读取事件分析报告的DOCX文档内容 | 复用文档修订工作流的 `read_document` 工具 |
| 2. LLM节点(分析处置过程) | 从报告中提取应急处置过程，输出结构化JSON | 按时间顺序排列事件，标记节点类型 |
| 3. LLM节点(校验) | 检查第2步输出的完整性和时间顺序 | 确保时间无矛盾、事件无遗漏 |
| 4. MCP节点(生成时间轴图) | 根据结构化数据生成PNG时间轴图 | 输出专业美观的可视化图片 |

## 核心能力

| 功能 | 说明 |
|------|------|
| **横向递进布局** | 时间轴水平从左到右，符合阅读习惯 |
| **蛇形多行排布** | 超过10个时间点自动换行，支持海量事件 |
| **同一时间多事件** | 支持 `descriptions` 数组，同一时刻多处置项 |
| **节点类型分类** | start(绿色)/end(红色)/special(橙色菱形)/normal(蓝色) |
| **按处理人着色** | 自动按不同处理人/系统分配颜色 |
| **内容块在上方** | 所有描述内容统一在时间轴线上方，避免干扰 |
| **箭头连接** | 主节点到描述块带箭头虚线，行间虚线连接 |

## 输入数据格式

第2步LLM节点需要输出以下JSON格式：

```json
{
  "title": "核心交易系统生产事件处置时间轴",
  "incident_id": "INC-2024-001",
  "start_time": "2024-01-15 14:00",
  "end_time": "2024-01-15 17:10",
  "events": [
    {
      "time": "14:00",
      "descriptions": ["Zabbix监控告警触发", "核心交易系统响应超时"],
      "handler": "监控系统",
      "system": "Zabbix",
      "node_type": "start"
    },
    {
      "time": "14:10",
      "descriptions": ["通知技术经理", "通知DBA团队"],
      "handler": "张三",
      "system": "电话通知",
      "node_type": "normal"
    },
    {
      "time": "14:20",
      "descriptions": ["启动应急响应流程", "成立应急指挥小组"],
      "handler": "李四",
      "system": "应急指挥平台",
      "node_type": "special"
    },
    {
      "time": "17:00",
      "descriptions": ["事件关闭"],
      "handler": "监控系统",
      "system": "Zabbix",
      "node_type": "end"
    }
  ]
}
```

### 字段说明

| 字段 | 必填 | 说明 |
|------|------|------|
| `title` | 是 | 时间轴标题 |
| `incident_id` | 否 | 事件编号 |
| `events[].time` | 是 | 事件发生时间 |
| `events[].descriptions` | 是 | 处置内容数组（支持多个） |
| `events[].description` | 兼容 | 单个处置内容（兼容旧格式） |
| `events[].handler` | 否 | 处理人/系统（用于着色） |
| `events[].system` | 否 | 关联系统名称 |
| `events[].node_type` | 否 | 节点类型：start/end/special/normal |

### 节点类型说明

| 类型 | 标记 | 样式 |
|------|------|------|
| `start` | 开始节点 | 绿色圆形 + 双圈效果 + START标签 |
| `end` | 结束节点 | 红色圆形 + 双圈效果 + END标签 |
| `special` | 特殊节点 | 橙色菱形 + 暖色背景框（关键决策/异常） |
| `normal` | 普通节点 | 蓝色圆形 + 白色半透明背景框 |

## 输出

```json
{
  "success": true,
  "output_path": "timeline.png",
  "file_size_readable": "267.1 KB",
  "event_count": 30,
  "total_event_items": 30,
  "image_format": "PNG"
}
```

时间轴图示例布局（30节点蛇形排布）：

```
 第1行 [09:00]─[09:30]─...─[13:30]    ← 从左到右，内容块在上方
                         │
                         ╲ 虚线连接
                         │
 第2行 [14:00]─[14:30]─...─[18:30]    ← 从右到左（蛇形反转），内容块在上方
                         │
                         ╲ 虚线连接
                         │
 第3行 [19:00]─[19:30]─...─[23:30]    ← 从左到右，内容块在上方
```

## Dify配置示例

**第1步 - 读取报告**：
```yaml
节点类型: MCP
服务: unified_mcp_server  # 或 timeline_mcp_server
工具: read_document
输入:
  file_path: {{ start.file_path }}
输出变量:
  document_content: {{ document_content }}
```

**第2步 - LLM分析处置过程**：
```yaml
节点类型: LLM
提示词: 从事件分析报告中提取应急处置过程，按时间顺序排列。
         每条事件包含 time(时间)、descriptions(处置内容数组)、handler(处理人)、node_type(节点类型)
         输出JSON格式，参考 get_timeline_template 工具返回的模板
```

**第3步 - LLM校验**：
```yaml
节点类型: LLM
提示词: 检查第2步输出：1)时间顺序是否递增 2)是否有遗漏的关键事件 3)node_type是否合理
```

**第4步 - 生成时间轴图**：
```yaml
节点类型: MCP
服务: timeline_mcp_server   # 端口8002（http://40.129.21.85:8002）
工具: generate_timeline
输入:
  data_json: {{ step2.output }}
输出变量:
  output_path: {{ output_path }}
```

---

# 通用部分

## MCP服务架构说明

本项目包含多个独立的MCP服务，它们之间的关系如下：

```
┌─────────────────────────────────────────────────────────────────┐
│                      unified_mcp_server.py                      │
│                     MCP实例: "UnifiedDocumentServer"             │
│                     默认端口: 18080 (SSE模式)                     │
│  ┌─────────────────────┐  ┌──────────────────────────────┐      │
│  │  @mcp.tool() 工具   │  │  @mcp.tool() 工具            │      │
│  │  - read_document    │  │  - read_system_info          │      │
│  │  - get_document_info│  │  - match_system_from_document│      │
│  └─────────────────────┘  └──────────────────────────────┘      │
│  ┌─────────────────────┐  ┌──────────────────────────────┐      │
│  │  @mcp.tool() 工具   │  │  @mcp.tool() 工具            │      │
│  │  - revise_document  │  │  - validate_suggestions      │      │
│  │  - get_tools_info   │  │                              │      │
│  └─────────────────────┘  └──────────────────────────────┘      │
└─────────────────────────────────────────────────────────────────┘
                          │ 启动时自动拉起子进程
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│                     timeline_mcp_server.py                      │
│                     MCP实例: "TimelineDocumentServer"            │
│                     默认端口: 8002 (SSE模式)                     │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │  @mcp.tool() 工具                                      │    │
│  │  - read_document    ← 代理到 unified_mcp_server        │    │
│  │  - generate_timeline                                    │    │
│  │  - validate_timeline_data                               │    │
│  │  - get_timeline_template                                │    │
│  └─────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────┘
```

### 多服务不冲突的原因

两个源文件中都有 `read_document` 函数且都带有 `@mcp.tool()` 装饰器，但**不会冲突**，原因如下：

| 比较项 | unified_mcp_server.py | timeline_mcp_server.py |
|--------|----------------------|----------------------|
| **MCP实例** | `FastMCP("UnifiedDocumentServer")` | `FastMCP("TimelineDocumentServer")` |
| **进程** | 主进程（独立内存空间） | 子进程（独立内存空间） |
| **默认端口** | 18080 | 8002 |
| **`read_document`** | 独立的完整实现 | 代理调用 unifed 中的 `DocumentReader` |

**核心原理**：两个服务运行在**不同的操作系统进程**中，各自拥有独立的 `mcp` 对象和内存空间。`@mcp.tool()` 注册的工具只属于各自的 `FastMCP` 实例，互不干扰。就像两台电脑上各有一个同名的程序，名字相同但运行环境完全隔离。

在 Dify 中配置 MCP 工具时，通过**服务地址+端口**区分：
- `http://40.129.21.85:18080` → 文档修订工作流（unified_mcp_server）
- `http://40.129.21.85:8002` → 时间轴图工作流（timeline_mcp_server）

### 服务启动关系

- **SSE模式**：启动 `unified_mcp_server.py` 会自动以子进程拉起 `timeline_mcp_server.py`（40.129.21.85:8002），两者同时运行
- **stdio模式**：`unified_mcp_server.py` 不会自动启动 timeline 服务，需手动启动

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

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. Python API直接调用

**文档修订**：
```python
from unified_mcp_server import DocumentReader, DocumentReviser
import json

# 读取文档
reader = DocumentReader()
result = reader.read_document("原文档.docx")

# 修订文档（comment模式为默认）
reviser = DocumentReviser()
suggestions = [{"id": "ref_001", "original_text": "原文", "suggestion": "修改后文本"}]
result = reviser.revise_document(
    file_path="原文档.docx",
    suggestions_json=json.dumps(suggestions),
    output_path="修订后.docx",
    use_track_changes="comment"  # true/comment/false
)
print(f"应用修订数: {result['applied_revisions']}")
print(f"修改记录: {result.get('revision_log_path')}")
```

**时间轴图生成**：
```python
from timeline_mcp_server import TimelineChartGenerator

data = {
    "title": "事件处置时间轴",
    "incident_id": "INC-001",
    "events": [
        {"time": "14:00", "descriptions": ["告警触发"], "handler": "监控系统", "node_type": "start"},
        {"time": "14:30", "descriptions": ["分析定位", "执行操作"], "handler": "张三", "node_type": "normal"},
        {"time": "15:00", "descriptions": ["业务验证通过"], "handler": "李四", "node_type": "special"},
        {"time": "16:00", "descriptions": ["事件关闭"], "handler": "监控系统", "node_type": "end"},
    ]
}
result = TimelineChartGenerator.generate_timeline(data, output_path="timeline.png")
print(f"时间轴图: {result['output_path']} ({result.get('file_size_readable')})")
```

### 3. 启动服务

```bash
# 统一MCP服务器（文档修订工作流 + 自动联动时间轴服务）
python unified_mcp_server.py                                       # SSE模式（默认，端口18080）
python unified_mcp_server.py --transport stdio                     # stdio模式

# 独立时间轴MCP服务器
python timeline_mcp_server.py                                      # SSE模式（默认，端口8002）
python timeline_mcp_server.py --transport stdio                    # stdio模式
```

### 4. 运行测试

```bash
cd tests
python test_all.py
```

## 核心文件

| 文件 | 说明 |
|------|------|
| `unified_mcp_server.py` | **统一MCP服务器** - 文档读取+系统信息读取+文档修订，联动启动时间轴服务 |
| `timeline_mcp_server.py` | **时间轴图MCP服务器** - 根据事件数据生成PNG时间轴图 |
| `document_reader_mcp.py` | 独立文档读取MCP节点（支持智能分段） |
| `revision_mcp_node.py` | 旧版文档修订MCP节点（已不推荐） |
| `prompt_step1_5_analyze_systems.md` | 第1.5步文档系统关联分析提示词 |
| `prompt_step2_hard_rules.md` | 第2步硬性规则检查提示词 |
| `prompt_step3_soft_rules_v2.md` | 第3步软性规则检查提示词 |
| `prompt_step4_merge_prioritize.md` | 第4步合并去重排序提示词 |
| `prompt_step5_user_confirmation.md` | 第5步用户确认提示词 |
| `测试数据_系统模块信息.xlsx` | 系统信息Excel样例，包含6个系统/20条记录的测试数据 |
| `tests/` | 测试文件夹 |

## 代码结构

### unified_mcp_server.py

```
unified_mcp_server.py
├── DocumentMetadata         # 文档元数据模型
├── CheckSuggestion          # 建议数据模型
├── Severity                 # 严重程度枚举
├── DocumentReader           # 文档读取器
│   ├── read_document()      # 读取文档主入口
│   ├── _read_docx()         # Word文档解析
│   ├── _read_wps()          # WPS文档解析
│   └── _extract_docx_metadata()
├── DocumentReviser          # 文档修订器
│   ├── revise_document()    # 修订文档主入口
│   ├── _revise_docx()       # Word文档修订
│   ├── _revise_wps()        # WPS文档修订
│   ├── _apply_track_changes_to_paragraph()
│   ├── _create_track_changes_element()
│   ├── _add_red_comment_to_paragraph()
│   ├── _write_revision_log()
│   └── _get_revision_log_path()
├── SystemInfoReader          # 系统信息读取器（新增）
│   ├── read_system_info()   # 从Excel读取系统信息
│   └── _map_headers()       # 表头模糊匹配
└── MCP工具 (7个)
    ├── read_document()
    ├── get_document_info()
    ├── read_system_info()           # 新增：读取Excel系统信息表
    ├── match_system_from_document() # 新增：文档系统关键词匹配
    ├── revise_document()
    ├── validate_suggestions()
    └── get_tools_info()
```

### timeline_mcp_server.py

```
timeline_mcp_server.py
├── TimelineChartGenerator     # 时间轴图生成器
│   ├── validate_events_data()
│   ├── generate_timeline()
│   ├── generate_timeline_v2()
│   └── get_timeline_data_template()
└── MCP工具 (5个)
    ├── read_document()
    ├── generate_timeline()
    ├── validate_timeline_data()
    ├── get_timeline_template()
    └── get_tools_info()
```

## LLM节点 Temperature 设置

| 工作流 | 步骤 | Temperature | 说明 |
|--------|------|-------------|------|
| **文档修订** 第1.5步 系统关联分析 | 文档系统匹配 | 0.1-0.2 | 严格匹配 |
| **文档修订** 第2步 硬性规则检查 | 格式规范检查 | 0.2-0.3 | 精确执行 |
| **文档修订** 第3步 软性规则检查 | 内容质量检查 | 0.2-0.4 | 适当灵活 |
| **文档修订** 第4步 合并去重 | 结果整合 | 0.3-0.5 | 判断整合 |
| **文档修订** 第7步 复核 | 质量评估 | 0.3-0.5 | 综合判断 |
| **时间轴** 第2步 分析处置过程 | 事件提取 | 0.2-0.3 | 精确提取 |
| **时间轴** 第3步 校验 | 时间顺序检查 | 0.1-0.2 | 严格校验 |

## 注意事项

1. **WPS文档支持**：.wps/.wpsx 与 .docx 使用相同OOXML格式，可被WPS和Word同时识别
2. **编码问题**：中文环境注意Windows终端编码设置（建议 UTF-8）
3. **修订前备份**：建议保留原文档
4. **系统信息表格式**：Excel文件必须包含系统中文名称、系统英文简称、研发牵头部门、运维牵头部门等关键列，表头支持模糊匹配（如"系统中文名"→"系统中文名称"）
5. **性能说明**：系统信息表支持1100个以上系统，单次读取预估约0.2秒，关键词匹配约0.1秒，无性能瓶颈

## 应用场景

- **企业文档自动化检查** - 合同、报告、规范文档的智能审查与修订
- **安全事件应急响应** - 读取事件报告 → 分析处置过程 → 生成时间轴可视化
- **AI辅助写作** - 自动修正语法、逻辑、格式问题
- **合规性审查** - 确保文档符合企业规范和标准
- **运维故障复盘** - 记录问题发现到解决的完整时间线
