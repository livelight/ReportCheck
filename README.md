# 文档检查与时间轴生成MCP服务

## 项目简介

本项目基于 MCP (Model Context Protocol) 协议和 FastMCP 框架，提供三个独立的工作流服务：

| 工作流 | 说明 | 核心文件 |
|--------|------|----------|
| **文档修订工作流** | 读取文档 → 获取系统信息 → 规则检查 → 合并建议 → 自动修订 → 复核 | `unified_mcp_server.py` |
| **时间轴图工作流** | 读取处置过程 → LLM分析 → 生成专业时间轴PNG图 | `timeline_mcp_server.py` |
| **事件复盘分析工作流** | 读取事件情况表 → LLM分析复盘 → 生成复盘报告Word文档 | `incident_mcp_server.py` |

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
服务: unified_mcp_server   # 单端口模式下使用同一个服务地址
工具: generate_timeline
输入:
  data_json: {{ step2.output }}
输出变量:
  output_path: {{ output_path }}
```


---

# 工作流三：事件复盘分析工作流

## 业务场景

在Dify工作流中，根据生产事件情况表自动进行复盘分析，生成结构化的复盘报告文档：

1. **读取事件数据** - 从Excel生产事件情况表中读取事件记录
2. **LLM复盘分析** - 由大模型分析事件数据，包括总结、原因分析、问题暴露、改进建议、应急处置评估
3. **生成报告** - 将LLM分析结果输出为排版精美的Word复盘报告

## 工作流架构（3步工作流）

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│ 1. MCP节点  │────▶│ 2. LLM节点  │────▶│ 3. MCP节点  │
│ (读取事件表) │     │ (复盘分析)   │     │ (生成报告)   │
└─────────────┘     └─────────────┘     └─────────────┘
```

### 各节点功能

| 节点 | 功能 | 说明 |
|------|------|------|
| 1. MCP节点(读取事件表) | 读取Excel事件情况表，输出结构化JSON | 支持模糊表头匹配，自动统计等级分布、影响系统等 |
| 2. LLM节点(复盘分析) | 根据事件数据进行深度复盘分析 | 按`get_incident_report_template`模板输出JSON |
| 3. MCP节点(生成报告) | 将LLM分析结果输出为Word文档 | 包含6大章节，自动排版，支持自定义输出路径 |

## 核心能力

| 功能 | 说明 |
|------|------|
| **事件表读取** | 支持.xlsx/.xls格式，14列标准字段，表头模糊匹配 |
| **多字段支持** | 事件编号、名称、时间、等级、影响系统、事件描述、处置方式、根因、监控告警、应急预案等 |
| **自动汇总统计** | 按等级分布、影响系统统计、总影响时长自动计算 |
| **结构化输出** | LLM分析输出严格遵循JSON模板，确保下游MCP节点正确解析 |
| **报告自动生成** | 生成包含事件总结、原因分析、问题暴露、改进建议、处置评估6大章节的Word报告 |
| **改进建议优先级** | 支持高/中/低三级，自动颜色标记（红/橙/绿） |
| **处置时间线** | 支持在报告中嵌入时间线，展示完整处置过程 |

## 输入数据格式

### 第1步输出（读取事件表）

```json
{
  "success": true,
  "file_path": "事件情况.xlsx",
  "total_events": 5,
  "summary": {
    "total_events": 5,
    "severity_distribution": {
      "P0重大": 1,
      "P1严重": 2,
      "P2一般": 2
    },
    "affected_systems": {
      "CRM系统": 1,
      "ERP系统": 1
    },
    "total_impact_minutes": 845
  },
  "events": [
    {
      "incident_id": "INC-2024-001",
      "incident_name": "核心数据库连接池耗尽导致服务不可用",
      "start_time": "2024-01-15 09:30",
      "end_time": "2024-01-15 11:15",
      "impact_duration_minutes": "105",
      "severity": "P0重大",
      "affected_systems": "CRM系统,ERP系统",
      "impact_scope": "全部客户无法登录及操作",
      "description": "因数据库连接数突增...",
      "emergency_response": "1.紧急重启数据库连接池；2.临时扩容...",
      "response_duration_minutes": "45",
      "root_cause": "某业务模块因代码缺陷...",
      "monitoring_alert_status": "未配置连接池使用率告警阈值...",
      "emergency_plan": "已制定数据库连接池扩容预案..."
    }
  ]
}
```

### 第2步LLM节点输出模板

参考 `get_incident_report_template` 工具，输出以下JSON结构：

```json
{
  "summary": {
    "overview": "一句话事件概述",
    "impact_duration": "影响时长描述",
    "severity": "事件等级",
    "involved_systems": "涉及系统列表",
    "impact_scope": "影响范围详细描述"
  },
  "cause_analysis": {
    "direct_cause": "直接原因",
    "indirect_cause": "间接原因",
    "root_cause": "根本原因分析"
  },
  "issues": {
    "monitoring_alerts": "监控告警方面的问题",
    "architecture_design": "系统架构设计方面的问题",
    "emergency_response": "应急处置和应急预案方面的问题",
    "others": "其他问题"
  },
  "suggestions": [
    {
      "category": "建议类别",
      "content": "具体改进建议",
      "priority": "高/中/低"
    }
  ],
  "response_info": {
    "response_method": "应急处置方式",
    "response_duration": "处置时长",
    "effectiveness": "处置效果评估",
    "timeline": [
      {"time": "时间点", "action": "处置动作"}
    ]
  }
}
```

### 第3步输出（生成报告）

```json
{
  "success": true,
  "output_path": "reports/事件复盘报告_20240624_183000.docx",
  "file_name": "事件复盘报告_20240624_183000.docx",
  "file_size": 38247,
  "report_sections": [
    "事件总体总结",
    "深入原因分析",
    "暴露出来的问题",
    "改进建议",
    "应急处置方式与处置时长"
  ]
}
```

## Dify配置示例

**第1步 - 读取事件情况表**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: read_incident_events
输入:
  file_path: {{ start.file_path }}    # 事件情况Excel文件路径
输出变量:
  events_data: {{ events_data }}      # 事件结构化数据，传递给第2步LLM节点
```

**第2步 - LLM复盘分析**：
```yaml
节点类型: LLM
提示词: |
  你是一个资深的生产事件复盘分析专家。根据以下事件数据，进行深入复盘分析。
  请严格按照 get_incident_report_template 工具返回的JSON模板格式输出。
  特别注意：
  - summary.overview 要简明扼要
  - cause_analysis 要深入分析技术和管理层面的根因
  - issues 要分四个方面系统梳理暴露的问题
  - suggestions 要给出可操作的具体改进建议，标注优先级(高/中/低)
  - response_info 要客观评估处置时效和效果

输入变量:
  events_data: {{ step1.events_data }}
输出变量:
  analysis_result: {{ analysis_result }}  # 结构化分析结果JSON
```

**第3步 - 生成复盘报告**：
```yaml
节点类型: MCP
服务: unified_mcp_server
工具: generate_incident_report
输入:
  analysis_result: {{ step2.analysis_result }}
  incident_file_name: {{ start.file_name }}   # 原始文件名（可选）
输出变量:
  report_path: {{ report_path }}             # 生成的Word报告路径
```

## 测试数据

提供 `事件情况_示例数据.xlsx` 作为测试数据，包含5条典型生产事件：

| 事件 | 等级 | 影响系统 | 处置时长 |
|------|------|----------|----------|
| 核心数据库连接池耗尽 | P0重大 | CRM,ERP | 45分钟 |
| 支付网关响应超时 | P1严重 | 支付系统 | 50分钟 |
| 文件存储磁盘写满 | P1严重 | OA系统 | 60分钟 |
| Session风暴登录异常 | P2一般 | IAM | 40分钟 |
| 数据中台任务延迟 | P2一般 | 数据中台 | 180分钟 |

测试方法：
```python
# 手动测试
python -c "
from incident_mcp_server import IncidentEventReader
import json
result = IncidentEventReader.read_events('事件情况_示例数据.xlsx')
print(json.dumps(result, ensure_ascii=False, indent=2))
"

# 测试报告生成
python -c "
from incident_mcp_server import IncidentReportGenerator
import json
analysis = {'summary': {'overview': '测试'}, ...}  # 构造分析数据
result = IncidentReportGenerator.generate_report(
    analysis_result=json.dumps(analysis),
    incident_file_name='事件情况_示例数据.xlsx'
)
print(result['output_path'])
"
```

---

# 通用部分
## MCP服务架构说明

本项目支持两种启动模式，可根据需要选择：

### 模式一：单端口聚合模式（推荐）

所有工具在同一个 FastMCP 实例中注册，通过一个端口对外提供服务。

```
                               mcp_instance.py
                         全局唯一 FastMCP 实例
                              "UnifiedDocumentServer"
                    ┌──────────────────────────────────┐
                    │  导入并注册 @mcp.tool()           │
                    └──────────────────────────────────┘
                               ▲        ▲         ▲
                导入共享实例    │        │         │ 导入共享实例
                    ┌──────────┘        └────┬─────┘
                    ▼                        ▼
      unified_mcp_server.py        incident_mcp_server.py
      ┌─────────────────────┐      ┌─────────────────────────┐
      │ 文档修订工作流工具   │      │ 事件复盘分析工作流工具   │
      │  - read_document    │      │  - read_incident_events │
      │  - get_document_info│      │  - generate_incident_rep│
      │  - read_system_info │      │  - get_incident_report_ │
      │  - match_system...  │      └─────────────────────────┘
      │  - revise_document  │
      │  - validate_suggest │
      └─────────────────────┘
                    │                          ▲
                    ▼                          │
         timeline_mcp_server.py ───────────────┘
      ┌─────────────────────────┐  (同样导入共享实例)
      │ 时间轴图工作流工具       │
      │  - read_timeline_document│
      │  - generate_timeline     │
      │  - validate_timeline_data│
      │  - get_timeline_template │
      └─────────────────────────┘
                    │
                    ▼
                 main.py (统一启动入口，默认端口18080)
                 通过 import 触发工具注册，单进程运行
```

**核心原理**：`mcp_instance.py` 创建全局唯一的 `FastMCP` 实例。所有工具模块（`unified_mcp_server.py`、`timeline_mcp_server.py`、`incident_mcp_server.py`）都导入此实例，通过 `@mcp.tool()` 注册各自的工具。`main.py` 导入所有模块（触发注册），然后调用 `mcp.run()` 启动服务。所有工具在**同一个进程、同一个端口**下可用。

Dify 中只需配置一个 MCP 服务地址：
- `http://40.129.21.85:18080` → 所有工具（文档修订 + 时间轴图 + 事件复盘）

### 模式二：独立双端口模式（向后兼容）

保留原始的多进程、多端口方案，用于需要独立部署的场景。

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
│  │  - get_document_... │  │                              │      │
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
│  │  - read_timeline_document                              │    │
│  │  - generate_timeline                                    │    │
│  │  - validate_timeline_data                               │    │
│  │  - get_timeline_template                                │    │
│  └─────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────┘
```

**核心原理**：两个服务运行在**不同的操作系统进程**中，各自拥有独立的 `FastMCP` 实例和内存空间。`unified_mcp_server.py` 启动时会自动以子进程拉起 `timeline_mcp_server.py`（端口8002）。

Dify 中需要配置两个 MCP 服务地址：
- `http://40.129.21.85:18080` → 文档修订工作流（unified_mcp_server）
- `http://40.129.21.85:8002` → 时间轴图工作流（timeline_mcp_server）

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

支持两种模式：

```bash
# 模式一（推荐）：单端口聚合模式 - 所有工具通过一个端口提供
python main.py                                                  # SSE模式（默认，端口18080）
python main.py --transport stdio                                # stdio模式
python main.py --port 18080                                     # 自定义端口

# 模式二（向后兼容）：独立部署模式
python unified_mcp_server.py                                    # 文档修订服务（端口18080）
python timeline_mcp_server.py                                   # 时间轴图服务（端口8002）
python incident_mcp_server.py                                   # 事件复盘分析服务（端口8003）
```

### 4. 运行测试

```bash
cd tests
python test_all.py
```

## 核心文件

| 文件 | 说明 |
|------|------|
| `unified_mcp_server.py` | **统一MCP服务器** - 文档读取+系统信息读取+文档修订（导入共享MCP实例） |
| `timeline_mcp_server.py` | **时间轴图MCP节点** - 根据事件数据生成PNG时间轴图（导入共享MCP实例） |
| `incident_mcp_server.py` | **事件复盘分析MCP节点** - 读取事件情况表+生成复盘报告（导入共享MCP实例） |
| `mcp_instance.py` | **全局唯一FastMCP实例** - 单端口模式下所有工具的共享MCP实例 |
| `main.py` | **统一启动入口** - 单端口聚合模式，导入所有工具模块并启动服务 |
| `document_reader_mcp.py` | 独立文档读取MCP节点（支持智能分段） |
| `revision_mcp_node.py` | 旧版文档修订MCP节点（已不推荐） |
| `prompt_step1_5_analyze_systems.md` | 第1.5步文档系统关联分析提示词 |
| `prompt_step2_hard_rules.md` | 第2步硬性规则检查提示词 |
| `prompt_step3_soft_rules_v2.md` | 第3步软性规则检查提示词 |
| `prompt_step4_merge_prioritize.md` | 第4步合并去重排序提示词 |
| `prompt_step5_user_confirmation.md` | 第5步用户确认提示词 |
| `测试数据_系统模块信息.xlsx` | 系统信息Excel样例，包含6个系统/20条记录的测试数据 |
| `事件情况_示例数据.xlsx` | 生产事件情况Excel样例，包含5条典型事件的测试数据 |
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
├── SystemInfoReader          # 系统信息读取器
│   ├── read_system_info()   # 从Excel读取系统信息
│   └── _map_headers()       # 表头模糊匹配
└── MCP工具 (7个，通过import mcp_instance注册)
    ├── read_document()
    ├── get_document_info()
    ├── read_system_info()
    ├── match_system_from_document()
    ├── revise_document()
    ├── validate_suggestions()
    └── get_document_tools_info()
```

### timeline_mcp_server.py

```
timeline_mcp_server.py
├── TimelineChartGenerator     # 时间轴图生成器
│   ├── validate_events_data()
│   ├── generate_timeline()
│   ├── generate_timeline_v2()
│   └── get_timeline_data_template()
└── MCP工具 (5个，通过import mcp_instance注册)
    ├── read_timeline_document()
    ├── generate_timeline()
    ├── validate_timeline_data()
    ├── get_timeline_template()
    └── get_timeline_tools_info()
```

### mcp_instance.py

```
mcp_instance.py
└── mcp = FastMCP("UnifiedDocumentServer")   # 全局唯一的FastMCP实例
```

### incident_mcp_server.py

```
incident_mcp_server.py
├── IncidentEventReader        # 事件情况表读取器
│   ├── read_events()          # 从Excel读取事件数据
│   └── _map_headers()         # 表头模糊匹配
├── IncidentReportGenerator    # 复盘报告生成器
│   ├── generate_report()      # 生成Word复盘报告
│   ├── _add_section_heading()
│   ├── _add_subsection_text()
│   └── _add_field_row()
└── MCP工具 (3个，通过import mcp_instance注册)
    ├── read_incident_events()
    ├── generate_incident_report()
    └── get_incident_report_template()

## LLM节点 Temperature 设置

| 工作流 | 步骤 | Temperature | 说明 |
|--------|------|-------------|------|
| **文档修订** 第1.5步 系统关联分析 | 文档系统匹配 | 0.1-0.2 | 严格匹配 |
| **文档修订** 第2步 硬性规则检查 | 格式规范检查 | 0.2-0.3 | 精确执行 |
| **文档修订** 第3步 软性规则检查 | 内容质量检查 | 0.2-0.4 | 适当灵活 |
| **文档修订** 第4步 合并去重 | 结果整合 | 0.3-0.5 | 判断整合 |
| **文档修订** 第7步 复核 | 质量评估 | 0.3-0.5 | 综合判断 |
| **时间轴** 第2步 分析处置过程 | 事件提取 | 0.2-0.3 | 精确提取 |
| **事件复盘** 第2步 复盘分析 | 事件总结与建议 | 0.2-0.4 | 综合分析判断 |

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
