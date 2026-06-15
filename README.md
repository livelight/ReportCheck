# 文档修订MCP节点 - 第6步实现

## 功能概述

本模块实现文档检查工作流的第6步：**MCP节点 - 调用MCP服务，输入为json字符串格式的检查建议、原文档，已修订模式对文档进行修订、并生成修订后的新文档。**

## 项目背景与目标

### 业务场景
在大模型应用（如Dify工作流）中，需要自动化处理企业文档，包括：
1. **文档检查** - 发现格式、内容、语言、逻辑问题
2. **生成建议** - 基于检查规则给出修改建议
3. **自动修订** - 根据建议自动修改文档
4. **复核确认** - 验证修订结果

### 项目定位
这是整个工作流的**第6步 - MCP节点（Model Context Protocol）**，负责**文档修订**。

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
| **修订应用** | 根据LLM建议自动替换文本 |
| **修订追踪** | 使用Word修订模式标记修改（删除线/下划线）|
| **格式兼容** | 生成的文档可被WPS/Word正常打开 |

### 技术特点

1. **FastMCP框架** - 使用`@mcp.tool()`注册工具，可被Dify等MCP客户端调用
2. **修订模式** - 修改内容用红色删除线（删除）和蓝色下划线（插入）标记
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

## 核心文件

| 文件 | 说明 |
|------|------|
| `revision_mcp_node.py` | 核心MCP节点实现（1600+行，包含解析器、修订器、MCP工具、数据转换函数） |
| `prompt_step2_hard_rules.md` | 第2步硬性规则检查提示词 |
| `prompt_step3_soft_rules_v2.md` | 第3步软性规则检查提示词（去重优化版） |
| `prompt_step4_merge_prioritize.md` | 第4步合并去重排序提示词 |
| `prompt_step5_user_confirmation.md` | 第5步用户确认提示词 |
| `mcp_server.py` | HTTP API服务（Flask，同步方式） |
| `mcp_sse_server.py` | SSE服务器（FastAPI，流式方式，推荐） |
| `test_revision.py` | 基础功能测试脚本 |
| `test_step4_to_step6.py` | 第4步到第6步数据流转测试 |
| `test_fastmcp.py` | FastMCP框架测试 |
| `test_track_changes.py` | 修订模式功能测试 |
| `test_wps_track_changes.py` | WPS/WPSX文档测试 |
| `test_sse.py` | SSE服务器测试脚本 |
| `check_revisions.py` | 修订标记验证工具 |
| `PROJECT_CHECK_REPORT.md` | 项目检查报告 |
| `OPTIMIZATION_SUMMARY.md` | 优化总结报告 |
| `FINAL_CHECK_REPORT.md` | 最终检查报告 |

## 代码结构

```
revision_mcp_node.py
├── CheckSuggestion          # 建议数据模型
├── DocumentParser           # 文档解析器
│   ├── parse_document()     # 主入口
│   ├── _parse_docx()        # Word解析
│   └── _parse_wps()         # WPS解析（使用OOXML格式）
├── TrackChangesReviser      # 修订器（核心）
│   ├── apply_suggestions()  # 应用建议
│   ├── _create_deleted_text()   # 创建删除标记（红色删除线）
│   ├── _create_inserted_text()  # 创建插入标记（蓝色下划线）
│   ├── _apply_to_docx_track_changes()  # DOCX修订
│   └── _apply_to_wps()      # WPS修订
├── MCP工具 (4个)
│   ├── revise_document()      # 修订文档（适配第4步输出格式）
│   ├── parse_document()       # 解析文档
│   ├── validate_suggestions() # 验证建议JSON
│   └── get_tools_info()       # 工具信息
└── 数据转换工具函数
    ├── transform_step4_to_step6()      # 第4步输出转换为第6步输入
    └── prepare_suggestions_for_revision()  # 准备建议数据用于修订
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

### 3. 在Dify中配置HTTP请求节点

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
from revision_mcp_node import RevisionMCPNode
import json

node = RevisionMCPNode()
result = node.revise_document(
    file_path="原文档.docx",
    suggestions_json=json.dumps(suggestions)
)
print(result)
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

使用OOXML标准标记（所有格式通用）：

```xml
<!-- 删除标记 - 红色删除线 -->
<w:del w:id="1" w:author="DocumentReviewer">
  <w:r>
    <w:rPr><w:strike/><w:color w:val="FF0000"/></w:rPr>
    <w:t>原文本</w:t>
  </w:r>
</w:del>

<!-- 插入标记 - 蓝色下划线 -->
<w:ins w:id="1" w:author="DocumentReviewer">
  <w:r>
    <w:rPr><w:u w:val="single"/><w:color w:val="0000FF"/></w:rPr>
    <w:t>新文本</w:t>
  </w:r>
</w:ins>
```

在WPS/Word中打开后：
- **删除的文本** → 显示为红色删除线
- **插入的文本** → 显示为蓝色下划线
- 可在"审阅"选项卡中逐条接受/拒绝修订

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

#### 方式一: stdio模式（MCP标准）

```bash
python revision_mcp_node.py
```

#### 方式二: SSE模式（推荐，适合Dify集成）

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

#### 方式三: HTTP API服务（Flask）

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
