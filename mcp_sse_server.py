"""
MCP SSE服务器 - 文档修订服务
使用FastAPI和SSE提供MCP协议支持

使用方法:
    python mcp_sse_server.py [--host 0.0.0.0] [--port 8000]

端点:
    GET /sse - SSE连接端点（用于接收服务器消息）
    POST /messages - 消息发送端点（用于向服务器发送请求）
    GET /health - 健康检查
    GET /tools - 获取工具定义
"""

import json
import asyncio
import argparse
from typing import Optional
from datetime import datetime
from contextlib import asynccontextmanager

# FastAPI和SSE
from fastapi import FastAPI, Request, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse, StreamingResponse
from sse_starlette.sse import EventSourceResponse

# MCP相关
from revision_mcp_node import (
    RevisionMCPNode,
    DocumentParser,
    CheckSuggestion,
    transform_step4_to_step6,
    prepare_suggestions_for_revision
)

# 全局MCP节点实例
mcp_node: Optional[RevisionMCPNode] = None


@asynccontextmanager
async def lifespan(app: FastAPI):
    """应用生命周期管理"""
    global mcp_node
    # 启动时初始化
    mcp_node = RevisionMCPNode()
    print("=" * 60)
    print("MCP SSE服务器启动")
    print("=" * 60)
    print(f"服务名称: {mcp_node.name}")
    print(f"版本: {mcp_node.version}")
    print(f"描述: {mcp_node.description}")
    print("-" * 60)
    print("端点:")
    print("  GET  /sse      - SSE连接")
    print("  POST /messages - 发送消息")
    print("  GET  /health   - 健康检查")
    print("  GET  /tools    - 工具定义")
    print("=" * 60)
    yield
    # 关闭时清理
    print("\n服务器关闭")


app = FastAPI(
    title="文档修订MCP SSE服务器",
    description="基于SSE的文档修订MCP服务",
    version="1.0.0",
    lifespan=lifespan
)

# 存储客户端连接
clients: dict[str, asyncio.Queue] = {}


@app.get("/health")
async def health_check():
    """健康检查"""
    return {
        "status": "healthy",
        "service": mcp_node.name if mcp_node else "unknown",
        "version": mcp_node.version if mcp_node else "unknown",
        "timestamp": datetime.now().isoformat()
    }


@app.get("/tools")
async def get_tools():
    """获取工具定义"""
    if not mcp_node:
        raise HTTPException(status_code=503, detail="服务未初始化")
    return mcp_node.get_tool_definition()


@app.get("/sse")
async def sse_endpoint(request: Request):
    """
    SSE连接端点
    客户端连接后，服务器会通过SSE发送消息
    """
    client_id = f"client_{datetime.now().strftime('%Y%m%d%H%M%S')}_{id(request)}"
    queue: asyncio.Queue = asyncio.Queue()
    clients[client_id] = queue

    print(f"[SSE] 客户端连接: {client_id}")

    async def event_generator():
        try:
            # 发送初始连接成功消息
            yield {
                "event": "connected",
                "data": json.dumps({
                    "client_id": client_id,
                    "status": "connected",
                    "timestamp": datetime.now().isoformat()
                })
            }

            # 持续监听消息队列
            while True:
                try:
                    # 使用wait_for实现超时检查连接状态
                    message = await asyncio.wait_for(
                        queue.get(),
                        timeout=30.0
                    )
                    yield {
                        "event": message.get("event", "message"),
                        "data": json.dumps(message.get("data", {}))
                    }
                except asyncio.TimeoutError:
                    # 发送心跳保持连接
                    yield {
                        "event": "heartbeat",
                        "data": json.dumps({"timestamp": datetime.now().isoformat()})
                    }

        except asyncio.CancelledError:
            print(f"[SSE] 客户端断开: {client_id}")
        finally:
            if client_id in clients:
                del clients[client_id]
                print(f"[SSE] 客户端清理: {client_id}")

    return EventSourceResponse(event_generator())


@app.post("/messages")
async def post_message(request: Request):
    """
    接收客户端消息并处理

    请求格式:
    {
        "client_id": "客户端ID",
        "tool": "工具名称",
        "params": {参数对象}
    }
    """
    if not mcp_node:
        raise HTTPException(status_code=503, detail="服务未初始化")

    try:
        body = await request.json()
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="无效的JSON")

    client_id = body.get("client_id")
    tool = body.get("tool")
    params = body.get("params", {})

    if not client_id:
        raise HTTPException(status_code=400, detail="缺少client_id")

    if not tool:
        raise HTTPException(status_code=400, detail="缺少tool参数")

    # 异步处理工具调用
    asyncio.create_task(process_tool_call(client_id, tool, params))

    return {"status": "accepted", "client_id": client_id, "tool": tool}


async def process_tool_call(client_id: str, tool: str, params: dict):
    """
    异步处理工具调用并发送结果到SSE
    """
    queue = clients.get(client_id)
    if not queue:
        print(f"[ERROR] 客户端不存在: {client_id}")
        return

    try:
        result = {"success": False, "error": "未知工具"}

        if tool == "revise_document":
            result = await async_revise_document(params)
        elif tool == "parse_document":
            result = await async_parse_document(params)
        elif tool == "validate_suggestions":
            result = await async_validate_suggestions(params)
        elif tool == "get_tools_info":
            result = await async_get_tools_info()
        else:
            result = {"success": False, "error": f"未知工具: {tool}"}

        # 发送结果到客户端
        await queue.put({
            "event": "tool_result",
            "data": {
                "tool": tool,
                "result": result,
                "timestamp": datetime.now().isoformat()
            }
        })

    except Exception as e:
        print(f"[ERROR] 工具调用失败: {e}")
        await queue.put({
            "event": "error",
            "data": {
                "tool": tool,
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
        })


async def async_revise_document(params: dict) -> dict:
    """异步修订文档"""
    loop = asyncio.get_event_loop()

    file_path = params.get("file_path")
    suggestions_json = params.get("suggestions_json", "[]")
    output_path = params.get("output_path", "")
    use_track_changes = params.get("use_track_changes", "true")

    # 在线程池中执行同步代码
    result = await loop.run_in_executor(
        None,
        lambda: mcp_node.revise_document(
            file_path=file_path,
            suggestions_json=suggestions_json,
            output_path=output_path,
            use_track_changes=use_track_changes
        )
    )

    return result


async def async_parse_document(params: dict) -> dict:
    """异步解析文档"""
    loop = asyncio.get_event_loop()

    file_path = params.get("file_path")

    result = await loop.run_in_executor(
        None,
        lambda: DocumentParser.parse_document(file_path)
    )

    return {
        "success": True,
        "file_path": file_path,
        "type": result.get("type"),
        "paragraph_count": len(result.get("paragraphs", [])),
        "table_count": len(result.get("tables", [])),
        "paragraphs": [p.get("text", "") for p in result.get("paragraphs", [])]
    }


async def async_validate_suggestions(params: dict) -> dict:
    """异步验证建议"""
    suggestions_json = params.get("suggestions_json", "[]")

    try:
        data = json.loads(suggestions_json)

        if isinstance(data, dict):
            suggestions = data.get("suggestions", [])
        elif isinstance(data, list):
            suggestions = data
        else:
            return {"success": False, "error": "无效的格式"}

        return {
            "success": True,
            "valid": True,
            "count": len(suggestions),
            "message": f"验证通过，共{len(suggestions)}条建议"
        }

    except json.JSONDecodeError as e:
        return {
            "success": False,
            "valid": False,
            "error": f"JSON解析错误: {str(e)}"
        }


async def async_get_tools_info() -> dict:
    """获取工具信息"""
    return mcp_node.get_tool_definition()


# ============ 同步HTTP API (保留兼容) ============

@app.post("/revise")
async def revise_endpoint(request: Request):
    """同步修订文档接口（HTTP JSON）"""
    if not mcp_node:
        raise HTTPException(status_code=503, detail="服务未初始化")

    body = await request.json()

    file_path = body.get("file_path")
    suggestions_json = body.get("suggestions_json")
    output_path = body.get("output_path", "")

    if not file_path or not suggestions_json:
        raise HTTPException(status_code=400, detail="缺少必需参数")

    loop = asyncio.get_event_loop()
    result = await loop.run_in_executor(
        None,
        lambda: mcp_node.revise_document(
            file_path=file_path,
            suggestions_json=suggestions_json,
            output_path=output_path
        )
    )

    return JSONResponse(content=result)


@app.post("/revise/stream")
async def revise_stream_endpoint(request: Request):
    """
    流式修订文档接口（SSE Stream）
    实时返回修订进度
    """
    if not mcp_node:
        raise HTTPException(status_code=503, detail="服务未初始化")

    body = await request.json()
    file_path = body.get("file_path")
    suggestions_json = body.get("suggestions_json")

    if not file_path or not suggestions_json:
        raise HTTPException(status_code=400, detail="缺少必需参数")

    async def progress_generator():
        """生成进度事件"""
        yield {
            "event": "start",
            "data": json.dumps({"message": "开始修订文档", "timestamp": datetime.now().isoformat()})
        }

        # 模拟进度更新
        for i, progress in enumerate([25, 50, 75, 100]):
            await asyncio.sleep(0.5)
            yield {
                "event": "progress",
                "data": json.dumps({
                    "progress": progress,
                    "message": f"修订进度: {progress}%",
                    "timestamp": datetime.now().isoformat()
                })
            }

        # 执行实际修订
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None,
            lambda: mcp_node.revise_document(
                file_path=file_path,
                suggestions_json=suggestions_json
            )
        )

        yield {
            "event": "complete",
            "data": json.dumps({
                "success": result.get("success", False),
                "output_path": result.get("output_path"),
                "applied_revisions": result.get("applied_revisions", 0),
                "timestamp": datetime.now().isoformat()
            })
        }

    return EventSourceResponse(progress_generator())


def main():
    """主入口"""
    parser = argparse.ArgumentParser(description='MCP SSE服务器')
    parser.add_argument('--host', default='0.0.0.0', help='主机地址')
    parser.add_argument('--port', type=int, default=8000, help='端口号')
    parser.add_argument('--reload', action='store_true', help='开发模式（热重载）')

    args = parser.parse_args()

    import uvicorn
    uvicorn.run(
        "mcp_sse_server:app",
        host=args.host,
        port=args.port,
        reload=args.reload,
        log_level="info"
    )


if __name__ == "__main__":
    main()
