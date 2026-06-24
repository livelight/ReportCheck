"""
全局唯一的 FastMCP 实例（单例模式）

所有工具模块通过导入此文件共享同一个 mcp 实例，
从而实现所有工具在同一个端口下提供服务。

使用方法：
    from mcp_instance import mcp

    @mcp.tool()
    def my_tool(...) -> str:
        ...
"""

from fastmcp import FastMCP

# 全局唯一的 MCP 实例
# 所有模块导入此实例并注册 @mcp.tool()
mcp = FastMCP("UnifiedDocumentServer")
