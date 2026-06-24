"""
统一MCP服务器入口（单端口模式）

在同一个进程中加载所有MCP工具的注册，通过全局唯一的 FastMCP 实例
（mcp_instance.py）在同一端口下暴露所有工具。

使用方法:
    python main.py [--transport stdio|sse] [--host 0.0.0.0] [--port 18080]

端点(SSE模式):
    GET /sse - SSE连接端点
    POST /messages - 发送消息
    GET /health - 健康检查
    GET /tools - 获取工具定义
"""

import sys
import argparse
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger('UnifiedEntry')

# ============================================================
# 导入所有工具模块（触发 @mcp.tool() 注册到共享实例）
# ============================================================

# 1. 导入共享MCP实例
try:
    from mcp_instance import mcp
except ImportError:
    print("错误: mcp_instance.py 未找到")
    print("请确保 mcp_instance.py 存在")
    sys.exit(1)

# 2. 导入文档修订工作流工具（第1步、第1.5步、第6步）
try:
    import unified_mcp_server
    logger.info("已加载文档修订工作流工具（read_document, get_document_info, read_system_info, match_system_from_document, revise_document, validate_suggestions）")
except ImportError as e:
    logger.error(f"加载 unified_mcp_server 失败: {e}")
    sys.exit(1)

# 3. 导入时间轴图工作流工具（第4步）
try:
    import timeline_mcp_server
    logger.info("已加载时间轴图工作流工具（read_timeline_document, generate_timeline, validate_timeline_data, get_timeline_template）")
except ImportError as e:
    logger.error(f"加载 timeline_mcp_server 失败: {e}")
    sys.exit(1)

# 4. 导入事件复盘分析工作流工具（第1步、第3步）
try:
    import incident_mcp_server
    logger.info("已加载事件复盘分析工作流工具（read_incident_events, generate_incident_report, get_incident_report_template）")
except ImportError as e:
    logger.error(f"加载 incident_mcp_server 失败: {e}")
    sys.exit(1)

# 5. 导入依赖检查
from unified_mcp_server import DOCX_AVAILABLE, OPENPYXL_AVAILABLE
from timeline_mcp_server import MATPLOTLIB_AVAILABLE


# ============================================================
# 主入口
# ============================================================

def main():
    """主入口函数"""
    parser = argparse.ArgumentParser(description='统一MCP服务器（单端口模式）')
    parser.add_argument(
        '--transport',
        choices=['stdio', 'sse'],
        default='sse',
        help='传输方式: stdio 或 sse (默认)'
    )
    parser.add_argument(
        '--host',
        default='40.129.21.85',
        help='SSE模式下的主机地址 (默认: 40.129.21.85)'
    )
    parser.add_argument(
        '--port',
        type=int,
        default=18080,
        help='SSE模式下的端口号 (默认: 18080)'
    )

    args = parser.parse_args()

    print("=" * 70)
    print("统一MCP服务器（单端口聚合模式）")
    print("包含：文档修订工作流 + 时间轴图工作流 + 事件复盘分析工作流")
    print("=" * 70)

    # 依赖检查
    if not DOCX_AVAILABLE:
        print("\n错误: python-docx未安装")
        print("请运行: pip install python-docx")
        sys.exit(1)

    if not OPENPYXL_AVAILABLE:
        print("\n警告: openpyxl未安装，系统信息读取功能不可用")
        print("如需使用系统信息读取功能，请运行: pip install openpyxl")

    if not MATPLOTLIB_AVAILABLE:
        print("\n警告: matplotlib未安装，时间轴图生成功能不可用")
        print("请运行: pip install matplotlib")

    print(f"\n传输方式: {args.transport.upper()}")

    if args.transport == 'sse':
        print(f"服务地址: http://{args.host}:{args.port}")
        print(f"SSE端点: http://{args.host}:{args.port}/sse")
        print(f"消息端点: http://{args.host}:{args.port}/messages")

    print("\n可用工具（全部在同一端口下）:")
    print("  【第1步 - 文档读取】")
    print("    1. read_document              - 读取文档内容")
    print("    2. get_document_info          - 获取文档基本信息")
    print("  【第1.5步 - 系统信息获取】")
    print("    3. read_system_info           - 从Excel系统信息表读取系统基础信息")
    print("    4. match_system_from_document - 从文档文本中匹配关联的系统信息")
    print("  【第6步 - 文档修订】")
    print("    5. revise_document            - 根据建议修订文档")
    print("    6. validate_suggestions       - 验证建议JSON")
    print("  【第4步 - 时间轴图生成】")
    print("    7. read_timeline_document     - 读取时间轴源文档")
    print("    8. generate_timeline          - 根据事件数据生成时间轴PNG图")
    print("    9. validate_timeline_data     - 验证时间轴数据格式")
    print("   10. get_timeline_template      - 获取时间轴数据模板")
    print("   11. get_timeline_tools_info    - 时间轴工具信息")
    print("  【事件复盘分析 - 生产事件复盘】")
    print("   12. read_incident_events       - 第1步：读取事件情况表")
    print("   13. generate_incident_report   - 第3步：生成复盘报告")
    print("   14. get_incident_report_template - 获取报告模板")
    print("\n启动服务...")
    print("-" * 70)

    # 启动FastMCP服务
    if args.transport == 'sse':
        mcp.run(transport='sse', host=args.host, port=args.port)
    else:
        mcp.run()


if __name__ == "__main__":
    main()
