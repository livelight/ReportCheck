"""
生成30个节点的时间轴测试图
"""
import sys
import os

# 添加父目录到路径（timeline_mcp_server.py在父目录）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from timeline_mcp_server import TimelineChartGenerator
import json

def generate_30_node_timeline():
    """生成包含30个节点的时间轴图"""
    
    # 构建30个事件数据 - 模拟一个完整的安全事件处置过程
    data = {
        "title": "核心交易系统生产事件处置时间轴",
        "incident_id": "INC-2024-0622-001",
        "start_time": "2024-06-22 09:00",
        "end_time": "2024-06-22 18:30",
        "events": [
            # 第1行：上午时段 (10个节点)
            {
                "time": "09:00",
                "descriptions": ["Zabbix监控告警触发", "核心交易系统响应超时"],
                "handler": "监控系统",
                "system": "Zabbix",
                "node_type": "start"
            },
            {
                "time": "09:05",
                "descriptions": ["值班人员收到告警通知"],
                "handler": "张三",
                "system": "企业微信",
                "node_type": "normal"
            },
            {
                "time": "09:10",
                "descriptions": ["初步确认系统异常", "检查应用日志"],
                "handler": "张三",
                "system": "日志平台",
                "node_type": "normal"
            },
            {
                "time": "09:15",
                "descriptions": ["通知技术经理"],
                "handler": "张三",
                "system": "电话通知",
                "node_type": "normal"
            },
            {
                "time": "09:20",
                "descriptions": ["启动应急响应流程", "成立应急指挥小组"],
                "handler": "李四",
                "system": "应急指挥平台",
                "node_type": "special"
            },
            {
                "time": "09:30",
                "descriptions": ["DBA团队介入分析"],
                "handler": "王五",
                "system": "数据库监控",
                "node_type": "normal"
            },
            {
                "time": "09:40",
                "descriptions": ["发现数据库连接池耗尽"],
                "handler": "王五",
                "system": "Oracle DB",
                "node_type": "normal"
            },
            {
                "time": "09:50",
                "descriptions": ["尝试重启应用服务"],
                "handler": "赵六",
                "system": "Kubernetes",
                "node_type": "normal"
            },
            {
                "time": "10:00",
                "descriptions": ["重启后问题依然存在"],
                "handler": "赵六",
                "system": "应用监控",
                "node_type": "normal"
            },
            {
                "time": "10:10",
                "descriptions": ["升级事件等级为P1"],
                "handler": "李四",
                "system": "事件管理平台",
                "node_type": "special"
            },
            
            # 第2行：中午时段 (10个节点)
            {
                "time": "10:30",
                "descriptions": ["架构师团队加入排查"],
                "handler": "陈七",
                "system": "远程会议",
                "node_type": "normal"
            },
            {
                "time": "10:45",
                "descriptions": ["分析慢查询日志"],
                "handler": "王五",
                "system": "SQL分析工具",
                "node_type": "normal"
            },
            {
                "time": "11:00",
                "descriptions": ["定位到问题SQL语句"],
                "handler": "王五",
                "system": "Oracle AWR报告",
                "node_type": "normal"
            },
            {
                "time": "11:15",
                "descriptions": ["制定优化方案"],
                "handler": "陈七",
                "system": "技术方案评审",
                "node_type": "special"
            },
            {
                "time": "11:30",
                "descriptions": ["准备SQL优化脚本"],
                "handler": "王五",
                "system": "SQL Developer",
                "node_type": "normal"
            },
            {
                "time": "11:45",
                "descriptions": ["在测试环境验证优化效果"],
                "handler": "测试团队",
                "system": "测试环境",
                "node_type": "normal"
            },
            {
                "time": "12:00",
                "descriptions": ["测试验证通过", "性能提升80%"],
                "handler": "测试团队",
                "system": "性能测试报告",
                "node_type": "normal"
            },
            {
                "time": "12:15",
                "descriptions": ["提交变更申请"],
                "handler": "李四",
                "system": "变更管理系统",
                "node_type": "normal"
            },
            {
                "time": "12:30",
                "descriptions": ["变更审批通过"],
                "handler": "变更委员会",
                "system": "CAB审批",
                "node_type": "normal"
            },
            {
                "time": "12:45",
                "descriptions": ["准备生产环境执行计划"],
                "handler": "赵六",
                "system": "运维手册",
                "node_type": "normal"
            },
            
            # 第3行：下午时段 (10个节点)
            {
                "time": "13:00",
                "descriptions": ["开始执行生产变更"],
                "handler": "赵六",
                "system": "自动化运维平台",
                "node_type": "special"
            },
            {
                "time": "13:15",
                "descriptions": ["备份数据库"],
                "handler": "王五",
                "system": "Oracle RMAN",
                "node_type": "normal"
            },
            {
                "time": "13:30",
                "descriptions": ["执行SQL优化脚本"],
                "handler": "王五",
                "system": "SQL*Plus",
                "node_type": "normal"
            },
            {
                "time": "13:45",
                "descriptions": ["重建相关索引"],
                "handler": "王五",
                "system": "Oracle DB",
                "node_type": "normal"
            },
            {
                "time": "14:00",
                "descriptions": ["更新应用配置"],
                "handler": "赵六",
                "system": "配置中心",
                "node_type": "normal"
            },
            {
                "time": "14:15",
                "descriptions": ["重启应用服务"],
                "handler": "赵六",
                "system": "Kubernetes",
                "node_type": "normal"
            },
            {
                "time": "14:30",
                "descriptions": ["业务功能验证"],
                "handler": "业务团队",
                "system": "业务流程测试",
                "node_type": "normal"
            },
            {
                "time": "14:45",
                "descriptions": ["性能压测验证"],
                "handler": "测试团队",
                "system": "JMeter",
                "node_type": "normal"
            },
            {
                "time": "15:00",
                "descriptions": ["监控系统指标恢复正常"],
                "handler": "监控系统",
                "system": "Grafana",
                "node_type": "normal"
            },
            {
                "time": "15:15",
                "descriptions": ["持续观察30分钟"],
                "handler": "运维团队",
                "system": "监控大屏",
                "node_type": "normal"
            },
            
            # 第4行：收尾阶段 (最后几个节点)
            {
                "time": "15:45",
                "descriptions": ["确认系统稳定运行"],
                "handler": "李四",
                "system": "健康检查",
                "node_type": "normal"
            },
            {
                "time": "16:00",
                "descriptions": ["通知业务方恢复"],
                "handler": "李四",
                "system": "企业微信",
                "node_type": "normal"
            },
            {
                "time": "16:30",
                "descriptions": ["编写事件复盘报告"],
                "handler": "张三",
                "system": "文档平台",
                "node_type": "normal"
            },
            {
                "time": "17:00",
                "descriptions": ["召开事件复盘会议"],
                "handler": "全体参与人员",
                "system": "会议室",
                "node_type": "special"
            },
            {
                "time": "17:30",
                "descriptions": ["制定改进措施"],
                "handler": "技术团队",
                "system": "改进计划",
                "node_type": "normal"
            },
            {
                "time": "18:00",
                "descriptions": ["关闭应急通道"],
                "handler": "李四",
                "system": "应急指挥平台",
                "node_type": "normal"
            },
            {
                "time": "18:30",
                "descriptions": ["事件正式关闭"],
                "handler": "监控系统",
                "system": "事件管理平台",
                "node_type": "end"
            }
        ]
    }
    
    print("=" * 80)
    print("开始生成30节点时间轴测试图...")
    print("=" * 80)
    print(f"标题: {data['title']}")
    print(f"事件ID: {data['incident_id']}")
    print(f"时间范围: {data['start_time']} - {data['end_time']}")
    print(f"事件数量: {len(data['events'])}")
    print()
    
    # 打印所有事件概览
    print("事件列表:")
    for i, event in enumerate(data['events'], 1):
        descriptions = event.get('descriptions', [event.get('description', '')])
        desc_text = ' | '.join(descriptions)
        print(f"  {i:2d}. [{event['time']}] {desc_text[:50]}... ({event.get('handler', 'N/A')})")
    
    print()
    print("-" * 80)
    
    # 生成时间轴图
    try:
        output_path = "test_30_nodes_timeline.png"
        result = TimelineChartGenerator.generate_timeline(data, output_path=output_path)
        
        print("\n[OK] 时间轴图生成成功!")
        print(f"输出文件: {result['output_path']}")
        print(f"文件大小: {result.get('file_size_readable', 'N/A')}")
        print(f"事件数量: {result.get('event_count', len(data['events']))}")
        print(f"总事件项数: {result.get('total_event_items', 'N/A')}")
        print(f"图片格式: {result.get('image_format', 'PNG')}")
        print()
        print("=" * 80)
        print("提示: 请查看生成的 test_30_nodes_timeline.png 文件")
        print("=" * 80)
        
        return result
        
    except Exception as e:
        print(f"\n[FAIL] 生成失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    generate_30_node_timeline()
