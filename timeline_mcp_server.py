"""
时间轴图MCP节点 - 根据事件处置过程生成专业时间轴图

工作流位置：第4步 - MCP节点（共享MCP实例模式）
功能：根据LLM分析并格式化后的处置过程描述，生成PNG格式的时间轴图

依赖安装：
    pip install matplotlib pillow fastmcp

运行方式（共享MCP实例模式）：
    python main.py                                       # 与unified_mcp_server共享一个端口

运行方式（独立模式-向后兼容）：
    python timeline_mcp_server.py                        # stdio模式
    python timeline_mcp_server.py --transport sse --port 8002   # SSE模式
"""

import json
import os
import sys
import logging
import argparse
import hashlib
from pathlib import Path
from typing import Optional, List, Dict, Any
from datetime import datetime

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger('TimelineMCPServer')

# FastMCP导入 - 优先使用共享MCP实例（单端口模式），fallback到独立实例
try:
    from mcp_instance import mcp
    FASTMCP_AVAILABLE = True
    MCP_SHARED_MODE = True  # 共享模式标志
except ImportError:
    try:
        from fastmcp import FastMCP
        FASTMCP_AVAILABLE = True
        MCP_SHARED_MODE = False  # 独立模式
    except ImportError:
        FASTMCP_AVAILABLE = False
        MCP_SHARED_MODE = False
        logger.warning("fastmcp未安装，请运行: pip install fastmcp")

# 图表绘制库
try:
    import matplotlib
    matplotlib.use('Agg')  # 非交互式后端，支持服务器环境
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import matplotlib.dates as mdates
    from matplotlib.patches import FancyBboxPatch
    from matplotlib.font_manager import FontProperties
    
    # 配置中文字体
    _CJK_FONT = None
    for _font_name in ['Microsoft YaHei', 'SimHei', 'Noto Sans SC', 'STHeiti', 'Arial Unicode MS']:
        try:
            _CJK_FONT = FontProperties(family=_font_name)
            plt.text(0, 0, '测试', fontproperties=_CJK_FONT)
            plt.close()
            # 设置全局字体，确保所有文本绘制使用中文字体
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [_font_name, 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            logger.info(f"使用中文字体: {_font_name}")
            break
        except Exception:
            _CJK_FONT = None
            continue
    
    if _CJK_FONT is None:
        try:
            plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            _CJK_FONT = FontProperties()
            logger.warning("使用fallback字体配置，部分中文可能不显示")
        except Exception:
            logger.warning("中文配置失败，使用默认字体")
            _CJK_FONT = FontProperties()
    
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    logger.warning("matplotlib未安装，请运行: pip install matplotlib")

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    logger.warning("Pillow未安装，请运行: pip install pillow")


# ============================================================
# 安全检查工具
# ============================================================

def safe_path_join(base_dir: str, user_path: str) -> str:
    """
    安全拼接路径，防止目录穿越攻击
    
    Args:
        base_dir: 基础目录
        user_path: 用户输入的路径
        
    Returns:
        安全拼接后的绝对路径
        
    Raises:
        ValueError: 检测到目录穿越
    """
    base = os.path.abspath(base_dir)
    # 如果user_path是绝对路径，直接校验
    if os.path.isabs(user_path):
        abs_path = os.path.normpath(user_path)
    else:
        abs_path = os.path.normpath(os.path.join(base, user_path))
    
    # 检查是否在base目录内（防止 ../ 穿越）
    if not abs_path.startswith(base):
        raise ValueError(f"路径安全校验失败：不允许访问 {base} 之外的路径")
    
    return abs_path


# ============================================================
# 时间轴图生成器
# ============================================================

class TimelineChartGenerator:
    """时间轴图生成器 - 根据事件数据生成专业时间轴图"""
    
    # 节点类型配置
    NODE_STYLES = {
        'start': {
            'marker': 'o',
            'markersize': 18,
            'color': '#2E7D32',       # 深绿色
            'edgecolor': '#1B5E20',
            'label': '开始',
            'zorder': 5
        },
        'end': {
            'marker': 'o',
            'markersize': 18,
            'color': '#C62828',       # 深红色
            'edgecolor': '#B71C1C',
            'label': '结束',
            'zorder': 5
        },
        'special': {
            'marker': 'D',            # 菱形
            'markersize': 14,
            'color': '#E65100',       # 橙色
            'edgecolor': '#BF360C',
            'label': '特殊节点',
            'zorder': 4
        },
        'normal': {
            'marker': 'o',
            'markersize': 12,
            'color': '#1565C0',       # 深蓝色
            'edgecolor': '#0D47A1',
            'label': '普通节点',
            'zorder': 3
        }
    }
    
    # 处理方颜色映射
    HANDLER_COLORS = [
        '#1565C0',  # 蓝
        '#2E7D32',  # 绿
        '#E65100',  # 橙
        '#6A1B9A',  # 紫
        '#00838F',  # 青
        '#C62828',  # 红
        '#F9A825',  # 黄
        '#4E342E',  # 棕
    ]

    @staticmethod
    def validate_events_data(data: Dict[str, Any]) -> Dict[str, Any]:
        """
        验证事件数据格式
        
        Args:
            data: 事件数据字典
            
        Returns:
            校验结果，包含 valid、error、event_count 等字段
        """
        if not data:
            return {'valid': False, 'error': '输入数据为空'}
        
        if not isinstance(data, dict):
            return {'valid': False, 'error': '输入必须是JSON对象'}
        
        events = data.get('events', [])
        if not events:
            return {'valid': False, 'error': 'events数组为空'}
        
        if not isinstance(events, list):
            return {'valid': False, 'error': 'events必须是数组'}
        
        # 校验每个事件
        errors = []
        for i, event in enumerate(events):
            if not isinstance(event, dict):
                errors.append(f'events[{i}]不是对象')
                continue
            if 'time' not in event:
                errors.append(f'events[{i}]缺少time字段')
            if 'description' not in event and 'descriptions' not in event:
                errors.append(f'events[{i}]缺少description或descriptions字段')
        
        if errors:
            return {'valid': False, 'error': '；'.join(errors), 'error_count': len(errors)}
        
        # 查找开始和结束节点
        node_types = [e.get('node_type', 'normal') for e in events]
        start_count = node_types.count('start')
        end_count = node_types.count('end')
        
        warnings = []
        if start_count == 0:
            warnings.append('未标记开始节点(node_type=start)')
        if end_count == 0:
            warnings.append('未标记结束节点(node_type=end)')
        if start_count > 1:
            warnings.append(f'有{start_count}个开始节点')
        if end_count > 1:
            warnings.append(f'有{end_count}个结束节点')
        
        result = {
            'valid': True,
            'event_count': len(events),
            'node_types': {t: node_types.count(t) for t in set(node_types)}
        }
        if warnings:
            result['warnings'] = '；'.join(warnings)
        
        return result

    @staticmethod
    def generate_timeline(data: Dict[str, Any], output_path: str = "",
                           width: int = 2000, height: int = 1000,
                           dpi: int = 150) -> Dict[str, Any]:
        """
        生成时间轴图（横向递进布局，支持同一时间多事件）
        
        Args:
            data: 事件数据字典
            output_path: 输出图片路径（可选，默认自动生成）
            width: 图片宽度（像素）
            height: 图片高度（像素）
            dpi: 图片分辨率
            
        Returns:
            包含success、output_path、event_count等字段的字典
        """
        logger.info("开始生成时间轴图")
        
        if not MATPLOTLIB_AVAILABLE:
            return {
                'success': False,
                'error': 'matplotlib未安装，无法生成图表',
                'error_type': 'DependencyError'
            }
        
        # 验证数据
        validation = TimelineChartGenerator.validate_events_data(data)
        if not validation['valid']:
            return {
                'success': False,
                'error': f'数据验证失败: {validation["error"]}',
                'error_type': 'ValidationError'
            }
        
        events = data.get('events', [])
        title = data.get('title', '事件处置时间轴')
        incident_id = data.get('incident_id', '')
        
        try:
            if not output_path:
                safe_name = "timeline_" + hashlib.md5(
                    json.dumps(data, ensure_ascii=False).encode()
                ).hexdigest()[:8]
                output_path = f"{safe_name}.png"
            
            output_dir = os.path.dirname(os.path.abspath(output_path))
            os.makedirs(output_dir, exist_ok=True)
            
            # ==================== 数据分组 ====================
            
            # 按时间分组
            time_groups = {}
            time_order = []
            
            for event in events:
                t = event.get('time', '')
                descriptions = event.get('descriptions', None)
                if descriptions is None:
                    descriptions = [event.get('description', '')]
                elif isinstance(descriptions, str):
                    descriptions = [descriptions]
                
                if t not in time_groups:
                    time_groups[t] = {
                        'time': t,
                        'descriptions': [],
                        'handlers': [],
                        'systems': [],
                        'node_type': event.get('node_type', 'normal'),
                    }
                    time_order.append(t)
                
                for d in descriptions:
                    if d:
                        time_groups[t]['descriptions'].append(d)
                
                handler = event.get('handler', '').strip()
                if handler and handler not in time_groups[t]['handlers']:
                    time_groups[t]['handlers'].append(handler)
                
                system = event.get('system', '').strip()
                if system and system not in time_groups[t]['systems']:
                    time_groups[t]['systems'].append(system)
                
                nt = event.get('node_type', 'normal')
                if nt in ('special', 'start', 'end'):
                    time_groups[t]['node_type'] = nt
            
            if time_order:
                time_groups[time_order[0]]['node_type'] = 'start'
                time_groups[time_order[-1]]['node_type'] = 'end'
            
            grouped_events = [time_groups[t] for t in time_order]
            n_groups = len(grouped_events)
            
            if n_groups < 2:
                return {
                    'success': False,
                    'error': '时间点数量不足，至少需要2个不同时间点',
                    'error_type': 'ValidationError'
                }
            
            # 收集handler
            all_handlers = set()
            for group in grouped_events:
                for h in group['handlers']:
                    all_handlers.add(h)
            handler_list = sorted(all_handlers)
            handler_color_map = {}
            for i, h in enumerate(handler_list):
                handler_color_map[h] = TimelineChartGenerator.HANDLER_COLORS[
                    i % len(TimelineChartGenerator.HANDLER_COLORS)
                ]
            
            # ==================== 自适应宽度 ====================
            
            # 计算每个时间点描述最大宽度（字符数）
            max_desc_chars = 0
            for g in grouped_events:
                for d in g['descriptions']:
                    # 中文字符算2个宽度
                    w = sum(2 if ord(c) > 127 else 1 for c in d)
                    if w > max_desc_chars:
                        max_desc_chars = w
            
            # 根据事件数和描述长度自适应宽度
            base_width_per_group = 1.0 / max(n_groups, 6)  # 每个时间点最小宽度
            desc_width = max_desc_chars * 0.008  # 描述文本所需宽度
            required_width = max(base_width_per_group, desc_width)
            adjusted_width = max(width, int(n_groups * required_width * dpi + 200))
            
            # ==================== 创建画布 ====================
            
            fig, ax = plt.subplots(figsize=(adjusted_width / dpi, height / dpi), dpi=dpi)
            fig.patch.set_facecolor('#FAFBFC')

            # ==================== 多行时间轴布局 ====================
            
            MAX_PER_ROW = 7  # 每行最多7个时间点
            margin_left = 0.03
            margin_right = 0.03
            
            # 蛇形布局：第一行从左到右，第二行从右到左，第三行从左到右...
            if n_groups <= MAX_PER_ROW:
                row_assignments = [(i, 0) for i in range(n_groups)]
                n_rows = 1
            else:
                row_assignments = []
                for gi in range(n_groups):
                    row_idx = gi // MAX_PER_ROW
                    row_assignments.append((gi, row_idx))
                n_rows = max(ri for _, ri in row_assignments) + 1
            
            # 每行的Y位置：row 0 在顶部，最后一行在底部
            row_spacing = 0.50
            row_centers = []
            for ri in range(n_rows):
                y_pos = 0.75 - ri * row_spacing
                row_centers.append(y_pos)
            
            # 分配每行内的时间点，蛇形排布
            row_groups = {}
            for gi, ri in row_assignments:
                if ri not in row_groups:
                    row_groups[ri] = []
                row_groups[ri].append(gi)
            
            # 对奇数行反转顺序（蛇形），并计算坐标
            node_positions = {}
            for ri, group_indices in row_groups.items():
                n_in_row = len(group_indices)
                if ri % 2 == 1:
                    # 奇数行：从右到左（蛇形反转）
                    ordered = list(reversed(group_indices))
                else:
                    # 偶数行：从左到右
                    ordered = group_indices
                
                if n_in_row <= 1:
                    for gi in ordered:
                        node_positions[gi] = (0.5, row_centers[ri])
                else:
                    for j, gi in enumerate(ordered):
                        x_pos = margin_left + j * (1 - margin_left - margin_right) / (n_in_row - 1)
                        node_positions[gi] = (x_pos, row_centers[ri])
            
            # 绘制每行的时间轴主线
            for ri, center_y in enumerate(row_centers):
                group_indices = row_groups[ri]
                n_in_row = len(group_indices)
                if n_in_row <= 1:
                    x_start = x_end = 0.5
                else:
                    x_start = margin_left
                    x_end = 1 - margin_right
                
                ax.axhline(y=center_y, xmin=0.02, xmax=0.98,
                          color='#78909C', linewidth=2.5, linestyle='-', alpha=0.6, zorder=1)
                # 两端装饰
                ax.plot([x_start, x_start], [center_y - 0.01, center_y + 0.01],
                       color='#78909C', linewidth=2, alpha=0.6, zorder=1)
                ax.plot([x_end, x_end], [center_y - 0.01, center_y + 0.01],
                       color='#78909C', linewidth=2, alpha=0.6, zorder=1)
                
                # 行间连接：从当前行最后一点虚线连接到下一行第一个点
                if ri < n_rows - 1:
                    last_idx = group_indices[-1]
                    first_idx_next = row_groups[ri + 1][0]
                    last_x, last_y = node_positions[last_idx]
                    first_x_next, first_y_next = node_positions[first_idx_next]
                    ax.plot([last_x, first_x_next], [last_y, first_y_next],
                           color='#78909C', linewidth=1.2, linestyle='--', alpha=0.4, zorder=1)
            
            # ==================== 绘制事件节点 ====================
            
            # 已存储的节点覆盖范围，用于冲突检测
            drawn_texts = []  # [(x_center, y_center, height), ...]

            def has_overlap(new_x, new_y, new_h, existing):
                """检测新文本是否与已有文本重叠"""
                for ex, ey, eh in existing:
                    if abs(new_x - ex) < 0.035 and abs(new_y - ey) < (new_h + eh) * 0.4:
                        return True
                return False

            # 绘制事件节点
            for gi in range(n_groups):
                group = grouped_events[gi]
                x_pos, timeline_y = node_positions[gi]

                time_str = group['time']
                descriptions = group['descriptions']
                handlers = group['handlers']
                systems = group['systems']
                node_type = group.get('node_type', 'normal')

                style = TimelineChartGenerator.NODE_STYLES.get(
                    node_type, TimelineChartGenerator.NODE_STYLES['normal']
                )

                node_color = style['color']
                node_edgecolor = style['edgecolor']
                if handlers:
                    primary_handler = handlers[0]
                    if primary_handler in handler_color_map:
                        node_color = handler_color_map[primary_handler]

                # 所有行的内容块都显示在上方
                
                # 事件点
                ax.plot(x_pos, timeline_y, marker=style['marker'],
                       markersize=style['markersize'],
                       color=node_color, markeredgecolor=node_edgecolor,
                       markeredgewidth=2, zorder=style['zorder'])
                
                if node_type in ('start', 'end'):
                    ax.plot(x_pos, timeline_y, marker='o', markersize=style['markersize'] + 6,
                           color=node_color, alpha=0.12, zorder=style['zorder'] - 1)
                
                # 时间标签
                time_label_y = timeline_y + 0.055
                ax.text(x_pos, time_label_y, time_str,
                       fontsize=11, fontweight='bold',
                       ha='center', va='center',
                       color='#37474F')
                
                # ==================== 分支节点布局 ====================
                # 同一时间点多条内容时，各条独立分支显示
                # 单条内容居中显示，多条内容在两侧扇形排布
                
                MAX_CHARS_PER_LINE = 12
                desc_lines = descriptions[:]
                
                # handler和system信息
                system_str = ' | '.join(list(set(systems))) if systems else ''
                handler_str = '  '.join([f'[{h}]' for h in handlers]) if handlers else ''
                
                # 为每条描述构建独立的显示文本
                desc_items = []
                for dl in desc_lines:
                    # 自动换行
                    wrapped_lines = []
                    remaining = dl
                    while len(remaining) > MAX_CHARS_PER_LINE:
                        wrapped_lines.append(remaining[:MAX_CHARS_PER_LINE])
                        remaining = remaining[MAX_CHARS_PER_LINE:]
                    if remaining:
                        wrapped_lines.append(remaining)
                    desc_items.append({
                        'text_lines': wrapped_lines,
                        'total_lines': len(wrapped_lines)
                    })
                
                n_desc = len(desc_items)
                # handler/system信息行
                info_line = ''
                info_parts = []
                if handler_str:
                    info_parts.append(handler_str)
                if system_str:
                    info_parts.append(system_str)
                has_info = len(info_parts) > 0
                
                if n_desc == 1:
                    # 单条：居中显示（原有逻辑）
                    display_lines = desc_items[0]['text_lines']
                    if has_info:
                        display_lines.append('')
                        display_lines.append('  '.join(info_parts))
                    display_text = '\n'.join(display_lines)
                    n_total_lines = len(display_lines)
                    
                    base_offset = 0.08
                    text_offset = base_offset + max(0, n_total_lines - 1) * 0.032
                    label_y = timeline_y + text_offset
                    
                    text_height = n_total_lines * 0.032 + 0.03
                    
                    # 特殊节点背景框
                    if node_type == 'special':
                        bbox_props = dict(boxstyle='round,pad=0.5', facecolor='#FFF8E1', edgecolor='#F57C00', linewidth=1.5)
                        ax.text(x_pos, label_y, display_text, fontsize=9, ha='center', va='bottom', color='#212121', bbox=bbox_props)
                    else:
                        bbox_props = dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='#E0E0E0', linewidth=0.5, alpha=0.8)
                        ax.text(x_pos, label_y, display_text, fontsize=9, ha='center', va='bottom', color='#424242', bbox=bbox_props)
                    
                    # 箭头连接线
                    ax.annotate('', xy=(x_pos, timeline_y + 0.05), xytext=(x_pos, timeline_y + 0.01),
                               arrowprops=dict(arrowstyle='->', color='#78909C', linestyle='--', lw=1.5, alpha=0.5))
                    
                else:
                    # 多条：纵向顺序排列（所有分支垂直堆叠在时间点上方）
                    branch_base_y_offset = 0.08
                    branch_y_step = 0.12
                    
                    # 计算所有分支的总高度，用于判断是否需要增大行间距
                    all_item_lines = sum(item['total_lines'] for item in desc_items) + n_desc
                    branch_height = branch_base_y_offset + n_desc * branch_y_step + all_item_lines * 0.032
                    
                    for di, item in enumerate(desc_items):
                        # 所有分支垂直居中堆叠在时间点正上方（不左右偏移）
                        bx = x_pos
                        
                        # Y位置：递增高度
                        item_lines = item['text_lines']
                        n_item_lines = item['total_lines']
                        
                        # 组装显示文本
                        d_lines = list(item_lines)
                        # 为该分支添加handler信息
                        # 如果所有描述共享同一个handler，每一条都显示
                        branch_info = []
                        if handlers:
                            h_idx = di % len(handlers)
                            branch_info.append(f'[{handlers[h_idx]}]')
                        if systems:
                            s_idx = di % len(systems)
                            branch_info.append(f'|{systems[s_idx]}|')
                        if branch_info:
                            d_lines.append(' '.join(branch_info))
                        
                        item_text = '\n'.join(d_lines)
                        item_total_lines = len(d_lines)
                        
                        # 从上到下排列分支
                        by = timeline_y + branch_base_y_offset + di * branch_y_step
                        # 更靠边的分支略低一点
                        if di > 1:
                            by += (di - 1) * 0.01
                        
                        # 分支连接线（从主节点到卡片）
                        ax.plot([x_pos, bx], [timeline_y + 0.01, by - 0.005],
                               color='#90A4AE', linewidth=1.0, linestyle='--', alpha=0.4, zorder=1)
                        
                        # 分支内容卡片（紧凑小框）
                        card_bbox = dict(
                            boxstyle='round,pad=0.2',
                            facecolor='white', edgecolor='#B0BEC5',
                            linewidth=0.5, alpha=0.85
                        )
                        if node_type == 'special' and di == 0:
                            card_bbox = dict(
                                boxstyle='round,pad=0.25',
                                facecolor='#FFF8E1', edgecolor='#F57C00',
                                linewidth=1.0
                            )
                        
                        ax.text(bx, by, item_text,
                               fontsize=7.5, ha='center', va='bottom',
                               color='#424242', bbox=card_bbox)
                
                # 开始/结束标记
                if node_type == 'start':
                    ax.text(x_pos, timeline_y + 0.10, 'START',
                           fontsize=8, fontweight='bold',
                           ha='center', va='center',
                           color='#2E7D32',
                           style='italic')
                if node_type == 'end':
                    ax.text(x_pos, timeline_y - 0.10, 'END',
                           fontsize=8, fontweight='bold',
                           ha='center', va='center',
                           color='#C62828',
                           style='italic')
            
            # ==================== 标题和图例 ====================
            
            title_text = title
            if incident_id:
                title_text += f'\n事件编号: {incident_id}'
            
            ax.set_title(title_text, fontsize=15, fontweight='bold',
                        color='#1A237E', pad=16, linespacing=1.3)
            
            ax.set_xlim(0, 1)
            y_min = min(row_centers) - row_spacing * 0.6
            y_max = max(row_centers) + row_spacing * 0.6
            ax.set_ylim(y_min - 0.1, y_max + 0.15)
            
            ax.axis('off')
            
            # 图例
            legend_elements = []
            for ntype, nstyle in TimelineChartGenerator.NODE_STYLES.items():
                legend_elements.append(
                    plt.Line2D([0], [0], marker=nstyle['marker'], color='w',
                              markerfacecolor=nstyle['color'],
                              markeredgecolor=nstyle['edgecolor'],
                              markersize=9, label=nstyle['label'])
                )
            
            if handler_list:
                legend_elements.append(
                    plt.Line2D([0], [0], marker='', color='w', label='')
                )
                for handler_name, color in handler_color_map.items():
                    legend_elements.append(
                        plt.Line2D([0], [0], marker='o', color='w',
                                  markerfacecolor=color, markersize=7,
                                  label=f'处理人: {handler_name}')
                    )
            
            if data.get('start_time') or data.get('end_time'):
                info_text = ''
                if data.get('start_time'):
                    info_text += f'开始: {data["start_time"]}  '
                if data.get('end_time'):
                    info_text += f'结束: {data["end_time"]}'
                ax.text(0.5, 0.01, info_text,
                       fontsize=8, color='#90A4AE',
                       ha='center', va='bottom',
                       transform=ax.transAxes)
            
            ax.text(0.98, 0.005, f'生成: {datetime.now().strftime("%Y-%m-%d %H:%M")}',
                   fontsize=6, color='#B0BEC5',
                   ha='right', va='bottom',
                   transform=ax.transAxes)
            
            if legend_elements:
                legend = fig.legend(
                    handles=legend_elements,
                    loc='lower center',
                    ncol=min(10, len(legend_elements)),
                    fontsize=7.5, framealpha=0.95,
                    edgecolor='#CFD8DC', facecolor='white',
                    title='图例说明', title_fontsize=8.5,
                    bbox_to_anchor=(0.5, -0.01)
                )
                legend.get_title().set_fontweight('bold')
            
            plt.tight_layout(pad=1.5)
            fig.savefig(output_path, dpi=dpi, bbox_inches='tight',
                       facecolor=fig.get_facecolor(), edgecolor='none')
            plt.close(fig)
            
            file_size = os.path.getsize(output_path)
            logger.info(f"时间轴图已生成: {output_path} ({file_size / 1024:.1f} KB)")
            
            return {
                'success': True,
                'output_path': os.path.abspath(output_path),
                'file_size': file_size,
                'file_size_readable': f"{file_size / 1024:.1f} KB",
                'event_count': n_groups,
                'total_event_items': len(events),
                'image_format': 'PNG',
                'width': adjusted_width,
                'height': height,
                'dpi': dpi
            }
            
        except Exception as e:
            logger.exception("时间轴图生成失败")
            return {
                'success': False,
                'error': f'生成失败: {str(e)}',
                'error_type': type(e).__name__
            }

    @staticmethod
    def generate_timeline_v2(data_json: str, output_path: str = "",
                              width: int = 1600, height: int = 900,
                              dpi: int = 150) -> str:
        """
        生成时间轴图（字符串参数版本，适合MCP工具调用）
        
        Args:
            data_json: 事件数据JSON字符串
            output_path: 输出图片路径（可选）
            width: 图片宽度
            height: 图片高度
            dpi: 图片分辨率
            
        Returns:
            JSON字符串，包含生成结果
        """
        start_time = datetime.now()
        call_id = f"timeline_{start_time.strftime('%Y%m%d_%H%M%S')}"
        
        logger.info(f"[{call_id}] 开始生成时间轴图")
        
        try:
            # 解析JSON
            data = json.loads(data_json)
            
            result = TimelineChartGenerator.generate_timeline(
                data, output_path, width, height, dpi
            )
            result['call_id'] = call_id
            result['timestamp'] = datetime.now().isoformat()
            
            return json.dumps(result, ensure_ascii=False, indent=2)
            
        except json.JSONDecodeError as e:
            logger.error(f"[{call_id}] JSON解析失败: {e}")
            return json.dumps({
                'success': False,
                'error': f'JSON解析失败: {str(e)}',
                'error_type': 'JSONParseError',
                'call_id': call_id,
                'timestamp': datetime.now().isoformat()
            }, ensure_ascii=False)
        except Exception as e:
            logger.exception(f"[{call_id}] 生成异常")
            return json.dumps({
                'success': False,
                'error': f'生成失败: {str(e)}',
                'error_type': type(e).__name__,
                'call_id': call_id,
                'timestamp': datetime.now().isoformat()
            }, ensure_ascii=False)

    @staticmethod
    def validate_timeline_data(data_json: str) -> str:
        """
        验证时间轴数据格式
        
        Args:
            data_json: 事件数据JSON字符串
            
        Returns:
            JSON字符串，包含验证结果
        """
        start_time = datetime.now()
        call_id = f"validate_{start_time.strftime('%Y%m%d_%H%M%S')}"
        
        logger.info(f"[{call_id}] 验证时间轴数据")
        
        try:
            data = json.loads(data_json)
            result = TimelineChartGenerator.validate_events_data(data)
            result['call_id'] = call_id
            result['timestamp'] = datetime.now().isoformat()
            return json.dumps(result, ensure_ascii=False, indent=2)
            
        except json.JSONDecodeError as e:
            return json.dumps({
                'valid': False,
                'error': f'JSON解析失败: {str(e)}',
                'call_id': call_id,
                'timestamp': datetime.now().isoformat()
            }, ensure_ascii=False)

    @staticmethod
    def get_timeline_data_template() -> str:
        """
        获取时间轴数据模板（帮助第2步LLM了解输出格式）
        
        Returns:
            JSON字符串，包含数据模板
        """
        template = {
            "title": "事件处置时间轴",
            "incident_id": "INC-2024-001",
            "start_time": "2024-01-15 14:00",
            "end_time": "2024-01-15 16:30",
            "events": [
                {
                    "time": "14:00",
                    "description": "事件描述",
                    "handler": "处理人/系统",
                    "system": "关联系统",
                    "node_type": "start",
                    "category": "automation"
                },
                {
                    "time": "14:05",
                    "description": "中间事件描述",
                    "handler": "处理人",
                    "system": "关联系统",
                    "node_type": "normal",
                    "category": "manual"
                },
                {
                    "time": "14:30",
                    "description": "关键决策/异常事件",
                    "handler": "处理人",
                    "system": "关联系统",
                    "node_type": "special",
                    "category": "decision"
                },
                {
                    "time": "16:30",
                    "description": "事件处置完成",
                    "handler": "处理人",
                    "system": "关联系统",
                    "node_type": "end",
                    "category": "manual"
                }
            ],
            "field_descriptions": {
                "title": "时间轴标题",
                "incident_id": "事件编号",
                "start_time": "开始时间（ISO格式）",
                "end_time": "结束时间（ISO格式）",
                "events": "事件数组，按时间顺序排列",
                "events[].time": "事件发生时间",
                "events[].description": "事件描述",
                "events[].handler": "处理人或处理系统",
                "events[].system": "关联系统名称",
                "events[].node_type": "节点类型：start(开始)/end(结束)/normal(普通)/special(特殊)",
                "events[].category": "分类：automation(自动)/manual(人工)/decision(决策)"
            }
        }
        return json.dumps(template, ensure_ascii=False, indent=2)


# ============================================================
# MCP服务器初始化
# ============================================================

if FASTMCP_AVAILABLE:
    # 如果不是共享模式，创建独立的 FastMCP 实例
    if not MCP_SHARED_MODE:
        mcp = FastMCP("TimelineDocumentServer")

    @mcp.tool()
    def read_timeline_document(file_path: str) -> str:
        """
        读取文档内容，返回文档文本和元数据（时间轴工作流第1步）
        
        Args:
            file_path: 文档路径，支持 .docx、.wps、.wpsx 格式
            
        Returns:
            JSON字符串，包含document_content和document_metadata
        """
        start_time = datetime.now()
        call_id = f"read_{start_time.strftime('%Y%m%d_%H%M%S')}"

        logger.info(f"[{call_id}] 开始读取文档: {file_path}")

        try:
            if not file_path or not file_path.strip():
                return json.dumps({
                    "success": False,
                    "error": "file_path不能为空",
                    "error_type": "ValidationError",
                    "call_id": call_id
                }, ensure_ascii=False)

            # 尝试从unified_mcp_server导入DocumentReader
            try:
                from unified_mcp_server import DocumentReader
                reader = DocumentReader()
                result = reader.read_document(file_path)
            except ImportError:
                # fallback: 使用document_reader_mcp
                try:
                    from document_reader_mcp import DocumentReaderMCPNode
                    reader = DocumentReaderMCPNode()
                    result = reader.read_document(file_path)
                except ImportError:
                    return json.dumps({
                        "success": False,
                        "error": "无法找到文档读取模块（unified_mcp_server.py或document_reader_mcp.py）",
                        "error_type": "ImportError",
                        "call_id": call_id
                    }, ensure_ascii=False)

            result['call_id'] = call_id
            result['timestamp'] = datetime.now().isoformat()

            logger.info(f"[{call_id}] 读取完成: {result.get('document_metadata', {}).get('paragraph_count', 0)}个段落")
            return json.dumps(result, ensure_ascii=False, indent=2)

        except FileNotFoundError as e:
            return json.dumps({
                "success": False,
                "error": str(e),
                "error_type": "FileNotFoundError",
                "call_id": call_id
            }, ensure_ascii=False)
        except Exception as e:
            logger.exception(f"[{call_id}] 读取异常")
            return json.dumps({
                "success": False,
                "error": f"读取失败: {str(e)}",
                "error_type": type(e).__name__,
                "call_id": call_id,
                "timestamp": datetime.now().isoformat()
            }, ensure_ascii=False)

    @mcp.tool()
    def generate_timeline(data_json: str, output_path: str = "",
                           width: int = 1600, height: int = 900,
                           dpi: int = 150) -> str:
        """
        根据事件处置数据生成时间轴图（第4步）
        
        输入数据格式说明：
        - title: 时间轴标题（可选）
        - incident_id: 事件编号（可选）
        - start_time: 开始时间（可选）
        - end_time: 结束时间（可选）
        - events: 事件数组（必需），每个事件包含：
            - time: 时间（必需）
            - description: 事件描述（必需）
            - handler: 处理人/系统（可选，用于着色）
            - system: 关联系统（可选）
            - node_type: 节点类型（可选，默认normal）
                start=开始节点(绿色) end=结束节点(红色) 
                special=特殊节点(橙色菱形) normal=普通节点(蓝色圆形)
            - category: 分类（可选）
        
        Args:
            data_json: 事件数据JSON字符串
            output_path: 输出图片路径（可选，默认自动生成）
            width: 图片宽度（像素，默认1600）
            height: 图片高度（像素，默认900）
            dpi: 图片分辨率（默认150）
            
        Returns:
            JSON字符串，包含生成结果（success、output_path、file_size等）
        """
        start_time = datetime.now()
        call_id = f"timeline_{start_time.strftime('%Y%m%d_%H%M%S')}"

        logger.info(f"[{call_id}] 开始生成时间轴图")
        logger.info(f"[{call_id}] output_path: {output_path}")

        try:
            # 解析JSON
            data = json.loads(data_json)
            
            result = TimelineChartGenerator.generate_timeline(
                data, output_path, width, height, dpi
            )
            result['call_id'] = call_id
            result['timestamp'] = datetime.now().isoformat()

            if result.get('success'):
                logger.info(f"[{call_id}] 时间轴图生成成功: {result.get('output_path')}, {result.get('event_count')}个事件")
            else:
                logger.error(f"[{call_id}] 时间轴图生成失败: {result.get('error')}")
            
            return json.dumps(result, ensure_ascii=False, indent=2)

        except json.JSONDecodeError as e:
            logger.error(f"[{call_id}] JSON解析失败: {e}")
            return json.dumps({
                "success": False,
                "error": f"JSON解析失败: {str(e)}",
                "error_type": "JSONParseError",
                "call_id": call_id,
                "timestamp": datetime.now().isoformat()
            }, ensure_ascii=False)
        except Exception as e:
            logger.exception(f"[{call_id}] 生成异常")
            return json.dumps({
                "success": False,
                "error": f"生成失败: {str(e)}",
                "error_type": type(e).__name__,
                "call_id": call_id,
                "timestamp": datetime.now().isoformat()
            }, ensure_ascii=False)

    @mcp.tool()
    def validate_timeline_data(data_json: str) -> str:
        """
        验证时间轴数据格式
        
        Args:
            data_json: 事件数据JSON字符串
            
        Returns:
            JSON字符串，包含验证结果（valid、event_count、warnings等）
        """
        start_time = datetime.now()
        call_id = f"validate_{start_time.strftime('%Y%m%d_%H%M%S')}"

        logger.info(f"[{call_id}] 验证时间轴数据")

        try:
            data = json.loads(data_json)
            result = TimelineChartGenerator.validate_events_data(data)
            result['call_id'] = call_id
            result['timestamp'] = datetime.now().isoformat()

            logger.info(f"[{call_id}] 验证结果: valid={result.get('valid')}, events={result.get('event_count', 0)}")
            return json.dumps(result, ensure_ascii=False, indent=2)

        except json.JSONDecodeError as e:
            return json.dumps({
                "valid": False,
                "error": f"JSON解析失败: {str(e)}",
                "call_id": call_id,
                "timestamp": datetime.now().isoformat()
            }, ensure_ascii=False)

    @mcp.tool()
    def get_timeline_template() -> str:
        """
        获取时间轴数据模板（帮助LLM了解输出格式）
        
        Returns:
            JSON字符串，包含完整的数据模板和字段说明
        """
        return TimelineChartGenerator.get_timeline_data_template()

    @mcp.tool()
    def get_timeline_tools_info() -> str:
        """获取时间轴图工具信息"""
        tools_info = {
            "name": "TimelineDocumentServer",
            "version": "1.0.0",
            "description": "时间轴图MCP服务器，支持文档读取和时间轴图生成",
            "tools": [
                {
                    "name": "read_timeline_document",
                    "description": "读取文档内容（时间轴工作流第1步）",
                    "category": "文档读取"
                },
                {
                    "name": "generate_timeline",
                    "description": "根据事件数据生成时间轴PNG图片（第4步）",
                    "category": "时间轴图生成"
                },
                {
                    "name": "validate_timeline_data",
                    "description": "验证时间轴数据格式",
                    "category": "时间轴图生成"
                },
                {
                    "name": "get_timeline_template",
                    "description": "获取时间轴数据模板",
                    "category": "时间轴图生成"
                }
            ]
        }
        return json.dumps(tools_info, ensure_ascii=False, indent=2)


# ============================================================
# 主入口
# ============================================================

def main():
    """主入口函数"""
    parser = argparse.ArgumentParser(description='时间轴图MCP服务器')
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
        default=8002,
        help='SSE模式下的端口号 (默认: 8002)'
    )

    args = parser.parse_args()

    print("=" * 70)
    print("时间轴图MCP服务器 v1.0.0")
    print("功能：文档读取 + 时间轴图生成")
    print("=" * 70)

    if not FASTMCP_AVAILABLE:
        print("\n错误: fastmcp未安装")
        print("请运行: pip install fastmcp")
        sys.exit(1)

    if not MATPLOTLIB_AVAILABLE:
        print("\n错误: matplotlib未安装")
        print("请运行: pip install matplotlib")
        sys.exit(1)

    print(f"\n传输方式: {args.transport.upper()}")

    if args.transport == 'sse':
        print(f"服务地址: http://{args.host}:{args.port}")
        print(f"SSE端点: http://{args.host}:{args.port}/sse")
        print(f"消息端点: http://{args.host}:{args.port}/messages")

    print("\n可用工具:")
    print("  【时间轴图生成】")
    tool_list = [
        ("read_timeline_document", "读取文档内容（时间轴工作流，仅独立模式可用）"),
        ("generate_timeline", "根据事件数据生成时间轴PNG图"),
        ("validate_timeline_data", "验证时间轴数据格式"),
        ("get_timeline_template", "获取时间轴数据模板"),
        ("get_tools_info", "工具信息"),
    ]
    for i, (name, desc) in enumerate(tool_list, 1):
        print(f"    {i}. {name:<35s} - {desc}")
    print("\n启动服务...")
    print("-" * 70)

    # 启动FastMCP服务
    if args.transport == 'sse':
        mcp.run(transport='sse', host=args.host, port=args.port)
    else:
        mcp.run()


if __name__ == "__main__":
    main()
