import sys
content = open('timeline_mcp_server.py', encoding='utf-8').read()

# 替换所有 fontfamily 为 fontproperties
# 模式: fontfamily='sans-serif' -> fontproperties=_CJK_FONT
# 注意: 有 fontfamily='sans-serif' 和 fontfamily='sans-serif') 两种

replacements = [
    ("fontfamily='sans-serif'", 'fontproperties=_CJK_FONT'),
    ("fontfamily='sans-serif')", 'fontproperties=_CJK_FONT)'),
]

for old, new in replacements:
    content = content.replace(old, new)

open('timeline_mcp_server.py', 'w', encoding='utf-8').write(content)
print('替换完成')
