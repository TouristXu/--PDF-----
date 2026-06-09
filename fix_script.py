import re

with open("gui_main.py", "r", encoding="utf-8") as f:
    content = f.read()

# 修改框框尺寸（再扩大1.2倍）
content = content.replace("cell_width = 120  # 80 * 1.5", "cell_width = 144  # 80 * 1.5 * 1.2")
content = content.replace("cell_height = 150  # 100 * 1.5", "cell_height = 180  # 100 * 1.5 * 1.2")
content = content.replace("padding = 15  # 10 * 1.5", "padding = 18  # 10 * 1.5 * 1.2")

with open("gui_main.py", "w", encoding="utf-8") as f:
    f.write(content)

print("修改完成")
