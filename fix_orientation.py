with open("gui_main.py", "r", encoding="utf-8") as f:
    content = f.read()

# 添加页面方向列表初始化
content = content.replace(
    "        page_images = []\n        first_pdf =",
    "        page_images = []\n        page_orientations = []  # 记录每页的方向\n        first_pdf ="
)

# 添加页面方向检测
content = content.replace(
    "                            page = doc.load_page(page_num - 1)\n                            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))",
    "                            page = doc.load_page(page_num - 1)\n                            # 检测页面方向\n                            page_rect = page.rect\n                            is_landscape = page_rect.width > page_rect.height\n                            page_orientations.append(is_landscape)\n                            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))"
)

# 添加异常处理中的方向记录
content = content.replace(
    "                        except Exception as e:\n                            page_images.append(None)",
    "                        except Exception as e:\n                            page_images.append(None)\n                            page_orientations.append(False)"
)

with open("gui_main.py", "w", encoding="utf-8") as f:
    f.write(content)

print("页面方向检测已添加")
