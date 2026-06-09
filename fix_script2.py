import re

with open("gui_main.py", "r", encoding="utf-8") as f:
    content = f.read()

# 添加页面方向检测
old_thumbnails = '''        # 预加载 PDF 缩略图（保持引用防止被垃圾回收）
        page_images = []
        first_pdf = self.swap_pdfs[0] if self.swap_pdfs else None
        
        def load_page_thumbnails():
            if not first_pdf:
                return
            
            try:
                doc = fitz.open(first_pdf)
                for page_num in range(1, min(max_pages + 1, 51)):
                    if page_num <= len(doc):
                        try:
                            page = doc.load_page(page_num - 1)
                            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
                            img_data = pix.tobytes("ppm")
                            photo = tk.PhotoImage(data=img_data)
                            page_images.append(photo)
                        except Exception as e:
                            page_images.append(None)
                doc.close()
            except Exception as e:
                print(f"加载 PDF 缩略图失败：{e}")
        
        load_page_thumbnails()'''

new_thumbnails = '''        # 预加载 PDF 缩略图（保持引用防止被垃圾回收）
        page_images = []
        page_orientations = []  # 记录每页的方向（True=横向，False=纵向）
        first_pdf = self.swap_pdfs[0] if self.swap_pdfs else None
        
        def load_page_thumbnails():
            if not first_pdf:
                return
            
            try:
                doc = fitz.open(first_pdf)
                for page_num in range(1, min(max_pages + 1, 51)):
                    if page_num <= len(doc):
                        try:
                            page = doc.load_page(page_num - 1)
                            # 检测页面方向
                            page_rect = page.rect
                            is_landscape = page_rect.width > page_rect.height
                            page_orientations.append(is_landscape)
                            
                            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
                            img_data = pix.tobytes("ppm")
                            photo = tk.PhotoImage(data=img_data)
                            page_images.append(photo)
                        except Exception as e:
                            page_images.append(None)
                            page_orientations.append(False)
                doc.close()
            except Exception as e:
                print(f"加载 PDF 缩略图失败：{e}")
        
        load_page_thumbnails()'''

content = content.replace(old_thumbnails, new_thumbnails)

# 修改绘制逻辑以适配横向页面
old_draw = '''            for i in range(display_pages):
                page_num = self.swap_page_order[i]
                row = i // cols
                col = i % cols
                x = col * (cell_width + padding) + padding
                y = row * (cell_height + padding) + padding

                # 绘制方框背景
                bg_color = "#ffffff" if i != selected_index[0] else "#fff3cd"   
                pages_canvas.create_rectangle(x, y, x + cell_width, y + cell_height,
                                             outline="#3366cc" if i != selected_index[0] else "#ff6600",
                                             width=2 if i != selected_index[0] else 3,
                                             fill=bg_color)'''

new_draw = '''            for i in range(display_pages):
                page_num = self.swap_page_order[i]
                row = i // cols
                col = i % cols
                
                # 根据页面方向调整框框尺寸
                is_landscape = page_orientations[i] if i < len(page_orientations) else False
                if is_landscape:
                    current_cell_width = cell_height
                    current_cell_height = cell_width
                else:
                    current_cell_width = cell_width
                    current_cell_height = cell_height
                
                x = col * (cell_width + padding) + padding
                y = row * (cell_height + padding) + padding

                # 绘制方框背景
                bg_color = "#ffffff" if i != selected_index[0] else "#fff3cd"   
                pages_canvas.create_rectangle(x, y, x + current_cell_width, y + current_cell_height,
                                             outline="#3366cc" if i != selected_index[0] else "#ff6600",
                                             width=2 if i != selected_index[0] else 3,
                                             fill=bg_color)'''

content = content.replace(old_draw, new_draw)

with open("gui_main.py", "w", encoding="utf-8") as f:
    f.write(content)

print("修改完成")
