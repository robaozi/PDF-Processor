import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path
import os
import threading
import re


def parse_page_range(page_range_str, total_pages):
    """
    解析页码范围字符串，返回需要编号的页码列表
    支持格式：全部 | 1-5 | 1,3,5 | 1-3,6,8-10
    """
    if not page_range_str.strip() or page_range_str.strip() == "全部":
        return list(range(1, total_pages + 1))

    pages = set()
    parts = page_range_str.split(',')

    for part in parts:
        part = part.strip()
        if '-' in part:
            # 处理范围格式 如 1-5
            try:
                start, end = map(int, part.split('-', 1))
                start = max(1, start)
                end = min(total_pages, end)
                if start <= end:
                    pages.update(range(start, end + 1))
            except ValueError:
                continue
        else:
            # 处理单个页码
            try:
                page = int(part)
                if 1 <= page <= total_pages:
                    pages.add(page)
            except ValueError:
                continue

    return sorted(list(pages))


def resize_to_a4(img, dpi=300, direction='portrait'):
    """
    将图片调整为A4纸尺寸，保持比例并居中
    """
    if direction == 'portrait':
        a4_px = (int(8.27 * dpi), int(11.69 * dpi))  # A4纵向像素尺寸
    else:
        a4_px = (int(11.69 * dpi), int(8.27 * dpi))  # A4横向像素尺寸

    img_ratio = img.width / img.height
    a4_ratio = a4_px[0] / a4_px[1]

    # 按比例缩放，确保完全适应A4纸张
    if img_ratio > a4_ratio:
        new_width = a4_px[0]
        new_height = int(new_width / img_ratio)
    else:
        new_height = a4_px[1]
        new_width = int(new_height * img_ratio)

    # 缩放图片
    resized = img.resize((new_width, new_height), Image.LANCZOS)

    # 创建A4白色背景
    bg = Image.new("RGB", a4_px, (255, 255, 255))

    # 计算居中位置并粘贴
    paste_x = (a4_px[0] - new_width) // 2
    paste_y = (a4_px[1] - new_height) // 2
    bg.paste(resized, (paste_x, paste_y))
    return bg


def merge_images(images, rows, cols):
    """
    将图片按指定行列数拼接成页面
    """
    merged = []
    for i in range(0, len(images), rows * cols):
        group = images[i:i + rows * cols]  # 获取当前组的图片
        if not group:
            break

        w, h = group[0].size
        # 创建拼接画布
        canvas = Image.new("RGB", (cols * w, rows * h), (255, 255, 255))

        # 按行列顺序粘贴图片
        for idx, img in enumerate(group):
            r = idx // cols  # 计算行位置
            c = idx % cols  # 计算列位置
            canvas.paste(img, (c * w, r * h))
        merged.append(canvas)
    return merged


def process_pdf(params, progress_callback=None, done_callback=None):
    """
    主要的PDF处理函数
    """
    try:
        input_pdf = params['pdf_path']
        dpi = int(params['dpi'])
        rows = int(params['rows'])
        cols = int(params['cols'])
        direction = "portrait" if params['direction'] == "纵向" else "landscape"
        prefix = params['prefix']
        fontsize = int(params['fontsize'])
        x_pos = int(params['x'])
        y_pos = int(params['y'])
        font_name = params['font']
        font_color = tuple(map(int, params['color'].split(',')))
        page_range_str = params['page_range']
        output_format = params['output_format'].lower()

        # 字体文件映射
        font_map = {
            "黑体": "simhei.ttf",
            "宋体": "simsun.ttc",
            "微软雅黑": "msyh.ttc",
            "Times New Roman": "times.ttf"
        }
        font_path = font_map.get(font_name, "simhei.ttf")

        # 转换PDF全文为图片
        images = convert_from_path(input_pdf, dpi=dpi)
        total_pages = len(images)

        if not images:
            raise ValueError("未能转换任何页面")

        # 解析需要编号的页码
        numbering_pages = parse_page_range(page_range_str, total_pages)

        total_steps = len(images) + 2
        processed_images = []

        # 处理每一页
        for idx, img in enumerate(images):
            current_page = idx + 1  # 当前页码

            # 调整为A4尺寸
            img = resize_to_a4(img, dpi, direction)

            # 判断当前页是否需要添加编号
            if current_page in numbering_pages:
                draw = ImageDraw.Draw(img)
                try:
                    # 加载字体
                    font = ImageFont.truetype(font_path, int(fontsize * dpi / 72))
                except:
                    # 如果字体加载失败，使用默认字体
                    font = ImageFont.load_default()

                # 添加编号文字，使用在编号页面中的顺序
                numbering_index = numbering_pages.index(current_page) + 1
                draw.text((x_pos, y_pos), f"{prefix}.{numbering_index}", fill=font_color, font=font)

            processed_images.append(img)

            # 更新进度
            if progress_callback:
                progress_callback((idx + 1) / total_steps * 100)

        # 拼接图片
        merged_images = merge_images(processed_images, rows, cols)

        # 生成输出文件路径
        output_path = input_pdf.replace(".pdf", f"_输出.{output_format}")

        # 保存文件
        if output_format == 'pdf':
            # 保存为PDF
            merged_images[0].save(output_path, save_all=True, append_images=merged_images[1:])
        else:
            # 保存为JPG
            for i, img in enumerate(merged_images):
                img_path = output_path.replace('.jpg', f'_{i + 1}.jpg')
                img.save(img_path, dpi=(dpi, dpi))

        # 处理完成回调
        if done_callback:
            done_callback(output_path)

    except Exception as e:
        messagebox.showerror("错误", f"处理失败：{str(e)}")


def run_gui():
    """
    运行图形界面
    """
    root = tk.Tk()
    root.title("PDF 编号与拼页工具")
    root.geometry("600x700")

    # 定义所有输入变量
    vars = {
        "pdf_path": tk.StringVar(),  # PDF文件路径
        "prefix": tk.StringVar(value="2.2.21"),  # 编号前缀
        "font": tk.StringVar(value="黑体"),  # 字体
        "fontsize": tk.StringVar(value="65"),  # 字号
        "color_name": tk.StringVar(value="红色"),  # 颜色名称
        "color": tk.StringVar(value="255,0,0"),  # RGB颜色值
        "x": tk.StringVar(value="40"),  # X坐标
        "y": tk.StringVar(value="40"),  # Y坐标
        "dpi": tk.StringVar(value="300"),  # 渲染DPI
        "rows": tk.StringVar(value="2"),  # 拼页行数
        "cols": tk.StringVar(value="2"),  # 拼页列数
        "direction": tk.StringVar(value="纵向"),  # 页面方向
        "page_range": tk.StringVar(value="全部"),  # 编号页码范围
        "output_format": tk.StringVar(value="PDF"),  # 输出格式
    }

    def browse():
        """浏览并选择PDF文件"""
        file = filedialog.askopenfilename(filetypes=[("PDF 文件", "*.pdf")])
        if file:
            vars["pdf_path"].set(file)

    def update_color(_=None):
        """更新颜色RGB值"""
        color_map = {
            "红色": "255,0,0",
            "黑色": "0,0,0",
            "蓝色": "0,0,255",
            "绿色": "0,128,0",
            "灰色": "128,128,128"
        }
        vars["color"].set(color_map.get(vars["color_name"].get(), "255,0,0"))

    def start_process():
        """开始处理PDF"""
        # 获取所有参数
        params = {k: v.get() for k, v in vars.items()}

        # 验证PDF文件
        if not os.path.isfile(params["pdf_path"]):
            messagebox.showerror("错误", "请选择有效的 PDF 文件")
            return

        # 重置进度条
        progress["value"] = 0

        # 在新线程中处理，避免界面卡死
        def thread_func():
            process_pdf(
                params,
                lambda p: progress.configure(value=p),
                lambda p: messagebox.showinfo("完成", f"处理完成！\n文件已保存：{p}")
            )

        threading.Thread(target=thread_func, daemon=True).start()

    # 创建主框架并居中
    main_frame = ttk.Frame(root)
    main_frame.pack(expand=True, fill='both', padx=20, pady=20)

    # 创建内容框架
    content = ttk.Frame(main_frame)
    content.pack(anchor='center')

    def create_row(parent, label_text, widget_type, var, options=None, command=None, browse_func=None):
        """创建一行输入控件"""
        # 创建行框架
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill='x', pady=5)

        # 标签 - 左对齐，固定宽度
        label = ttk.Label(row_frame, text=label_text, width=15, anchor='w')
        label.pack(side='left', padx=(0, 10))

        # 输入控件
        if widget_type == 'entry':
            widget = ttk.Entry(row_frame, textvariable=var, width=25)
            widget.pack(side='left')
            if browse_func:
                ttk.Button(row_frame, text="浏览", command=browse_func).pack(side='left', padx=(5, 0))
        elif widget_type == 'combobox':
            widget = ttk.Combobox(row_frame, textvariable=var, values=options, width=22, state='readonly')
            widget.pack(side='left')
            if command:
                widget.bind('<<ComboboxSelected>>', command)

    # 创建所有输入行
    create_row(content, "选择 PDF 文件:", 'entry', vars["pdf_path"], browse_func=browse)
    create_row(content, "编号前缀:", 'entry', vars["prefix"])
    create_row(content, "字体:", 'combobox', vars["font"], options=["黑体", "宋体", "微软雅黑", "Times New Roman"])
    create_row(content, "字号 (pt):", 'entry', vars["fontsize"])

    # 颜色选择行（特殊处理）
    color_frame = ttk.Frame(content)
    color_frame.pack(fill='x', pady=5)
    ttk.Label(color_frame, text="颜色:", width=15, anchor='w').pack(side='left', padx=(0, 10))
    color_combo = ttk.Combobox(color_frame, textvariable=vars["color_name"],
                               values=["红色", "黑色", "蓝色", "绿色", "灰色"],
                               width=12, state='readonly')
    color_combo.pack(side='left')
    color_combo.bind('<<ComboboxSelected>>', update_color)
    ttk.Entry(color_frame, textvariable=vars["color"], width=10).pack(side='left', padx=(5, 0))

    create_row(content, "编号 X 坐标:", 'entry', vars["x"])
    create_row(content, "编号 Y 坐标:", 'entry', vars["y"])
    create_row(content, "渲染 DPI:", 'entry', vars["dpi"])
    create_row(content, "拼页行数:", 'entry', vars["rows"])
    create_row(content, "拼页列数:", 'entry', vars["cols"])
    create_row(content, "页面方向:", 'combobox', vars["direction"], options=["纵向", "横向"])

    # 编号页码范围输入（特殊处理，添加说明）
    page_frame = ttk.Frame(content)
    page_frame.pack(fill='x', pady=5)
    ttk.Label(page_frame, text="编号页码范围:", width=15, anchor='w').pack(side='left', padx=(0, 10))
    ttk.Entry(page_frame, textvariable=vars["page_range"], width=25).pack(side='left')

    # 添加页码范围说明
    help_frame = ttk.Frame(content)
    help_frame.pack(fill='x', pady=(0, 10))
    ttk.Label(help_frame, text="", width=15).pack(side='left', padx=(0, 10))  # 占位
    help_text = ttk.Label(help_frame, text="示例: 全部 | 1-5 | 1,3,5 | 1-3,6,8-10 (仅这些页添加编号)",
                          foreground='gray', font=('TkDefaultFont', 8))
    help_text.pack(side='left', anchor='w')

    create_row(content, "输出格式:", 'combobox', vars["output_format"], options=["PDF", "JPG"])

    # 进度条和按钮
    button_frame = ttk.Frame(content)
    button_frame.pack(fill='x', pady=20)

    progress = ttk.Progressbar(button_frame, orient="horizontal", mode="determinate")
    progress.pack(fill='x', pady=(0, 10))

    ttk.Button(button_frame, text="开始处理", command=start_process).pack()

    # 初始化颜色
    update_color()

    root.mainloop()


if __name__ == "__main__":
    run_gui()
