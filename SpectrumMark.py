#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# --------------------- 辅助函数 ---------------------
def is_chinese(char):
    """
    判断单个字符是否为中文（常见汉字 Unicode 区间 \u0800 ~ \uFFFF）
    """
    return "\u0800" <= char <= "\uFFFF"


def segment_text(text_line):
    """
    将文本行分割为连续同类型字符段（中文或非中文）列表
    返回形式：[(segment, is_chinese), ...]
    """
    segments = []
    if text_line:
        current_seg = text_line[0]
        current_is_ch = is_chinese(text_line[0])
        for ch in text_line[1:]:
            if is_chinese(ch) == current_is_ch:
                current_seg += ch
            else:
                segments.append((current_seg, current_is_ch))
                current_seg = ch
                current_is_ch = is_chinese(ch)
        segments.append((current_seg, current_is_ch))
    return segments


def compute_line_width(text_line, chinese_font, latin_font, font_size, char_adjust=-0.5):
    """
    计算整行文本的宽度，区分中文和英文部分
    """
    total_width = 0.0
    segments = segment_text(text_line)
    for seg, seg_is_ch in segments:
        if seg_is_ch:
            seg_width = chinese_font.text_length(seg, fontsize=font_size)
            total_width += seg_width
        else:
            for letter in seg:
                letter_width = latin_font.text_length(letter, fontsize=font_size)
                total_width += (letter_width + char_adjust)
    return total_width


def split_text_line(text_line, allowed_width, chinese_font, latin_font, font_size, char_adjust=-0.5):
    """
    将文本拆分成两部分，使得第一部分文字宽度不超过 allowed_width
    返回 (first_part, second_part)
    """
    cum_width = 0.0
    split_index = 0
    for i, ch in enumerate(text_line):
        if is_chinese(ch):
            ch_width = chinese_font.text_length(ch, fontsize=font_size)
        else:
            ch_width = latin_font.text_length(ch, fontsize=font_size) + char_adjust
        if cum_width + ch_width <= allowed_width:
            cum_width += ch_width
            split_index = i + 1
        else:
            break
    if split_index == 0:
        split_index = 1
    first_part = text_line[:split_index]
    second_part = text_line[split_index:]
    return first_part, second_part


# --------------------- PDF 文本插入核心函数 ---------------------
def perform_insertion(input_pdf, overlay_txt, output_pdf, insert_x, insert_y,
                      font_size, color_str, text_align, log_func):
    """
    执行 PDF 文本插入操作：
      - input_pdf：输入 PDF 文件路径
      - overlay_txt：待插入文本文件路径（每行对应一页）
      - output_pdf：输出 PDF 文件路径
      - insert_x, insert_y, font_size, color_str, text_align 为各项参数（均为字符串，由 GUI 中获取）
      - log_func：日志打印回调，可用于在 GUI 界面输出日志信息
    """
    try:
        insert_x = float(insert_x)
        insert_y = float(insert_y)
        font_size = float(font_size)
    except ValueError:
        log_func("错误：插入 X/Y 坐标和字体大小必须为数字格式！")
        return

    text_align = text_align.lower().strip()
    if text_align not in ("left", "center", "right"):
        log_func("警告：文字对齐方式无效，采用默认 left 对齐")
        text_align = "left"

    try:
        # 支持“0,0,0”格式的 RGB 数值（0～255 或 0～1）
        color_values = [float(c.strip()) for c in color_str.split(",")]
        if len(color_values) != 3:
            raise ValueError
        if any(c > 1 for c in color_values):
            color_black = tuple(c / 255.0 for c in color_values)
        else:
            color_black = tuple(color_values)
    except Exception:
        log_func("配置中颜色参数有误，采用默认黑色 (0,0,0)")
        color_black = (0, 0, 0)

    if not os.path.exists(overlay_txt):
        log_func("未找到文本文件: " + overlay_txt)
        return

    try:
        with open(overlay_txt, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip()]
    except Exception as e:
        log_func("读取文本文件失败: " + str(e))
        return

    try:
        doc = fitz.open(input_pdf)
    except Exception as e:
        log_func("打开 PDF 文件失败: " + str(e))
        return

    # 定义字体（中文用内置 "china-s"，英文用 "Times-Roman"）
    chinese_font_name = "china-s"
    latin_font_name = "Times-Roman"
    try:
        chinese_font_obj = fitz.Font(chinese_font_name)
    except Exception as e:
        log_func(f"加载中文字体 '{chinese_font_name}' 失败: " + str(e))
        return
    latin_font_obj = fitz.Font(latin_font_name)
    char_adjust = -0.5

    num_lines = len(lines)
    num_pages = doc.page_count
    count = min(num_lines, num_pages)
    toc = []  # 用于存储书签数据

    for i in range(count):
        text_line = lines[i]
        page = doc[i]
        page_height = page.rect.height
        page_width = page.rect.width
        adjusted_y = page_height - insert_y

        # 计算文本宽度与允许的最大区域
        line_width = compute_line_width(text_line, chinese_font_obj, latin_font_obj, font_size, char_adjust)
        allowed_width = page_width - 2 * insert_x

        log_func(f"在第 {i+1} 页插入文本: {text_line}")
        log_func(f"预计算整行宽度: {line_width}，允许最大宽度: {allowed_width}，对齐方式: {text_align}")

        # 如果文本超过允许宽度，则拆分成两行
        if line_width > allowed_width:
            first_part, second_part = split_text_line(text_line, allowed_width, chinese_font_obj, latin_font_obj, font_size, char_adjust)
            first_line_width = compute_line_width(first_part, chinese_font_obj, latin_font_obj, font_size, char_adjust)
            second_line_width = compute_line_width(second_part, chinese_font_obj, latin_font_obj, font_size, char_adjust)

            if text_align == "left":
                x1 = insert_x
                x2 = insert_x
            elif text_align == "center":
                x1 = (page_width - first_line_width) / 2
                x2 = (page_width - second_line_width) / 2
            elif text_align == "right":
                x1 = page_width - insert_x - first_line_width
                x2 = page_width - insert_x - second_line_width

            first_line_y = adjusted_y - font_size - 2
            second_line_y = adjusted_y

            log_func(f"文本过长，拆为两行插入：\n  第一行: '{first_part}'，起始坐标: ({x1}, {first_line_y})\n  第二行: '{second_part}'，起始坐标: ({x2}, {second_line_y})")

            # 插入第一行
            segments_first = segment_text(first_part)
            current_x = x1
            for seg, seg_is_ch in segments_first:
                if seg_is_ch:
                    page.insert_text(
                        (current_x, first_line_y),
                        seg,
                        fontname=chinese_font_name,
                        fontsize=font_size,
                        color=color_black,
                        overlay=True
                    )
                    seg_width = chinese_font_obj.text_length(seg, fontsize=font_size)
                    current_x += seg_width
                else:
                    for letter in seg:
                        page.insert_text(
                            (current_x, first_line_y),
                            letter,
                            fontname=latin_font_name,
                            fontsize=font_size,
                            color=color_black,
                            overlay=True
                        )
                        letter_width = latin_font_obj.text_length(letter, fontsize=font_size)
                        current_x += (letter_width + char_adjust)

            # 插入第二行
            segments_second = segment_text(second_part)
            current_x = x2
            for seg, seg_is_ch in segments_second:
                if seg_is_ch:
                    page.insert_text(
                        (current_x, second_line_y),
                        seg,
                        fontname=chinese_font_name,
                        fontsize=font_size,
                        color=color_black,
                        overlay=True
                    )
                    seg_width = chinese_font_obj.text_length(seg, fontsize=font_size)
                    current_x += seg_width
                else:
                    for letter in seg:
                        page.insert_text(
                            (current_x, second_line_y),
                            letter,
                            fontname=latin_font_name,
                            fontsize=font_size,
                            color=color_black,
                            overlay=True
                        )
                        letter_width = latin_font_obj.text_length(letter, fontsize=font_size)
                        current_x += (letter_width + char_adjust)
        else:
            # 单行插入
            if text_align == "left":
                current_x = insert_x
            elif text_align == "center":
                current_x = (page_width - line_width) / 2
            elif text_align == "right":
                current_x = page_width - insert_x - line_width

            log_func("单行插入，起始 X 坐标: " + str(current_x))
            segments = segment_text(text_line)
            for seg, seg_is_ch in segments:
                if seg_is_ch:
                    page.insert_text(
                        (current_x, adjusted_y),
                        seg,
                        fontname=chinese_font_name,
                        fontsize=font_size,
                        color=color_black,
                        overlay=True
                    )
                    seg_width = chinese_font_obj.text_length(seg, fontsize=font_size)
                    log_func(f"中文段: '{seg}' 宽度: {seg_width}")
                    current_x += seg_width
                else:
                    for letter in seg:
                        page.insert_text(
                            (current_x, adjusted_y),
                            letter,
                            fontname=latin_font_name,
                            fontsize=font_size,
                            color=color_black,
                            overlay=True
                        )
                        letter_width = latin_font_obj.text_length(letter, fontsize=font_size)
                        log_func(f"英文字符: '{letter}' 宽度: {letter_width}")
                        current_x += (letter_width + char_adjust)
        toc.append([1, text_line, i + 1])

# 添加书签信息（页面编号从 1 开始）
    doc.set_toc(toc)
    for level, title, page in toc:
        log_func(f"生成书签: 层级={level}, 标题='{title}', 页面={page}")
    
    try:
        doc.save(output_pdf)
        log_func("文本成功插入对应页，输出文件为：" + output_pdf)
    except Exception as e:
        log_func("保存输出 PDF 文件失败: " + str(e))
    doc.close()


# --------------------- GUI 界面 ---------------------
class PDFInsertionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 文本插入工具        by unruffle")
        self.geometry("600x550")
        self.create_widgets()

    def create_widgets(self):
        # 框架：参数设置区域
        frame = tk.Frame(self)
        frame.pack(padx=10, pady=10, fill="x")

        # 输入 PDF 文件
        tk.Label(frame, text="输入 PDF 文件：").grid(row=0, column=0, sticky="w")
        self.input_pdf_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.input_pdf_var, width=50).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="浏览...", command=self.browse_input_pdf).grid(row=0, column=2, padx=5)

        # 文本文件（待插入文本）
        tk.Label(frame, text="插入文本文件：").grid(row=1, column=0, sticky="w")
        self.overlay_txt_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.overlay_txt_var, width=50).grid(row=1, column=1, padx=5)
        tk.Button(frame, text="浏览...", command=self.browse_overlay_txt).grid(row=1, column=2, padx=5)

        # 输出 PDF 文件
        tk.Label(frame, text="输出 PDF 文件：").grid(row=2, column=0, sticky="w")
        self.output_pdf_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.output_pdf_var, width=50).grid(row=2, column=1, padx=5)
        tk.Button(frame, text="浏览...", command=self.browse_output_pdf).grid(row=2, column=2, padx=5)

        # 插入 X 坐标
        tk.Label(frame, text="X 坐标：").grid(row=3, column=0, sticky="w")
        self.insert_x_var = tk.StringVar(value="60")
        tk.Entry(frame, textvariable=self.insert_x_var, width=20).grid(row=3, column=1, sticky="w", padx=5)

        # 插入 Y 坐标
        tk.Label(frame, text="Y 坐标：").grid(row=4, column=0, sticky="w")
        self.insert_y_var = tk.StringVar(value="8")
        tk.Entry(frame, textvariable=self.insert_y_var, width=20).grid(row=4, column=1, sticky="w", padx=5)

        # 字体大小
        tk.Label(frame, text="字 号：").grid(row=5, column=0, sticky="w")
        self.font_size_var = tk.StringVar(value="12")
        tk.Entry(frame, textvariable=self.font_size_var, width=20).grid(row=5, column=1, sticky="w", padx=5)

        # 颜色配置
        tk.Label(frame, text="文本颜色 (R,G,B)：").grid(row=6, column=0, sticky="w")
        self.color_black_var = tk.StringVar(value="0,0,0")
        tk.Entry(frame, textvariable=self.color_black_var, width=20).grid(row=6, column=1, sticky="w", padx=5)

        # 文本对齐方式
        tk.Label(frame, text="对齐方式：").grid(row=7, column=0, sticky="w")
        self.text_align_var = tk.StringVar(value="center")
        tk.OptionMenu(frame, self.text_align_var, "left", "center", "right").grid(row=7, column=1, sticky="w", padx=5)

        # “开始执行”按钮
        tk.Button(self, text="开始执行", command=self.start_process, width=20).pack(pady=10)

        # 日志输出文本框
        tk.Label(self, text="日志信息：").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self, width=70, height=15, state="normal")
        self.log_text.pack(padx=10, pady=5, fill="both", expand=True)

    def browse_input_pdf(self):
        filename = filedialog.askopenfilename(title="选择输入 PDF 文件", filetypes=[("PDF 文件", "*.pdf")])
        if filename:
            self.input_pdf_var.set(filename)

    def browse_overlay_txt(self):
        filename = filedialog.askopenfilename(title="选择文本文件", filetypes=[("文本文件", "*.txt")])
        if filename:
            self.overlay_txt_var.set(filename)

    def browse_output_pdf(self):
        filename = filedialog.asksaveasfilename(title="选择输出 PDF 文件", defaultextension=".pdf", filetypes=[("PDF 文件", "*.pdf")])
        if filename:
            self.output_pdf_var.set(filename)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def start_process(self):
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        input_pdf = self.input_pdf_var.get().strip()
        overlay_txt = self.overlay_txt_var.get().strip()
        output_pdf = self.output_pdf_var.get().strip()

        if not input_pdf or not overlay_txt or not output_pdf:
            messagebox.showerror("错误", "请确保选择输入 PDF、文本文件和输出 PDF 文件！")
            return

        # 获取各个参数
        insert_x = self.insert_x_var.get().strip()
        insert_y = self.insert_y_var.get().strip()
        font_size = self.font_size_var.get().strip()
        color_black = self.color_black_var.get().strip()
        text_align = self.text_align_var.get().strip()

        self.log("开始处理……")
        try:
            perform_insertion(input_pdf, overlay_txt, output_pdf,
                              insert_x, insert_y, font_size, color_black, text_align,
                              self.log)
            self.log("处理完成！")
        except Exception as e:
            self.log("出现异常：" + str(e))
            messagebox.showerror("错误", str(e))


if __name__ == "__main__":
    app = PDFInsertionApp()
    app.mainloop()