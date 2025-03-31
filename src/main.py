# Copyright 2023 agenius666
# GitHub: https://github.com/agenius666/PDF-Index-Adder
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import os
import threading
import queue
import io
import re
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


class PDFProcessor(threading.Thread):
    def __init__(self, excel_path, output_dir):
        super().__init__()
        self.excel_path = excel_path
        self.output_dir = output_dir
        self.queue = queue.Queue()
        self._stop_event = threading.Event()

        # 注册中文字体（需要字体文件）
        pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))

    def run(self):
        try:
            excel_dir = os.path.dirname(os.path.abspath(self.excel_path))
            # 跳过标题行
            df = pd.read_excel(self.excel_path, header=None, usecols=[0, 1], dtype=str, skiprows=1)
            df.columns = ["path", "index"]

            total = len(df)
            processed = 0
            errors = []

            for i, row in df.iterrows():
                if self.stopped():
                    break

                # 处理文件路径
                raw_path = str(row["path"]).strip()
                if not raw_path:
                    errors.append(f"第{i + 2}行文件路径为空")
                    continue

                abs_path = os.path.join(excel_dir, raw_path) if not os.path.isabs(raw_path) else raw_path

                if not os.path.isfile(abs_path):
                    errors.append(f"第{i + 2}行文件不存在：{abs_path}")
                    continue

                # 处理索引号
                index = str(row["index"]).strip()
                if not index:
                    errors.append(f"第{i + 2}行索引号为空")
                    continue

                safe_index = re.sub(r'[\\/*?:"<>|]', '_', index)
                original_name = os.path.basename(abs_path)
                new_name = f"{safe_index}_{original_name}"
                output_path = os.path.join(self.output_dir, new_name)

                try:
                    self.add_index_to_pdf(abs_path, safe_index, output_path)
                    processed += 1
                except Exception as e:
                    errors.append(f"第{i + 2}行处理失败：{str(e)}")

                self.queue.put((i + 1, total))

            result_msg = f"成功处理 {processed}/{total} 个文件"
            if errors:
                error_log = os.path.join(self.output_dir, "error.log")
                with open(error_log, "w", encoding="utf-8") as f:
                    f.write("\n".join(errors))
                result_msg += f"\n失败 {len(errors)} 个，详见：{error_log}"

            self.queue.put(("success", result_msg))
        except Exception as e:
            self.queue.put(("error", f"致命错误：{str(e)}"))

    def add_index_to_pdf(self, input_path, index, output_path):
        """使用ReportLab添加文本层"""
        try:
            # 获取第一页尺寸
            with pdfplumber.open(input_path) as pdf:
                first_page = pdf.pages[0]
                page_width = first_page.width
                page_height = first_page.height

            # 创建文本层
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=(page_width, page_height))

            # 计算字体大小和位置
            font_size = page_height * 0.03
            right_margin = page_width * 0.03
            top_margin = page_height * 0.03

            # 设置中文字体
            c.setFont("SimSun", font_size)
            text_width = c.stringWidth(index, "SimSun", font_size)

            x = page_width - right_margin - text_width
            y = page_height - top_margin - font_size * 1.2  # 微调垂直位置

            c.drawString(x, y, index)
            c.save()

            # 合并文本层到原PDF
            packet.seek(0)
            text_pdf = PdfReader(packet)

            original_pdf = PdfReader(open(input_path, "rb"))
            output = PdfWriter()

            for page in original_pdf.pages:
                page.merge_page(text_pdf.pages[0])
                output.add_page(page)

            with open(output_path, "wb") as f:
                output.write(f)

        except Exception as e:
            raise RuntimeError(f"PDF处理失败：{os.path.basename(input_path)}\n{str(e)}")

    def stop(self):
        self._stop_event.set()

    def stopped(self):
        return self._stop_event.is_set()


class PDFIndexApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF索引号添加工具 - 1.0.0")
        self.processor = None

        if sys.platform == "win32":
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)

        # 配置自适应布局
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(4, weight=1)

        self.create_widgets()
        self.add_template_button()

    def create_widgets(self):
        # 文件选择部分（添加sticky参数和权重配置）
        ttk.Label(self.root, text="Excel文件路径:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.excel_entry = ttk.Entry(self.root)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Button(self.root, text="浏览", command=self.select_excel) \
            .grid(row=0, column=2, padx=10, pady=5, sticky="e")

        # 输出目录选择
        ttk.Label(self.root, text="输出目录:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ttk.Entry(self.root)
        self.output_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.root, text="浏览", command=self.select_output_dir) \
            .grid(row=1, column=2, padx=10, pady=5, sticky="e")

        # 进度条（扩展至全宽）
        self.progress = ttk.Progressbar(self.root, orient="horizontal", mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        # 控制按钮（居中显示）
        btn_frame = ttk.Frame(self.root)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10)
        self.start_btn = ttk.Button(btn_frame, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(btn_frame, text="停止处理", command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # 状态栏（底部对齐）
        self.status = ttk.Label(self.root, text="准备就绪", foreground="gray")
        self.status.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky="sw")

    def add_template_button(self):
        self.template_btn = ttk.Button(
            self.root,
            text="生成Excel模板",
            command=self.generate_template,
            style="Accent.TButton"
        )
        self.template_btn.grid(row=5, column=0, pady=15, padx=15, sticky="w")

        style = ttk.Style()
        style.configure("Accent.TButton", foreground="black", background="#0078D4")

    def generate_template(self):
        try:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                title="保存模板文件",
                initialfile="PDF索引模板.xlsx"
            )

            if not save_path:
                return

            sample_data = {
                "PDF文件路径（支持相对路径）": [
                    "files/报告.pdf",
                    "../财务文件/2023年报.pdf",
                    "C:/documents/重要文件.pdf"
                ],
                "索引号（支持任意文本）": [
                    "FIN-2023-001",
                    "财务部-0056",
                    "XM-2023-Q4"
                ]
            }

            df = pd.DataFrame(sample_data)
            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                worksheet.set_column('A:A', 50)
                worksheet.set_column('B:B', 30)

            messagebox.showinfo(
                "模板生成成功",
                "模板文件已生成，请注意：\n"
                "1. 第一列为PDF文件路径\n"
                "2. 支持绝对路径和相对路径\n"
                "3. 第二列可为任意文本索引号\n"
                "4. 需要中文字体文件（simsun.ttc）"
            )

        except Exception as e:
            messagebox.showerror("生成失败", f"错误信息：{str(e)}")

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)

    def select_output_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, path)
            os.makedirs(path, exist_ok=True)

    def start_processing(self):
        excel_path = self.excel_entry.get()
        output_dir = self.output_entry.get()

        if not os.path.isfile(excel_path):
            messagebox.showerror("错误", "请选择有效的Excel文件")
            return
        if not os.path.isdir(output_dir):
            messagebox.showerror("错误", "无法创建输出目录")
            return

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress["value"] = 0

        self.processor = PDFProcessor(excel_path, output_dir)
        self.processor.start()
        self.root.after(100, self.check_queue)

    def stop_processing(self):
        if self.processor:
            self.processor.stop()
            self.processor = None
            self.status.config(text="处理已中止")
            self.reset_ui()

    def check_queue(self):
        try:
            while True:
                message = self.processor.queue.get_nowait()
                if message[0] == "success":
                    self.show_result(message[1])
                    self.reset_ui()
                elif message[0] == "error":
                    messagebox.showerror("错误", message[1])
                    self.reset_ui()
                else:
                    current, total = message
                    self.progress["maximum"] = total
                    self.progress["value"] = current
                    self.status.config(text=f"正在处理 {current}/{total}...")
        except queue.Empty:
            pass

        if self.processor and self.processor.is_alive():
            self.root.after(100, self.check_queue)
        else:
            self.processor = None

    def show_result(self, message):
        result_window = tk.Toplevel(self.root)
        result_window.title("处理结果")

        text_frame = ttk.Frame(result_window)
        text_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        text = tk.Text(text_frame, wrap=tk.WORD, width=60, height=15)
        scrollbar = ttk.Scrollbar(text_frame, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)

        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text.insert(tk.END, message)
        text.configure(state=tk.DISABLED)

        ttk.Button(result_window, text="关闭", command=result_window.destroy).pack(pady=10)

    def reset_ui(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status.config(text="准备就绪")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("800x300")
    app = PDFIndexApp(root)
    root.mainloop()
