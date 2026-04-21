"""主窗口GUI - 财务数据处理工具"""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from processors import DataMatcher, ExcelReader, ExcelWriter
from utils import setup_logger


class MainWindow:
    """主窗口类"""

    def __init__(self, root: tk.Tk):
        """初始化主窗口

        Args:
            root: Tkinter根窗口
        """
        self.root = root
        self.root.title("财务数据处理工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        # 初始化变量
        self.file_path = tk.StringVar()
        self.tolerance = tk.DoubleVar(value=10.0)
        self.is_processing = False

        # 设置日志
        self.logger = setup_logger('gui')

        # 创建界面
        self._create_widgets()

    def _create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # 文件路径显示
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        self.file_entry.grid(row=0, column=0, padx=(0, 10))

        # 浏览按钮
        browse_btn = ttk.Button(file_frame, text="浏览...", command=self._browse_file)
        browse_btn.grid(row=0, column=1)

        # 参数配置区域
        config_frame = ttk.LabelFrame(main_frame, text="参数配置", padding="10")
        config_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))

        # 容差值
        ttk.Label(config_frame, text="金额容差值（元）：").grid(row=0, column=0, sticky=tk.W)
        tolerance_spinbox = ttk.Spinbox(
            config_frame,
            textvariable=self.tolerance,
            from_=0,
            to=100,
            increment=1,
            width=10
        )
        tolerance_spinbox.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))

        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 15))

        # 开始处理按钮
        self.process_btn = ttk.Button(
            button_frame,
            text="开始处理",
            command=self._start_processing,
            width=20
        )
        self.process_btn.grid(row=0, column=0, padx=10)

        # 清空按钮
        clear_btn = ttk.Button(
            button_frame,
            text="清空",
            command=self._clear_form,
            width=20
        )
        clear_btn.grid(row=0, column=1, padx=10)

        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="10")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 日志文本框
        self.log_text = tk.Text(log_frame, height=10, width=60, state='disabled')
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 滚动条
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # 配置行和列的权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    def _browse_file(self):
        """浏览文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.logger.info(f"选择文件: {filename}")

    def _clear_form(self):
        """清空表单"""
        self.file_path.set("")
        self.tolerance.set(10.0)
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        self.logger.info("表单已清空")

    def _log_message(self, message: str, level: str = 'INFO'):
        """添加日志消息

        Args:
            message: 日志消息
            level: 日志级别
        """
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, f"[{level}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.root.update()

        # 同时记录到日志文件
        if level == 'INFO':
            self.logger.info(message)
        elif level == 'ERROR':
            self.logger.error(message)
        elif level == 'WARNING':
            self.logger.warning(message)

    def _start_processing(self):
        """开始处理文件"""
        if self.is_processing:
            messagebox.showwarning("警告", "处理正在进行中，请稍候...")
            return

        # 验证输入
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("错误", "请选择要处理的Excel文件")
            return

        file_path = Path(file_path)
        if not file_path.exists():
            messagebox.showerror("错误", "文件不存在")
            return

        # 开始处理
        self.is_processing = True
        self.process_btn.config(state='disabled')
        self._log_message("开始处理文件...")

        try:
            # 读取Excel文件
            self._log_message("正在读取Excel文件...")
            reader = ExcelReader(file_path)
            sheet1_records, sheet2_records = reader.read_excel()
            self._log_message(f"读取成功: {len(sheet1_records)}条组织记录, {len(sheet2_records)}条明细记录")

            # 匹配数据
            self._log_message("正在匹配数据...")
            matcher = DataMatcher(tolerance=self.tolerance.get())
            matched_records = matcher.match_data(sheet1_records, sheet2_records)
            self._log_message(f"匹配完成: {len(matched_records)}条记录")

            # 统计匹配结果
            matched_count = sum(1 for r in matched_records if r.organization)
            self._log_message(f"已归类记录: {matched_count}条")

            # 写入结果
            self._log_message("正在写入结果文件...")
            writer = ExcelWriter()
            output_file = writer.write_results(file_path, matched_records)
            self._log_message(f"处理完成！输出文件: {output_file.name}")

            # 显示完成消息
            messagebox.showinfo("完成", f"文件处理完成！\n输出文件: {output_file.name}")

        except Exception as e:
            error_msg = str(e)
            self._log_message(f"处理失败: {error_msg}", 'ERROR')
            messagebox.showerror("错误", f"处理失败:\n{error_msg}")
            self.logger.exception("处理异常")

        finally:
            self.is_processing = False
            self.process_btn.config(state='normal')
            self._log_message("处理结束")


def run_gui():
    """运行GUI应用"""
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == '__main__':
    run_gui()
