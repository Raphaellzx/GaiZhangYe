# GaiZhangYe/ui/batch_convert_panel.py
"""
批量Word转PDF面板 - 参考reference/stamp_panel.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import os
from threading import Thread


class BatchConvertPanel:
    """批量Word转PDF面板"""

    def __init__(self, parent, service):
        self.service = service
        self.frame = ttk.Frame(parent)

        # 初始化变量
        self.input_dir_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()

        # 创建UI
        self._create_widgets()

    def _create_widgets(self):
        """创建UI组件"""
        # 创建主容器
        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左侧配置区域
        config_frame = ttk.LabelFrame(main_frame, text="配置", padding="10")
        config_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))

        # 1. 输入目录选择
        input_frame = ttk.LabelFrame(config_frame, text="1. 输入目录", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # 目录选择
        ttk.Label(input_frame, text="待转换Word目录:").pack(anchor=tk.W, pady=(0, 5))

        input_dir_frame = ttk.Frame(input_frame)
        input_dir_frame.pack(fill=tk.X, pady=5)

        self.input_dir_entry = ttk.Entry(input_dir_frame, textvariable=self.input_dir_var, width=40)
        self.input_dir_entry.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            input_dir_frame,
            text="选择目录",
            command=self.select_input_dir,
            width=10
        ).pack(side=tk.LEFT)

        # 2. 输出目录选择
        output_frame = ttk.LabelFrame(config_frame, text="2. 输出目录", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        # 目录选择
        ttk.Label(output_frame, text="PDF输出目录:").pack(anchor=tk.W, pady=(0, 5))

        output_dir_frame = ttk.Frame(output_frame)
        output_dir_frame.pack(fill=tk.X, pady=5)

        self.output_dir_entry = ttk.Entry(output_dir_frame, textvariable=self.output_dir_var, width=40)
        self.output_dir_entry.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            output_dir_frame,
            text="选择目录",
            command=self.select_output_dir,
            width=10
        ).pack(side=tk.LEFT)

        # 3. 开始按钮
        self.start_btn = ttk.Button(
            config_frame,
            text="开始批量转换",
            command=self.start_batch_convert,
            width=20,
            style="Accent.TButton"
        )
        self.start_btn.pack(fill=tk.X, pady=10)

        # 右侧结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="操作结果", padding="10")
        result_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            height=25,
            font=('Consolas', 9)
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)

    def select_input_dir(self):
        """选择输入目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.input_dir_var.set(dir_path)

    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir_var.set(dir_path)

    def start_batch_convert(self):
        """开始批量Word转PDF"""
        self.result_text.delete(1.0, tk.END)
        self.start_btn.state(['disabled'])

        try:
            from pathlib import Path

            # 获取参数
            input_dir_path = self.input_dir_var.get()
            if not input_dir_path:
                self.result_text.insert(tk.END, "❌ 错误：请选择输入目录！\n")
                self.start_btn.state(['!disabled'])
                return

            output_dir_path = self.output_dir_var.get()
            if not output_dir_path:
                self.result_text.insert(tk.END, "❌ 错误：请选择输出目录！\n")
                self.start_btn.state(['!disabled'])
                return

            input_dir = Path(input_dir_path)
            output_dir = Path(output_dir_path)

            # 创建线程执行长时间任务
            def run_task():
                try:
                    self.result_text.insert(tk.END, "正在执行批量Word转PDF...\n")
                    self.result_text.insert(tk.END, f"输入目录：{input_dir}\n")
                    self.result_text.insert(tk.END, f"输出目录：{output_dir}\n\n")
                    self.result_text.see(tk.END)

                    # 执行服务
                    result_files = self.service.run(input_dir, output_dir)

                    self.result_text.insert(tk.END, f"✅ 批量转换完成！\n")
                    self.result_text.insert(tk.END, f"共转换 {len(result_files)} 个文件：\n")
                    for file in result_files:
                        self.result_text.insert(tk.END, f"  - {file}\n")
                    self.result_text.see(tk.END)

                except Exception as e:
                    import logging
                    logger = logging.getLogger("BatchConvertPanel")
                    logger.error(f"批量Word转PDF失败", exc_info=True)
                    self.result_text.insert(tk.END, f"❌ 失败：{str(e)}\n")
                    self.result_text.see(tk.END)
                finally:
                    self.start_btn.state(['!disabled'])

            # 启动线程
            thread = Thread(target=run_task)
            thread.daemon = True
            thread.start()

        except Exception as e:
            self.result_text.insert(tk.END, f"❌ 失败：{str(e)}\n")
            self.result_text.see(tk.END)
            self.start_btn.state(['!disabled'])

    def refresh(self):
        """刷新面板"""
        pass
