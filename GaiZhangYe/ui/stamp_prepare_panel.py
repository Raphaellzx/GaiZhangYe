# GaiZhangYe/ui/stamp_prepare_panel.py
"""
准备盖章页面板 - 参考reference/stamp_panel.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import os
from threading import Thread


class StampPreparePanel:
    """准备盖章页面板"""

    def __init__(self, parent, service):
        self.service = service
        self.frame = ttk.Frame(parent)

        # 初始化变量
        self.word_dir_var = tk.StringVar()
        self.page_var = tk.StringVar(value="1")

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

        # 1. 未盖章Word目录
        word_frame = ttk.LabelFrame(config_frame, text="1. 未盖章Word目录", padding="10")
        word_frame.pack(fill=tk.X, pady=(0, 10))

        # 目录选择
        ttk.Label(word_frame, text="Word文件目录:").pack(anchor=tk.W, pady=(0, 5))

        dir_select_frame = ttk.Frame(word_frame)
        dir_select_frame.pack(fill=tk.X, pady=5)

        self.word_dir_entry = ttk.Entry(dir_select_frame, textvariable=self.word_dir_var, width=40)
        self.word_dir_entry.pack(side=tk.LEFT, padx=(0, 5))

        from GaiZhangYe.utils.config import get_settings
        default_word_dir = str(get_settings().business_data_root / "func1" / "Nostamped_Word")
        self.word_dir_var.set(default_word_dir)

        ttk.Button(
            dir_select_frame,
            text="选择目录",
            command=self.select_word_dir,
            width=10
        ).pack(side=tk.LEFT)

        # 2. 提取页码设置
        page_frame = ttk.LabelFrame(config_frame, text="2. 提取页码设置", padding="10")
        page_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(page_frame, text="提取页码 (逗号分隔):").pack(anchor=tk.W)

        page_entry_frame = ttk.Frame(page_frame)
        page_entry_frame.pack(fill=tk.X, pady=5)

        ttk.Entry(
            page_entry_frame,
            textvariable=self.page_var,
            width=20
        ).pack(side=tk.LEFT, padx=(0, 5))

        # 3. 开始按钮
        self.start_btn = ttk.Button(
            config_frame,
            text="开始准备",
            command=self.start_prepare,
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

    def select_word_dir(self):
        """选择Word目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.word_dir_var.set(dir_path)

    def start_prepare(self):
        """开始准备盖章页"""
        self.result_text.delete(1.0, tk.END)
        self.start_btn.state(['disabled'])

        try:
            from pathlib import Path

            # 获取参数
            word_dir = Path(self.word_dir_var.get())
            page_str = self.page_var.get()
            target_pages = [int(p.strip()) for p in page_str.split(",") if p.strip().isdigit()]

            if not target_pages:
                self.result_text.insert(tk.END, "❌ 错误：请输入有效的页码！\n")
                return

            # 创建线程执行长时间任务
            def run_task():
                try:
                    self.result_text.insert(tk.END, "正在准备盖章页...\n")
                    self.result_text.insert(tk.END, f"Word目录：{word_dir}\n")
                    self.result_text.insert(tk.END, f"提取页码：{target_pages}\n\n")

                    # 执行服务
                    result_pdf = self.service.run(target_pages, word_dir)

                    self.result_text.insert(tk.END, f"✅ 准备完成！\n")
                    self.result_text.insert(tk.END, f"生成的盖章页：{result_pdf}\n")

                except Exception as e:
                    import logging
                    logger = logging.getLogger("StampPreparePanel")
                    logger.error(f"准备盖章页失败", exc_info=True)
                    self.result_text.insert(tk.END, f"❌ 失败：{str(e)}\n")
                finally:
                    self.start_btn.state(['!disabled'])

            # 启动线程
            thread = Thread(target=run_task)
            thread.daemon = True
            thread.start()

        except Exception as e:
            self.result_text.insert(tk.END, f"❌ 失败：{str(e)}\n")
            self.start_btn.state(['!disabled'])

    def refresh(self):
        """刷新面板"""
        pass
