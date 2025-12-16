# GaiZhangYe/ui/stamp_overlay_panel.py
"""
盖章页覆盖面板 - 参考reference/stamp_panel.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import os
from threading import Thread


class StampOverlayPanel:
    """盖章页覆盖面板"""

    def __init__(self, parent, service):
        self.service = service
        self.frame = ttk.Frame(parent)

        # 初始化变量
        self.stamp_file_var = tk.StringVar()
        self.image_width_var = tk.StringVar()
        self.target_word_dir_var = tk.StringVar()

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

        # 1. 盖章文件选择
        stamp_frame = ttk.LabelFrame(config_frame, text="1. 盖章文件选择", padding="10")
        stamp_frame.pack(fill=tk.X, pady=(0, 10))

        # 文件选择
        ttk.Label(stamp_frame, text="盖章PDF/图片:").pack(anchor=tk.W, pady=(0, 5))

        file_select_frame = ttk.Frame(stamp_frame)
        file_select_frame.pack(fill=tk.X, pady=5)

        self.stamp_file_entry = ttk.Entry(file_select_frame, textvariable=self.stamp_file_var, width=40)
        self.stamp_file_entry.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Button(
            file_select_frame,
            text="选择文件",
            command=self.select_stamp_file,
            width=10
        ).pack(side=tk.LEFT)

        # 2. 图片缩放设置
        scale_frame = ttk.LabelFrame(config_frame, text="2. 图片缩放设置", padding="10")
        scale_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(scale_frame, text="图片缩放宽度 (px):").pack(anchor=tk.W)

        scale_entry_frame = ttk.Frame(scale_frame)
        scale_entry_frame.pack(fill=tk.X, pady=5)

        from GaiZhangYe.utils.config import get_settings
        default_width = str(get_settings().image_default_width)
        self.image_width_var.set(default_width)

        ttk.Entry(
            scale_entry_frame,
            textvariable=self.image_width_var,
            width=15
        ).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(scale_frame, text="（0为保持原始尺寸）").pack(anchor=tk.W)

        # 3. 目标Word目录
        target_frame = ttk.LabelFrame(config_frame, text="3. 目标Word目录", padding="10")
        target_frame.pack(fill=tk.X, pady=(0, 10))

        # 目录选择
        ttk.Label(target_frame, text="目标Word文件目录:").pack(anchor=tk.W, pady=(0, 5))

        target_dir_frame = ttk.Frame(target_frame)
        target_dir_frame.pack(fill=tk.X, pady=5)

        self.target_word_dir_entry = ttk.Entry(target_dir_frame, textvariable=self.target_word_dir_var, width=40)
        self.target_word_dir_entry.pack(side=tk.LEFT, padx=(0, 5))

        from GaiZhangYe.utils.config import get_settings
        default_target_dir = str(get_settings().business_data_root / "func2" / "TargetFiles")
        self.target_word_dir_var.set(default_target_dir)

        ttk.Button(
            target_dir_frame,
            text="选择目录",
            command=self.select_target_word_dir,
            width=10
        ).pack(side=tk.LEFT)

        # 4. 开始按钮
        self.start_btn = ttk.Button(
            config_frame,
            text="开始覆盖",
            command=self.start_overlay,
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

    def select_stamp_file(self):
        """选择盖章文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF/Image Files", "*.pdf *.png *.jpg *.jpeg")]
        )
        if file_path:
            self.stamp_file_var.set(file_path)

    def select_target_word_dir(self):
        """选择目标Word目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.target_word_dir_var.set(dir_path)

    def start_overlay(self):
        """开始盖章页覆盖"""
        self.result_text.delete(1.0, tk.END)
        self.start_btn.state(['disabled'])

        try:
            from pathlib import Path

            # 获取参数
            stamp_file_path = self.stamp_file_var.get()
            if not stamp_file_path:
                self.result_text.insert(tk.END, "❌ 错误：请选择盖章文件！\n")
                self.start_btn.state(['!disabled'])
                return

            stamp_file = Path(stamp_file_path)
            target_word_dir = Path(self.target_word_dir_var.get())

            # 获取图片宽度
            image_width_str = self.image_width_var.get()
            image_width = None
            if image_width_str and image_width_str.isdigit():
                width = int(image_width_str)
                if width > 0:
                    image_width = width

            # 创建线程执行长时间任务
            def run_task():
                try:
                    self.result_text.insert(tk.END, "正在执行盖章页覆盖...\n")
                    self.result_text.insert(tk.END, f"盖章文件：{stamp_file}\n")
                    self.result_text.insert(tk.END, f"目标Word目录：{target_word_dir}\n")
                    if image_width:
                        self.result_text.insert(tk.END, f"图片缩放宽度：{image_width}px\n")
                    else:
                        self.result_text.insert(tk.END, f"图片缩放：保持原始尺寸\n")
                    self.result_text.insert(tk.END, "\n")
                    self.result_text.see(tk.END)

                    # 执行服务
                    result_files = self.service.run(stamp_file, target_word_dir, image_width)

                    self.result_text.insert(tk.END, f"✅ 盖章页覆盖完成！\n")
                    self.result_text.insert(tk.END, f"共处理 {len(result_files)} 个文件：\n")
                    for file in result_files:
                        self.result_text.insert(tk.END, f"  - {file}\n")
                    self.result_text.see(tk.END)

                except Exception as e:
                    import logging
                    logger = logging.getLogger("StampOverlayPanel")
                    logger.error(f"盖章页覆盖失败", exc_info=True)
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
