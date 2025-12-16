# ui/main_window.py
"""
主窗口模块
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import ConfigManager
from core.file_processor import FileProcessor
from .stamp_panel import StampPanel
from .pdf_panel import PDFPanel
from .tools_panel import ToolsPanel
from .styles import setup_styles


class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("盖章页覆盖工具 v3.0")
        
        # 设置窗口大小和位置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 1000
        window_height = 700
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(900, 600)
        
        # 设置图标
        self._set_icon()
        
        # 初始化配置和处理器
        self.config = ConfigManager()
        self.processor = FileProcessor()
        
        # 设置样式
        setup_styles()
        
        # 创建UI
        self._create_widgets()
        
        # 初始状态
        self._update_status()
        
        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _set_icon(self):
        """设置窗口图标"""
        try:
            # 尝试设置图标
            icon_path = os.path.join(os.path.dirname(__file__), "..", "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass
    
    def _create_widgets(self):
        """创建UI组件"""
        # 创建主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 顶部标题栏
        self._create_title_bar(main_container)
        
        # 创建主内容区域（选项卡）
        self._create_main_content(main_container)
        
        # 底部状态栏
        self._create_status_bar(main_container)
    
    def _create_title_bar(self, parent):
        """创建标题栏"""
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 标题
        title_label = ttk.Label(
            title_frame,
            text="📄 盖章页覆盖工具",
            font=("微软雅黑", 18, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(side=tk.LEFT)
        
        # 版本信息
        version_label = ttk.Label(
            title_frame,
            text="v3.0",
            font=("微软雅黑", 10),
            foreground="#7f8c8d"
        )
        version_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # 右侧操作按钮
        button_frame = ttk.Frame(title_frame)
        button_frame.pack(side=tk.RIGHT)
        
        # 刷新按钮
        refresh_btn = ttk.Button(
            button_frame,
            text="🔄 刷新",
            command=self._refresh,
            width=10
        )
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # 设置按钮
        settings_btn = ttk.Button(
            button_frame,
            text="⚙️ 设置",
            command=self._open_settings,
            width=10
        )
        settings_btn.pack(side=tk.LEFT, padx=5)
        
        # 帮助按钮
        help_btn = ttk.Button(
            button_frame,
            text="❓ 帮助",
            command=self._show_help,
            width=10
        )
        help_btn.pack(side=tk.LEFT)
    
    def _create_main_content(self, parent):
        """创建主内容区域"""
        # 创建选项卡控件
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 创建各个面板
        self.stamp_panel = StampPanel(self.notebook, self.processor)
        self.pdf_panel = PDFPanel(self.notebook, self.processor)
        self.tools_panel = ToolsPanel(self.notebook, self.processor)
        
        # 添加选项卡
        self.notebook.add(self.stamp_panel.frame, text="🏷️ 盖章处理")
        self.notebook.add(self.pdf_panel.frame, text="📄 PDF处理")
        self.notebook.add(self.tools_panel.frame, text="⚙️ 工具与设置")
        
        # 绑定选项卡切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
    
    def _create_status_bar(self, parent):
        """创建状态栏"""
        status_frame = ttk.Frame(parent, relief=tk.SUNKEN)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 左侧状态信息
        self.status_label = ttk.Label(
            status_frame,
            text="就绪",
            anchor=tk.W
        )
        self.status_label.pack(side=tk.LEFT, padx=10, pady=5)
        
        # 右侧系统信息
        sys_info = ttk.Label(
            status_frame,
            text=f"Python {sys.version.split()[0]}",
            anchor=tk.E
        )
        sys_info.pack(side=tk.RIGHT, padx=10, pady=5)
    
    def _update_status(self):
        """更新状态信息"""
        try:
            # 检查资源目录
            target_dir = self.config.target_dir
            images_dir = self.config.images_dir
            
            target_exists = os.path.exists(target_dir) if target_dir else False
            images_exists = os.path.exists(images_dir) if images_dir else False
            
            status_parts = []
            if target_exists:
                status_parts.append("文档目录正常")
            else:
                status_parts.append("文档目录缺失")
            
            if images_exists:
                status_parts.append("图片目录正常")
            else:
                status_parts.append("图片目录缺失")
            
            status_text = " | ".join(status_parts)
            self.status_label.config(text=status_text)
            
        except Exception as e:
            self.status_label.config(text=f"状态更新失败: {str(e)}")
    
    def _refresh(self):
        """刷新界面"""
        self._update_status()
        self.stamp_panel.refresh()
        self.pdf_panel.refresh()
        messagebox.showinfo("刷新", "界面已刷新")
    
    def _open_settings(self):
        """打开设置窗口"""
        # 这里可以添加设置窗口的实现
        messagebox.showinfo("设置", "设置功能开发中...")
    
    def _show_help(self):
        """显示帮助信息"""
        help_text = """
        📖 使用说明
        
        1. 盖章处理：
           - 将Word文档放入 resources/Target_Files/
           - 将盖章页PDF放入 resources/Images/（命名为'盖章页文件.pdf'）
           - 点击'提取图片'按钮
           - 配置插入页数
           - 点击'开始处理'
        
        2. PDF处理：
           - 将Word文档放入 resources/Nostamped_Word/
           - 点击'转换Word为PDF'
           - 选择合并或提取功能
        
        3. 工具与设置：
           - 打开相关文件夹
           - 查看文件信息
           - 管理缓存和日志
        """
        messagebox.showinfo("帮助", help_text)
    
    def _on_tab_changed(self, event):
        """选项卡切换事件"""
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        
        if "盖章处理" in current_tab:
            self.stamp_panel.refresh()
        elif "PDF处理" in current_tab:
            self.pdf_panel.refresh()
    
    def _on_closing(self):
        """窗口关闭事件"""
        if messagebox.askokcancel("退出", "确定要退出程序吗？"):
            self.root.quit()
            self.root.destroy()
    
    def run(self):
        """运行主循环"""
        self.root.mainloop()