# ui/tools_panel.py
"""
工具与设置面板
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import sys
import subprocess


class ToolsPanel:
    """工具与设置面板"""
    
    def __init__(self, parent, processor):
        self.processor = processor
        self.frame = ttk.Frame(parent)
        
        # 创建UI
        self._create_widgets()
    
    def _create_widgets(self):
        """创建UI组件"""
        # 主容器
        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧工具按钮
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 文件夹工具
        folder_frame = ttk.LabelFrame(left_frame, text="📁 文件夹工具", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        folders = [
            ("打开文档目录", self.processor.open_target_folder),
            ("打开图片目录", self.processor.open_images_folder),
            ("打开PDF目录", self.processor.open_nostamped_pdf_folder),
            ("打开结果目录", self.processor.open_result_nostamped_folder),
            ("打开Word结果", self.processor.open_word_results_folder),
            ("打开PDF结果", self.processor.open_pdf_results_folder)
        ]
        
        for text, command in folders:
            btn = ttk.Button(
                folder_frame,
                text=text,
                command=command,
                width=15
            )
            btn.pack(pady=2)
        
        # 系统工具
        sys_frame = ttk.LabelFrame(left_frame, text="⚙️ 系统工具", padding="10")
        sys_frame.pack(fill=tk.X)
        
        sys_tools = [
            ("查看日志", self.view_logs),
            ("清理缓存", self.clean_cache),
            ("检查环境", self.check_environment),
            ("关于程序", self.show_about)
        ]
        
        for text, command in sys_tools:
            btn = ttk.Button(
                sys_frame,
                text=text,
                command=command,
                width=15
            )
            btn.pack(pady=2)
        
        # 右侧信息区域
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 系统信息
        sys_info_frame = ttk.LabelFrame(right_frame, text="🖥️ 系统信息", padding="10")
        sys_info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.sys_info_text = scrolledtext.ScrolledText(
            sys_info_frame,
            width=40,
            height=8,
            font=("Consolas", 9)
        )
        self.sys_info_text.pack(fill=tk.BOTH, expand=True)
        
        # 更新系统信息
        self._update_system_info()
        
        # 配置信息
        config_frame = ttk.LabelFrame(right_frame, text="⚙️ 配置信息", padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True)
        
        self.config_text = scrolledtext.ScrolledText(
            config_frame,
            width=40,
            height=8,
            font=("Consolas", 9)
        )
        self.config_text.pack(fill=tk.BOTH, expand=True)
        
        # 更新配置信息
        self._update_config_info()
    
    def _update_system_info(self):
        """更新系统信息"""
        try:
            info = []
            
            # Python信息
            info.append(f"Python版本: {sys.version.split()[0]}")
            info.append(f"操作系统: {sys.platform}")
            info.append(f"工作目录: {os.getcwd()}")
            info.append(f"脚本目录: {os.path.dirname(os.path.abspath(__file__))}")
            
            # 磁盘信息
            import shutil
            total, used, free = shutil.disk_usage(os.getcwd())
            info.append(f"磁盘空间: {free // (2**30)} GB 可用")
            
            # 内存信息
            try:
                import psutil
                memory = psutil.virtual_memory()
                info.append(f"内存使用: {memory.percent}%")
            except:
                info.append("内存信息: 未安装psutil")
            
            self.sys_info_text.insert(tk.END, "\n".join(info))
            self.sys_info_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.sys_info_text.insert(tk.END, f"获取系统信息失败: {e}")
            self.sys_info_text.config(state=tk.DISABLED)
    
    def _update_config_info(self):
        """更新配置信息"""
        try:
            info = []
            
            # 路径配置
            info.append("📁 路径配置:")
            for key in ['target_dir', 'images_dir', 'nostamped_word_dir', 
                       'nostamped_pdf_dir', 'word_results_dir', 'pdf_results_dir',
                       'result_nostamped_dir']:
                path = getattr(self.processor.config, key, None)
                if path:
                    # 显示相对路径
                    rel_path = os.path.relpath(path, os.getcwd()) if os.path.exists(path) else path
                    info.append(f"  {key}: {rel_path}")
            
            # 处理配置
            info.append("\n⚙️ 处理配置:")
            info.append(f"  默认页数: {self.processor.config.default_page_count}")
            info.append(f"  Word可见: {self.processor.config.getboolean('word', 'word_visible', False)}")
            info.append(f"  日志级别: {self.processor.config.get('logging', 'log_level', 'INFO')}")
            
            self.config_text.insert(tk.END, "\n".join(info))
            self.config_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.config_text.insert(tk.END, f"获取配置信息失败: {e}")
            self.config_text.config(state=tk.DISABLED)
    
    def view_logs(self):
        """查看日志"""
        try:
            log_file = self.processor.config.get('logging', 'log_file', fallback='logs/application.log')
            
            if os.path.exists(log_file):
                # 用默认文本编辑器打开
                if sys.platform == "win32":
                    os.startfile(log_file)
                elif sys.platform == "darwin":
                    subprocess.run(['open', log_file])
                else:
                    subprocess.run(['xdg-open', log_file])
                
                messagebox.showinfo("成功", f"已打开日志文件: {log_file}")
            else:
                messagebox.showwarning("警告", f"日志文件不存在: {log_file}")
                
        except Exception as e:
            messagebox.showerror("错误", f"打开日志失败: {str(e)}")
    
    def clean_cache(self):
        """清理缓存"""
        try:
            # 这里可以添加清理缓存的具体逻辑
            # 例如：删除临时文件、清理图片缓存等
            
            cache_dirs = [
                self.processor.config.images_dir,
                self.processor.config.nostamped_pdf_dir
            ]
            
            deleted_count = 0
            for cache_dir in cache_dirs:
                if cache_dir and os.path.exists(cache_dir):
                    for f in os.listdir(cache_dir):
                        if f.startswith('temp_') or f.startswith('cache_'):
                            try:
                                os.remove(os.path.join(cache_dir, f))
                                deleted_count += 1
                            except:
                                pass
            
            messagebox.showinfo("成功", f"已清理 {deleted_count} 个缓存文件")
            
        except Exception as e:
            messagebox.showerror("错误", f"清理缓存失败: {str(e)}")
    
    def check_environment(self):
        """检查运行环境"""
        try:
            checks = []
            
            # 检查Python版本
            python_ok = sys.version_info >= (3, 7)
            checks.append(f"Python版本: {'✅' if python_ok else '❌'}")
            
            # 检查必要模块
            modules = ['win32com', 'fitz', 'PIL']
            for module in modules:
                try:
                    __import__(module)
                    checks.append(f"模块 {module}: ✅")
                except ImportError:
                    checks.append(f"模块 {module}: ❌")
            
            # 检查目录权限
            test_dir = self.processor.config.target_dir
            if test_dir:
                try:
                    test_file = os.path.join(test_dir, '.test_write')
                    with open(test_file, 'w') as f:
                        f.write('test')
                    os.remove(test_file)
                    checks.append("目录权限: ✅")
                except:
                    checks.append("目录权限: ❌")
            
            result = "🔍 环境检查结果:\n\n" + "\n".join(checks)
            messagebox.showinfo("环境检查", result)
            
        except Exception as e:
            messagebox.showerror("错误", f"环境检查失败: {str(e)}")
    
    def show_about(self):
        """显示关于信息"""
        about_text = """
        🏷️ 盖章页覆盖工具 v3.0
        
        功能：
        - 📄 PDF图片提取
        - 🏷️ Word文档盖章
        - 🔄 Word转PDF
        - 🔗 PDF页面合并
        - 📑 PDF页面提取
        
        作者：自动生成
        版本：3.0 (GUI版本)
        日期：2024年
        
        技术栈：
        - Python 3.10+
        - Tkinter (GUI)
        - PyMuPDF (PDF处理)
        - pywin32 (Word操作)
        - Pillow (图片处理)
        """
        
        messagebox.showinfo("关于", about_text)