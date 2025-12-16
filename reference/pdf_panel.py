# ui/pdf_panel.py
"""
PDF处理面板
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
from threading import Thread


class PDFPanel:
    """PDF处理面板"""
    
    def __init__(self, parent, processor):
        self.processor = processor
        self.frame = ttk.Frame(parent)
        
        # 初始化变量
        self.processing = False
        
        # 创建UI
        self._create_widgets()
    
    def _create_widgets(self):
        """创建UI组件"""
        # 主容器
        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 顶部说明
        info_frame = ttk.LabelFrame(main_frame, text="功能说明", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        info_text = """
        📄 PDF处理功能：
        
        1. Word转PDF：将Nostamped_Word中的文档批量转换为PDF
        2. PDF合并：将多个PDF的指定页面合并为一个文件
        3. 页面提取：批量提取每个PDF的最后一页
        
        📁 目录结构：
        输入目录：resources/Nostamped_Word/
        中间目录：resources/Nostamped_PDF/
        输出目录：resources/Result_NoStamped/
        """
        
        info_label = ttk.Label(
            info_frame,
            text=info_text,
            justify=tk.LEFT
        )
        info_label.pack(anchor=tk.W)
        
        # 功能选择框架
        func_frame = ttk.Frame(main_frame)
        func_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧功能按钮
        btn_frame = ttk.Frame(func_frame)
        btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 功能按钮
        functions = [
            ("📄 Word转PDF", self.convert_word_to_pdf),
            ("🔗 合并PDF(自动)", lambda: self.merge_pdfs(False)),
            ("🎯 合并PDF(自定义)", lambda: self.merge_pdfs(True)),
            ("📑 提取最后一页", self.extract_last_pages),
            ("📊 文件统计", self.show_stats),
            ("🔄 刷新目录", self.refresh)
        ]
        
        for text, command in functions:
            btn = ttk.Button(
                btn_frame,
                text=text,
                command=command,
                width=20
            )
            btn.pack(pady=5)
        
        # 右侧日志区域
        log_frame = ttk.LabelFrame(func_frame, text="处理日志", padding="10")
        log_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            width=50,
            height=20,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 进度条
        progress_frame = ttk.Frame(log_frame)
        progress_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='indeterminate'
        )
        self.progress_bar.pack(fill=tk.X)
    
    def refresh(self):
        """刷新面板"""
        self._log("面板已刷新")
    
    def convert_word_to_pdf(self):
        """Word转PDF"""
        if self.processing:
            return
        
        # 确认对话框
        if not messagebox.askyesno("确认", "将转换Nostamped_Word中的所有Word文档为PDF，确定吗？"):
            return
        
        self._start_processing()
        
        def convert_thread():
            try:
                self.frame.after(0, lambda: self._log("开始转换Word文档..."))
                
                success, failed = self.processor.batch_convert_nostamped_word()
                
                self.frame.after(0, lambda: self._log(f"转换完成: 成功 {success} 个，失败 {failed} 个"))
                
                if success > 0:
                    self.frame.after(0, lambda: messagebox.showinfo(
                        "成功",
                        f"转换完成！\n成功: {success} 个\n失败: {failed} 个"
                    ))
                else:
                    self.frame.after(0, lambda: messagebox.showwarning(
                        "警告",
                        "未成功转换任何文档"
                    ))
                    
            except Exception as e:
                self.frame.after(0, lambda: self._log(f"转换失败: {e}"))
                self.frame.after(0, lambda: messagebox.showerror(
                    "错误",
                    f"转换失败: {str(e)}"
                ))
            
            self.frame.after(0, self._stop_processing)
        
        Thread(target=convert_thread, daemon=True).start()
    
    def merge_pdfs(self, custom_selection=False):
        """合并PDF"""
        if self.processing:
            return
        
        mode_text = "自定义选择" if custom_selection else "自动最后一页"
        
        if not messagebox.askyesno("确认", f"将合并PDF文件（{mode_text}），确定吗？"):
            return
        
        self._start_processing()
        
        def merge_thread():
            try:
                self.frame.after(0, lambda: self._log(f"开始合并PDF ({mode_text})..."))
                
                success, message = self.processor.merge_pdf_last_pages(custom_selection)
                
                if success:
                    self.frame.after(0, lambda: self._log(f"合并成功: {message}"))
                    self.frame.after(0, lambda: messagebox.showinfo("成功", message))
                else:
                    self.frame.after(0, lambda: self._log(f"合并失败: {message}"))
                    self.frame.after(0, lambda: messagebox.showerror("失败", message))
                    
            except Exception as e:
                self.frame.after(0, lambda: self._log(f"合并失败: {e}"))
                self.frame.after(0, lambda: messagebox.showerror(
                    "错误",
                    f"合并失败: {str(e)}"
                ))
            
            self.frame.after(0, self._stop_processing)
        
        Thread(target=merge_thread, daemon=True).start()
    
    def extract_last_pages(self):
        """提取最后一页"""
        if self.processing:
            return
        
        if not messagebox.askyesno("确认", "将提取所有PDF的最后一页为单独文件，确定吗？"):
            return
        
        self._start_processing()
        
        def extract_thread():
            try:
                self.frame.after(0, lambda: self._log("开始提取最后一页..."))
                
                success, failed = self.processor.batch_extract_all_last_pages()
                
                if success > 0:
                    self.frame.after(0, lambda: self._log(f"提取完成: 成功 {success} 个，失败 {failed} 个"))
                    self.frame.after(0, lambda: messagebox.showinfo(
                        "成功",
                        f"提取完成！\n成功: {success} 个\n失败: {failed} 个"
                    ))
                else:
                    self.frame.after(0, lambda: self._log("未提取到任何文件"))
                    self.frame.after(0, lambda: messagebox.showwarning(
                        "警告",
                        "未提取到任何文件"
                    ))
                    
            except Exception as e:
                self.frame.after(0, lambda: self._log(f"提取失败: {e}"))
                self.frame.after(0, lambda: messagebox.showerror(
                    "错误",
                    f"提取失败: {str(e)}"
                ))
            
            self.frame.after(0, self._stop_processing)
        
        Thread(target=extract_thread, daemon=True).start()
    
    def show_stats(self):
        """显示文件统计"""
        try:
            # 获取目录信息
            word_dir = self.processor.config.nostamped_word_dir
            pdf_dir = self.processor.config.nostamped_pdf_dir
            result_dir = self.processor.config.result_nostamped_dir
            
            stats_text = "📊 文件统计\n\n"
            
            # 统计Word文档
            if word_dir and os.path.exists(word_dir):
                word_files = [f for f in os.listdir(word_dir) if f.lower().endswith(('.docx', '.doc'))]
                stats_text += f"Word文档目录: {len(word_files)} 个文件\n"
            
            # 统计PDF文件
            if pdf_dir and os.path.exists(pdf_dir):
                pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
                stats_text += f"PDF中间目录: {len(pdf_files)} 个文件\n"
            
            # 统计结果文件
            if result_dir and os.path.exists(result_dir):
                result_files = []
                for root, dirs, files in os.walk(result_dir):
                    result_files.extend([os.path.join(root, f) for f in files if f.lower().endswith('.pdf')])
                stats_text += f"结果目录: {len(result_files)} 个PDF文件\n"
            
            messagebox.showinfo("文件统计", stats_text)
            self._log("已显示文件统计")
            
        except Exception as e:
            self._log(f"统计失败: {e}")
            messagebox.showerror("错误", f"统计失败: {str(e)}")
    
    def _start_processing(self):
        """开始处理状态"""
        self.processing = True
        self.progress_bar.start(10)
    
    def _stop_processing(self):
        """停止处理状态"""
        self.processing = False
        self.progress_bar.stop()
    
    def _log(self, message):
        """记录日志"""
        import time
        timestamp = time.strftime("[%H:%M:%S]")
        self.log_text.insert(tk.END, f"{timestamp} {message}\n")
        self.log_text.see(tk.END)
        self.log_text.update()