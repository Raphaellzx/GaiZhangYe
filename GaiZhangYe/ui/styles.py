# ui/styles.py
"""
样式定义模块
"""

import tkinter as tk
from tkinter import ttk


def setup_styles():
    """设置应用程序样式"""
    style = ttk.Style()
    
    # 使用clam主题（更现代）
    style.theme_use('clam')
    
    # 基础样式
    style.configure('TFrame', background='#f5f5f5')
    style.configure('TLabel', background='#f5f5f5', font=('微软雅黑', 9))
    style.configure('TButton', font=('微软雅黑', 9), padding=6)
    style.configure('TLabelframe', background='#f5f5f5', font=('微软雅黑', 10, 'bold'))
    style.configure('TLabelframe.Label', background='#f5f5f5', font=('微软雅黑', 10, 'bold'))
    
    # 强调按钮样式
    style.configure('Accent.TButton', 
                   font=('微软雅黑', 10, 'bold'),
                   background='#3498db',
                   foreground='white',
                   padding=10)
    
    # 进度条样式
    style.configure('TProgressbar', 
                   background='#3498db',
                   troughcolor='#ecf0f1',
                   bordercolor='#bdc3c7',
                   lightcolor='#3498db',
                   darkcolor='#2980b9')
    
    # 选项卡样式
    style.configure('TNotebook', background='#ecf0f1', borderwidth=0)
    style.configure('TNotebook.Tab', 
                   padding=[20, 5],
                   font=('微软雅黑', 10),
                   background='#bdc3c7',
                   foreground='#2c3e50')
    style.map('TNotebook.Tab',
             background=[('selected', '#3498db'), ('active', '#2980b9')],
             foreground=[('selected', 'white'), ('active', 'white')])
    
    # 组合框样式
    style.configure('TCombobox', 
                   fieldbackground='white',
                   background='white',
                   arrowsize=12)
    
    # 滚动条样式
    style.configure('Vertical.TScrollbar',
                   background='#bdc3c7',
                   troughcolor='#ecf0f1',
                   bordercolor='#bdc3c7',
                   arrowcolor='#2c3e50')
    
    style.configure('Horizontal.TScrollbar',
                   background='#bdc3c7',
                   troughcolor='#ecf0f1',
                   bordercolor='#bdc3c7',
                   arrowcolor='#2c3e50')
    
    # 状态标签样式
    style.configure('Success.TLabel', foreground='#27ae60')
    style.configure('Error.TLabel', foreground='#e74c3c')
    style.configure('Warning.TLabel', foreground='#f39c12')
    
    # 复选框样式
    style.configure('TCheckbutton', background='#f5f5f5')
    
    # 微调框样式
    style.configure('TSpinbox',
                   fieldbackground='white',
                   background='white',
                   arrowsize=12)