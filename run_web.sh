#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Root directory shortcut to run web service
"""

import sys
import os

# Add the project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from GaiZhangYe.entrypoints.start_web import main
    main()
except ImportError:
    from GaiZhangYe.entrypoints.start_web import app
    import threading
    import webbrowser
    import time

    def open_browser():
        time.sleep(2)
        webbrowser.open('http://localhost:5000')

    print("启动HTML可视化服务...")
    print("服务将在 http://localhost:5000 启动")
    print("正在自动打开浏览器...")

    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()

    app.run(
        host='0.0.0.0',
        port=5000,
        debug=False
    )
