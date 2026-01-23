#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Minimal Flask app runner. Routes are provided by blueprints in web.routes."""
import sys
from pathlib import Path
import logging
from GaiZhangYe.web import create_app

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ensure project root is on sys.path for direct execution
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent))

app = create_app()

if __name__ == "__main__":
    logger.info("启动HTML可视化服务...")
    logger.info("服务将在 http://localhost:5001 启动")
    logger.info("按 Ctrl+C 停止服务")
    app.run(host="0.0.0.0", port=5001, debug=False)
