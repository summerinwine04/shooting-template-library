#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ASCII 入口：供 启动.bat 双击调用，转去运行 工具/启动模板库.py（避免 bat 里出现中文路径）。"""
import os, sys, importlib.util

here = os.path.dirname(os.path.abspath(__file__))
target = os.path.join(here, "工具", "启动模板库.py")
if not os.path.exists(target):
    print("找不到 工具/启动模板库.py，请确认 serve.py 放在 视觉模板库 根目录。")
    sys.exit(1)
spec = importlib.util.spec_from_file_location("tpl_server", target)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
m.main()
