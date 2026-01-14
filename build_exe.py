#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 使用PyInstaller将GUI版本打包成exe
"""

import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller参数
PyInstaller.__main__.run([
    'improve_inventory_gui.py',  # 主脚本
    '--name=库存表管理系统',  # 输出文件名
    '--onefile',  # 打包成单个exe
    '--windowed',  # 不显示控制台窗口
    '--clean',  # 清理临时文件
    '--noconfirm',  # 覆盖输出目录
    # '--icon=icon.ico',  # 图标文件(如果有)
    '--add-data=使用说明.md;.',  # 包含使用说明
    '--distpath=dist',  # 输出目录
    '--workpath=build',  # 临时目录
    '--specpath=.',  # spec文件目录
])

print("\n" + "=" * 60)
print("打包完成!")
print("=" * 60)
print(f"\n可执行文件位置: {os.path.join(current_dir, 'dist', '库存表管理系统.exe')}")
