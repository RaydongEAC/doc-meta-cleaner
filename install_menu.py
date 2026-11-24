#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安装 / 卸载「资源管理器右键菜单」：彻底清除文档属性。

只写当前用户注册表 HKCU，不需要管理员权限，随时可卸载。
用法:
    python install_menu.py install     安装
    python install_menu.py uninstall  卸载
    python install_menu.py status     查看状态
"""

import os
import sys
import winreg

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "clean_doc_meta.py")
MENU_TEXT = "彻底清除文档属性(&C)"
KEYS = [
    r"Software\Classes\*\shell\CleanDocMeta",          # 文件
    r"Software\Classes\Directory\shell\CleanDocMeta",  # 目录
]


def build_command():
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = pyw if os.path.exists(pyw) else sys.executable
    # 用 format 而不是 % 格式化，避免 "%1" 被当成占位符
    # 右键点击即彻底清理：作者/版本 + 时间戳 + WPS 指纹 + 可疑残留扫描
    return '"{}" "{}" --deep --neutralize --scan --gui "%1"'.format(exe, SCRIPT)


def install():
    if not os.path.exists(SCRIPT):
        print("找不到 clean_doc_meta.py，先确认文件完整。")
        return 1
    cmd = build_command()
    for k in KEYS:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, k) as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, MENU_TEXT)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, k + r"\command") as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, cmd)
    print("已安装右键菜单：" + MENU_TEXT)
    print("调用命令：" + cmd)
    print("选中文件/文件夹后点右键即可；窗口可能需要重启资源管理器或重新登录才刷新。")
    return 0


def uninstall():
    for k in KEYS:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, k + r"\command")
        except FileNotFoundError:
            pass
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, k)
        except FileNotFoundError:
            pass
    print("已卸载右键菜单。")
    return 0


def status():
    for k in KEYS:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, k + r"\command") as key:
                v = winreg.QueryValueEx(key, None)[0]
            print("已安装 " + k + "\n  " + v)
        except FileNotFoundError:
            print("未安装 " + k)
    return 0


if __name__ == "__main__":
    action = sys.argv[1] if len(sys.argv) > 1 else "status"
    sys.exit({"install": install, "uninstall": uninstall, "status": status}.get(action, status)())
