#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包便携版.py — 生成可拷到其它 Windows 电脑的「文档属性清理」发行包。

用法:
    python 打包便携版.py
    python 打包便携版.py --out "D:\\文档属性清理_便携版"

按白名单复制代码与说明，便携 Python 运行时一并带上（目标机无需安装 Python）。
绝不复制: _备份/ _日志/ 记录_*.md 外发前体检_*.md 旧版脚本 / 缓存 zip
—— 这些含原始作者名、本机路径、客户名，不能跟着工具走。

主脚本升级后重跑本脚本即可重新生成发行包。
"""

import argparse
import datetime
import os
import re
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, "便携版")
CACHE_ZIP = os.path.join(HERE, "_便携Python缓存_py314-embed.zip")
PY_URL = "https://www.python.org/ftp/python/3.14.3/python-3.14.3-embed-amd64.zip"

WHITELIST = [
    "clean_doc_meta.py",
    "install_menu.py",
    "selftest.py",
    "DROP_ON_ME.bat",
    "CHECK_ONLY.bat",
    "VIEW_LOG.bat",
    "INSTALL_MENU.bat",
    "UNINSTALL_MENU.bat",
    "使用说明.md",
]

README = """# 文档属性清理 · 便携版

> 生成时间：{{DATE}}　版本：v{{VERSION}}
> 把**整个文件夹**拷到任意 Windows 电脑，双击就能用：**不用装 Python、不用管理员、不用联网**。
> 包内自带 Python 3.14.3 官方 embeddable 运行时（约 12 MB，纯标准库）。

## 三步上手

1. 把整个文件夹拷到目标电脑（U 盘、NAS、任意路径均可；**路径含中文和空格没问题**）
2. 把要外发的文件/文件夹 **拖到 `DROP_ON_ME.bat`** 上，松手即清
3. 想先体检不改文件 → `CHECK_ONLY.bat`；想看每次清了什么 → `VIEW_LOG.bat`

⚠️ 顺序：**最后一次保存之后**再清。WPS 每次保存都会把用户名写回「最后保存者」。

## 换电脑后，哪些要重做

| 项目 | 换电脑后 | 说明 |
|---|---|---|
| 安装 Python / pip / 第三方包 | ❌ 不用 | 包内 `python\\` 自带运行时，脚本纯标准库 |
| 联网 | ❌ 不用 | 全流程离线 |
| 管理员权限 | ❌ 不用 | 拖放与右键菜单都只写 HKCU |
| 复制文件 | ✅ 要 | 拷整个文件夹即可 |
| **右键菜单** | ✅ **重新装一次** | 注册表里存的是**绝对路径**，换电脑/挪位置后必须重新双击 `INSTALL_MENU.bat`；卸载用 `UNINSTALL_MENU.bat` |
| 备份 / 日志 | 各机本地 | 运行时自动生成本文件夹下的 `_备份\\`、`_日志\\`，不跨机同步 |

## 常用命令（进阶）

用包内解释器，不依赖系统 Python：

```bat
python\\python.exe clean_doc_meta.py "D:\\待发目录\\" --dry-run --scan
python\\python.exe clean_doc_meta.py "D:\\待发目录\\" --deep --neutralize --scan
python\\python.exe clean_doc_meta.py --show-log
python\\python.exe selftest.py
```

参数说明见 `使用说明.md`（文档里的绝对路径是原电脑示例，**以本文件夹实际位置为准**）。

## 包内文件

```
文档属性清理_便携版/
├── README.md                 本文件
├── 使用说明.md               完整说明（清理清单 / 局限 / 验证方法）
├── clean_doc_meta.py         主脚本
├── install_menu.py           右键菜单安装 / 卸载（HKCU）
├── selftest.py               合成样本自检，不碰真实文件
├── DROP_ON_ME.bat            拖放即清理
├── CHECK_ONLY.bat            拖放只体检
├── VIEW_LOG.bat              打开最近一次日志
├── INSTALL_MENU.bat          安装右键菜单
├── UNINSTALL_MENU.bat        卸载右键菜单
└── python/                   官方 Python 3.14.3 embeddable（含 LICENSE.txt）
```

## 注意事项

- **本包不含任何备份、日志、体检报告或工作记录**，那些文件含真实作者名、本机路径、客户名，留在原电脑，不要一起拷走。
- 运行时生成的 `_备份\\` 和 `_日志\\` **含原始信息，绝对不要外发**；打包外发文件时排除它们。
- 目标电脑**已装 Python 3** 的话，可以直接删掉 `python\\` 目录，bat 会自动退回系统 `python` / `py -3`，包体只剩几十 KB。
- 这是 **64 位**运行时；若目标是 32 位 Windows，去 python.org 下 `python-3.14.3-embed-win32.zip` 解压覆盖 `python\\` 即可。
- 便携运行时来自 Python 官方，遵循 PSF License，许可文本见 `python\\LICENSE.txt`。
- 主脚本升级后，在原电脑运行 `python 打包便携版.py` 重新生成，再整包替换即可。
- 老 `.doc/.xls/.ppt/.wps` 仍不支持，需在 WPS 里另存为 `.docx/.xlsx/.pptx`；加密文档、带签名的 PDF 会跳过或提示风险。
"""

# 便携版里不该出现的真实信息（打包后再扫一遍，防止误带）
FORBIDDEN = []  # 例：["客户A", "项目B"] —— 填入你不希望随包外泄的关键词，留空则跳过该扫描


def main():
    ap = argparse.ArgumentParser(description="生成文档属性清理便携发行包")
    ap.add_argument("--out", default=DEFAULT_OUT, help="输出目录，默认当前目录下的 便携版")
    args = ap.parse_args()
    out = os.path.abspath(args.out)
    os.makedirs(out, exist_ok=True)

    print("输出目录：" + out)
    for name in WHITELIST:
        src = os.path.join(HERE, name)
        if not os.path.isfile(src):
            print("  ! 缺失，跳过：" + name)
            continue
        dst = os.path.join(out, name)
        if name == "使用说明.md":
            text = open(src, encoding="utf-8").read()
            text = re.sub(r"^> 位置：.*$", "> 位置：当前文件夹（便携版，可放在任意路径）",
                          text, count=1, flags=re.M)
            open(dst, "w", encoding="utf-8", newline="").write(text)
        else:
            shutil.copy2(src, dst)
        print("  + " + name)

    py_dir = os.path.join(out, "python")
    py_exe = os.path.join(py_dir, "python.exe")
    if os.path.isfile(py_exe):
        print("  = python\\ 便携运行时已存在，保留")
    elif os.path.isfile(CACHE_ZIP):
        with zipfile.ZipFile(CACHE_ZIP) as z:
            z.extractall(py_dir)
        print("  + 从缓存解压便携运行时")
    else:
        print("  ! 没有便携运行时：请下载 " + PY_URL + " 解压到 python\\，")
        print("    或让目标电脑自行安装 Python 3（bat 会自动退回系统 python）")

    readme = (README.replace("{{DATE}}", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
                    .replace("{{VERSION}}", detect_version()))
    open(os.path.join(out, "README.md"), "w", encoding="utf-8", newline="").write(readme)
    print("  + README.md")

    # 安全检查：包内不能出现真实姓名 / 客户名 / 本机业务路径
    leaked = []
    for root, dirs, files in os.walk(out):
        if os.path.basename(root) == "python":
            dirs[:] = []
            continue
        for f in files:
            p = os.path.join(root, f)
            if not f.lower().endswith((".py", ".bat", ".md", ".txt", ".log")):
                continue
            try:
                s = open(p, encoding="utf-8", errors="ignore").read()
            except OSError:
                continue
            for kw in FORBIDDEN:
                if kw in s:
                    leaked.append((os.path.relpath(p, out), kw))
    for dname in ("_备份", "_日志"):
        if os.path.isdir(os.path.join(out, dname)):
            print("  ! 警告：输出目录里有 %s\\，含敏感原件/日志，请勿随包拷贝" % dname)
    if leaked:
        print("  ! 警告：包内发现疑似真实信息，请人工核对：")
        for p, kw in leaked[:10]:
            print("      %s -> %s" % (p, kw))
    else:
        print("  ✓ 敏感信息扫描通过（包内无真实姓名 / 客户名 / 本机业务路径）")

    total = 0
    for root, _, files in os.walk(out):
        for f in files:
            total += os.path.getsize(os.path.join(root, f))
    print("\n完成，总大小 %.1f MB。把这个文件夹整个拷走即可。" % (total / 1024 / 1024))
    return 0


def detect_version():
    try:
        s = open(os.path.join(HERE, "clean_doc_meta.py"), encoding="utf-8").read()
        m = re.search(r'^VERSION\s*=\s*"([\d.]+)"', s, re.M)
        return m.group(1) if m else "?"
    except OSError:
        return "?"


if __name__ == "__main__":
    raise SystemExit(main())
