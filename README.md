# 文档属性清理工具

卸载 MS Office、只用 WPS 的电脑上，资源管理器「属性 → 详细信息」读不到也删不掉作者/版本信息。本工具绕开 Windows 界面，**直接清除文档内部的元数据**，纯 Python 标准库，不依赖 Office/WPS，离线可用。

## 支持

Word / Excel / PowerPoint（OOXML：docx/xlsx/pptx 等）、PDF。老二进制格式（.doc/.xls/.ppt/.wps）需先在 WPS 另存为新版格式；加密文档与带签名 PDF 会跳过或提示风险。

清除范围：作者、最后保存者、主题/关键词/备注、版本号与公司/管理者、WPS 账号与设备指纹（自定义属性、Application 版本 GUID、主题名、WPS 命名空间）、Excel 本机绝对路径 `x15ac:absPath`、批注/修订作者；PDF 的 `/Info` 与 XMP（含 FlateDecode 压缩流）。

## 快速使用

三种入口，效果一致（均为彻底清理：作者/版本 + 时间戳 + WPS 指纹 + 残留扫描）：

- **右键菜单**：双击 `INSTALL_MENU.bat` 安装一次，之后在任意文件/文件夹上右键 → 「彻底清除文档属性」。只写 HKCU，不需要管理员；卸载用 `UNINSTALL_MENU.bat`。
- **拖放**：把文件或文件夹拖到 `DROP_ON_ME.bat` 上。
- **命令行**：`python clean_doc_meta.py "<路径>" --deep --neutralize --scan`（命令行默认只清基础项，彻底清理需显式加这三个参数）。

`CHECK_ONLY.bat` 只体检不改动；`VIEW_LOG.bat` 查看最近一次运行明细。详细参数、清理清单、局限与验证方法见 `使用说明.md`。

## 便携版

在任意 Windows 电脑上使用：本机装好 Python 3 后，直接拷 `.py` + `.bat` + `使用说明.md` 即可；目标机没装 Python 时，从 python.org 下载 Python 3.14.3 embeddable amd64 解压到 `python/` 目录，bat 会自动优先使用它。也可运行 `python 打包便携版.py` 生成本地发行包（该脚本默认不携带运行时，需自行准备 `python/` 或缓存 zip）。

## 安全与隐私

- 清理**必须在最后一次保存之后**进行——WPS 每次保存都会把用户名写回去。
- 运行会在工具目录生成 `_备份/`（原件）和 `_日志/`（含原始作者名与本机路径），**两者都含敏感信息，绝不要随清理后的文件一起外发**。
- `selftest.py` 使用合成样本自检，不接触真实文件，样本数据已中性化。

## 许可与免责

本工具按「原样」提供，用于个人/内部的外发前隐私清理。请自行验证清理结果（WPS 文档属性、PDF 属性）后再外发；对因使用本工具造成的任何数据问题不承担责任。脚本仅使用 Python 标准库。
