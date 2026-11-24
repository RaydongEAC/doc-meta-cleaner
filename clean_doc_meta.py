#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_doc_meta.py  v2.2
外发前清除 Word / Excel / PowerPoint / PDF 里的作者信息与版本信息。

不依赖 Word / WPS / Office / 任何第三方包，只用 Python 标准库。
卸载了 Office、只装 WPS 的机器照样能跑。

用法:
    python clean_doc_meta.py 文件1.docx 待发目录\\ "报告.pdf"
    python clean_doc_meta.py 待发目录\\ --dry-run      只体检，不改动
    python clean_doc_meta.py 待发目录\\ --deep         连标题、时间戳一起清
    python clean_doc_meta.py 待发目录\\ --scan         额外扫描可疑残留
    python clean_doc_meta.py 文件.pdf --no-backup      不生成备份
    python clean_doc_meta.py --show-log                 打开最近一次运行日志

设计要点:
  * OOXML: 直接移除 docProps 里的作者/版本节点，清空自定义属性，
    并清除批注、修订、人员表里残留的作者名；另清除 Excel 自动记录的
    本机绝对路径 x15ac:absPath（里面常有客户名、业务目录）；重打包时保留条目顺序与时间戳。
  * PDF  : 采用「等长覆写」——把 /Info 字符串与 XMP 里的值替换成等长空格，
    文件总字节数、所有对象偏移、xref 表完全不变，不会破坏文档结构。
    FlateDecode 压缩的 XMP 会解压 -> 清理 -> 重压缩 -> 等长回填并改写 /Length。
  * 默认在同目录 _备份/ 下留一份原件；处理前先写临时文件再原子替换。
  * 每次运行把完整明细写入 _日志/清理日志_YYYY-MM-DD.log（含原始值，勿外发）。
  * 幂等：反复运行不会重复改动。
"""

import argparse
import ctypes
import datetime
import os
import re
import shutil
import sys
import zipfile
import zlib

VERSION = "2.2"
HERE = os.path.dirname(os.path.abspath(__file__))
BACKUP_ROOT = os.path.join(HERE, "_备份")
LOG_ROOT = os.path.join(HERE, "_日志")

# ---------------------------------------------------------------- 清理清单
# core.xml：作者 / 版本 / 单位 / 备注等
CORE_TAGS = [
    "dc:creator",            # 作者
    "cp:lastModifiedBy",     # 最后保存者
    "cp:keywords",           # 关键词（常含项目名、客户名）
    "dc:description",        # 备注
    "dc:subject",            # 主题
    "cp:category",           # 类别
    "cp:contentStatus",      # 内容状态
    "cp:lastPrinted",        # 最后打印时间
    "cp:revision",           # 修订次数（版本痕迹）
]
# --deep 追加
CORE_DEEP = ["dc:title", "dcterms:created", "dcterms:modified"]

# app.xml：公司 / 管理者 / 生成程序及版本（WPS 会写 WPS Office_12.1.0.xxxxx_设备GUID）
APP_TAGS = ["Company", "Manager", "HyperlinkBase", "Application", "AppVersion"]

CUSTOM_EMPTY = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
)

OOXML_EXT = {".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xltx",
             ".xltm", ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"}
PDF_EXT = {".pdf"}
LEGACY_EXT = {".doc", ".xls", ".ppt", ".wps", ".et", ".dps", ".wpt", ".dot", ".xlt", ".pot"}

# 正文里的作者痕迹：按部件精确处理，避免误伤形状名之类普通 name 属性
AUTHOR_PARTS = {
    # 文件路径模式 -> 要清空的 XML 属性名
    r"^word/(document|comments.*|people|footnotes|endnotes|header\d*|footer\d*|settings)\.xml$":
        ["author", "initials"],
    r"^ppt/commentAuthors\.xml$": ["name", "initials"],
    r"^ppt/comments/.*\.xml$": ["initials", "authorInitials"],
    r"^xl/persons/.*\.xml$": ["displayName", "userId", "id"],
    r"^xl/threadedComments/.*\.xml$": ["displayName"],
}
# Excel 旧式批注里作者是文本节点 <author>张三</author>
TEXT_PARTS = {r"^xl/(comments.*|threadedComments/.*)\.xml$": ["author"]}

# PDF /Info 字典
PDF_KEYS_BASE = [b"Author", b"Creator", b"Producer", b"Subject", b"Keywords",
                 b"CreationDate", b"ModDate", b"LastModified"]
PDF_KEYS_DEEP = [b"Title", b"Trapped"]

# XMP 里的作者/工具/时间字段
XMP_TAGS = [
    b"dc:creator", b"dc:title", b"dc:description", b"dc:subject",
    b"pdf:Producer", b"pdf:Keywords", b"pdf:Copyright",
    b"xmp:CreatorTool", b"xmp:CreateDate", b"xmp:ModifyDate", b"xmp:MetadataDate",
    b"xmp:Label", b"xmp:Nickname", b"xmpmm:DocumentID", b"xmpmm:InstanceID",
    b"xmpmm:OriginalDocumentID", b"photoshop:Author", b"photoshop:AuthorsPosition",
    b"photoshop:CaptionWriter", b"tiff:Artist", b"tiff:HostComputer",
    b"Iptc4xmpCore:CreatorWorkEmail", b"Iptc4xmpCore:CreatorWorkURL",
]

# 可疑残留（只报告，不自动改）
SUSPECT_STR = [
    re.compile(r"[A-Za-z]:\\+Users\\+[^\\/\"'<>\s]+", re.I),
    re.compile(r"[A-Za-z]:\\+[^\"'<>\r\n]{4,200}"),   # 任意本机绝对路径（含客户目录）
    re.compile(r"file:[/\\]{2,3}[^\"'<>\s]+", re.I),
    re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b"),
    re.compile(r"\b1[3-9]\d{9}\b"),
    re.compile(r"KSOTemplateDocerSaveRecord|KSOProductBuildVer|WPS Office_\d", re.I),
]

# Excel 自动写入的本机绝对路径（x15ac:absPath），常含客户名、业务目录、用户名
ABSPATH_RE = re.compile(r'(<(?:\w+:)?absPath\b[^>]*?\burl\s*=\s*")([^"]*)(")', re.I)
# WPS 编辑器指纹：主题名 + 各类 WPS 专属命名空间声明（workbook / worksheet 都会写）
WPS_THEME_RE = re.compile(r'\bname="WPS"', re.I)
WPS_NS_DECL_RE = re.compile(r'\s+xmlns:([\w\-]+)="([^"]*wps[^"]*)"', re.I)
SUSPECT_BYTES = [
    re.compile(rb"[A-Za-z]:\\+Users\\+[^\\/\"'<>\s]+", re.I),
    re.compile(rb"\b[\w.+-]+@[\w-]+\.[\w.-]+\b"),
    re.compile(rb"\b1[3-9]\d{9}\b"),
]


def clip(s, n=52):
    s = str(s).replace("\r", " ").replace("\n", " ").strip()
    return s if len(s) <= n else s[:n] + "..."


def match_any(name, patterns):
    return any(re.match(p, name, re.I) for p in patterns)


# ---------------------------------------------------------------- XML 工具
def strip_xml_tags(xml, tags):
    """移除带值的成对标签或自闭合标签，返回 (新xml, [(标签, 原值)])"""
    found = []
    for tag in tags:
        t = re.escape(tag)
        pair = re.compile(r"\s*<%s(?:\s[^>]*)?>(.*?)</%s\s*>" % (t, t), re.S)

        def repl(m):
            val = re.sub(r"<[^>]*>", "", m.group(1)).strip()
            if val:
                found.append((tag, clip(val)))
            return ""

        xml = pair.sub(repl, xml)
        self_closing = re.compile(r"\s*<%s(?:\s[^>]*)?/>" % t)
        xml = self_closing.sub("", xml)
    return xml, found


def blank_xml_attrs(xml, names):
    """把属性值清空（保留属性名与结构），返回 (新xml, [(属性, 原值)])"""
    found = []
    pat = re.compile(
        r'\b((?:[A-Za-z_][\w.\-]*:)?(?:%s)\s*=\s*")([^"]*)("\s*)'
        % "|".join(re.escape(n) for n in names), re.I)

    def repl(m):
        if m.group(2).strip():
            found.append((m.group(1).split("=")[0].strip(), clip(m.group(2))))
        return m.group(1) + m.group(3)

    return pat.sub(repl, xml), found


def blank_xml_text(xml, tags):
    """把 <tag>值</tag> 里的文本换成等长空格（保留标签）"""
    found = []
    for tag in tags:
        t = re.escape(tag)
        pat = re.compile(r"(<%s(?:\s[^>]*)?>)([^<>]*)(</%s\s*>)" % (t, t), re.S)

        def repl(m):
            if m.group(2).strip():
                found.append((tag, clip(m.group(2))))
                return m.group(1) + " " * len(m.group(2)) + m.group(3)
            return m.group(0)

        xml = pat.sub(repl, xml)
    return xml, found


def remove_unused_wps_ns(xml):
    """删除未被使用的 WPS 专属命名空间声明；返回 (新xml, 删掉的前缀列表)"""
    removed = []
    for m in list(WPS_NS_DECL_RE.finditer(xml)):
        pre = m.group(1)
        # 前缀若在元素名/属性名里被真正用到，就不能删，否则 XML 会失效
        if re.search(r"<" + re.escape(pre) + r":|\s" + re.escape(pre) + r":", xml):
            continue

        def drop_from_ignorable(mm):
            keep = [v for v in mm.group(2).split() if v != pre]
            return mm.group(1) + " ".join(keep) + mm.group(3)

        xml = re.sub(r'(mc:Ignorable=")([^"]*)(")', drop_from_ignorable, xml)
        xml = xml.replace(m.group(0), "", 1)
        removed.append(pre)
    return xml, removed


def scan_text(text, limit=6):
    hits = []
    for pat in SUSPECT_STR:
        for m in pat.finditer(text):
            v = clip(m.group(0))
            if v not in hits:
                hits.append(v)
            if len(hits) >= limit:
                return hits
    return hits


# ---------------------------------------------------------------- OOXML
def clean_ooxml_build(zin, args):
    """返回 (改动字典, 命中列表, 可疑残留列表)"""
    names = zin.namelist()
    new_data, hits, suspects = {}, [], []

    def get(name):
        try:
            return zin.read(name).decode("utf-8", "ignore")
        except KeyError:
            return None

    core = get("docProps/core.xml")
    if core:
        out, found = strip_xml_tags(core, CORE_TAGS + (CORE_DEEP if args.deep else []))
        if found:
            new_data["docProps/core.xml"] = out
            for k, v in found:
                hits.append("core/%s=%s" % (k, v))

    app = get("docProps/app.xml")
    if app:
        out, found = strip_xml_tags(app, APP_TAGS)
        if found:
            new_data["docProps/app.xml"] = out
            for k, v in found:
                hits.append("app/%s=%s" % (k, v))

    custom = get("docProps/custom.xml")
    if custom and re.search(r"<property\b", custom, re.I):
        props = re.findall(r'name="([^"]*)"', custom)
        new_data["docProps/custom.xml"] = CUSTOM_EMPTY
        hits.append("custom.xml 清空 %d 项%s" % (
            len(props), "（" + ", ".join(clip(p, 24) for p in props[:3]) + "）" if props else ""))

    for name in names:
        if not name.lower().endswith(".xml") and not name.lower().endswith(".rels"):
            continue
        xml = get(name)
        if not xml:
            continue
        for pat, attrs in AUTHOR_PARTS.items():
            if match_any(name, [pat]):
                out, found = blank_xml_attrs(xml, attrs)
                if found:
                    xml = out
                    for k, v in found[:4]:
                        hits.append("%s/%s=%s" % (name.split("/")[-1], k, v))
                break
        for pat, tags in TEXT_PARTS.items():
            if match_any(name, [pat]):
                out, found = blank_xml_text(xml, tags)
                if found:
                    xml = out
                    for k, v in found[:4]:
                        hits.append("%s/%s=%s" % (name.split("/")[-1], k, v))
                break

        # Excel 自动记录的本机绝对路径，里面常有客户名、业务目录、用户名
        def wipe_abspath(m):
            if m.group(2).strip():
                hits.append("%s/本机绝对路径=%s" % (name.split("/")[-1], clip(m.group(2), 100)))
                return m.group(1) + m.group(3)
            return m.group(0)

        xml = ABSPATH_RE.sub(wipe_abspath, xml)

        # 可选：把 WPS 编辑器指纹中性化（默认不改，避免动到主题样式）
        if args.neutralize and name.endswith(".xml"):
            if "/theme/" in name:
                n = len(WPS_THEME_RE.findall(xml))
                if n:
                    xml = WPS_THEME_RE.sub('name="Office"', xml)
                    hits.append("%s/主题名 WPS→Office ×%d" % (name.split("/")[-1], n))
            xml, removed = remove_unused_wps_ns(xml)
            for pre in removed:
                hits.append("%s 移除未使用的 WPS 命名空间声明 xmlns:%s"
                            % (name.split("/")[-1], pre))

        if name in new_data:
            xml = new_data[name]
        else:
            original = get(name)
            if xml != original:
                new_data[name] = xml
        if args.scan:
            for s in scan_text(xml, 4):
                if s not in suspects:
                    suspects.append(s)

    # 模板/外部关系里常带本机用户名路径，只报告不自动改
    if args.scan:
        for name in names:
            if name.endswith(".rels"):
                rel = get(name) or ""
                for s in scan_text(rel, 3):
                    if s not in suspects:
                        suspects.append(s)

    if "_xmlsignatures" in " ".join(names):
        suspects.append("含数字签名（_xmlsignatures），清理后签名会失效")
    if args.scan:
        # 基于清理后的内容判断残留，否则已经 --neutralize 清掉的指纹还会误报
        for name in names:
            if name.endswith(".xml"):
                s = new_data.get(name, get(name) or "")
                if WPS_THEME_RE.search(s) or WPS_NS_DECL_RE.search(s):
                    t = "仍残留 WPS 编辑器指纹（主题名/命名空间）；如需中性化加 --neutralize"
                    if t not in suspects:
                        suspects.append(t)
                        break
    return new_data, hits, suspects[:8]


def process_ooxml(path, args, backup_dir):
    try:
        zin = zipfile.ZipFile(path, "r")
    except Exception as e:
        return "fail", "打不开（可能加密或已损坏）：%s" % e, [], []
    with zin:
        names = zin.namelist()
        if "[Content_Types].xml" not in names:
            return "skip", "不是标准 OOXML 包（可能是加密文档）", [], []
        new_data, hits, suspects = clean_ooxml_build(zin, args)
        if not new_data:
            return "clean", "本来就没有可清的作者/版本信息", suspects, []
        if args.dry_run:
            return "would", "发现 %d 项元数据" % len(hits), suspects, hits
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]
        zip_comment = zin.comment

    tmp = path + ".cleantmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            zout.comment = zip_comment
            for info, data in entries:
                if info.filename in new_data:
                    data = new_data[info.filename].encode("utf-8")
                zout.writestr(info, data)
        if backup_dir:
            do_backup(path, backup_dir)
        os.replace(tmp, path)
    except Exception as e:
        return "fail", "写入失败（原件未动）：%s" % e, suspects, hits
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return "done", "已清除 %d 项元数据" % len(hits), suspects, hits


# ---------------------------------------------------------------- PDF 工具
def wipe_pdf_value(buf, p, end):
    """把 p 处的 PDF 字符串值替换为等长空格。返回 (原值bytes, 新位置) 或 None"""
    c = buf[p]
    if c == 0x28:  # (
        i, depth = p + 1, 1
        while i < end:
            b = buf[i]
            if b == 0x5C:            # 反斜杠转义，跳过下一字节
                i += 2
                continue
            if b == 0x28:
                depth += 1
            elif b == 0x29:
                depth -= 1
                if depth == 0:
                    break
            i += 1
        if i >= end or depth != 0:
            return None
        old = bytes(buf[p + 1:i])
        for j in range(p + 1, i):
            buf[j] = 0x20
        return old, i + 1
    if c == 0x3C:  # < hex 字符串（排除 << 字典）
        if p + 1 < end and buf[p + 1] == 0x3C:
            return None
        i = p + 1
        while i < end and buf[i] != 0x3E:
            i += 1
        if i >= end:
            return None
        old = bytes(buf[p + 1:i])
        for j in range(p + 1, i):
            buf[j] = 0x20
        return old, i + 1
    return None


def decode_pdf_str(b):
    # 已清成空格的值直接视为空：否则 "  " 会被 utf-16-be 解成匕首符号，误判成还有内容
    if not b or not b.strip():
        return ""
    if b.startswith(b"\xfe\xff"):
        return b[2:].decode("utf-16-be", "replace")
    for enc in ("utf-8", "latin-1"):
        try:
            s = b.decode(enc)
            if s.strip():
                return s
        except UnicodeError:
            pass
    return ""


def blank_info_dict(buf, start, end, keys):
    hits = []
    block = bytes(buf[start:end])
    for key in keys:
        pat = re.compile(rb"/" + key + rb"(?=[\s(<\[/])")
        for m in pat.finditer(block):
            p = start + m.end()
            while p < end and buf[p] in b" \t\r\n\x00\x0c":
                p += 1
            if p >= end or buf[p] not in (0x28, 0x3C):
                continue
            r = wipe_pdf_value(buf, p, end)
            if r:
                old, _ = r
                val = decode_pdf_str(old)
                if val.strip():
                    hits.append("/%s=%s" % (key.decode(), clip(val, 36)))
    return hits


def blank_xmp(data):
    """XMP 里把敏感字段值替换为等长空格；字节数不变，XML 结构不变"""
    out = bytearray(data)
    hits = []
    for tag in XMP_TAGS:
        t = re.escape(tag)
        # 容器形式：<dc:creator>...<rdf:li>张三</rdf:li>...</dc:creator>
        container = re.compile(rb"(<" + t + rb"\b[^>]*>)(.*?)(</" + t + rb"\s*>)", re.S)
        for m in container.finditer(data):
            whole = m.group(0)
            for txt in re.finditer(rb">([^<>]+)<", whole):
                s, e = m.start() + txt.start(1), m.start() + txt.end(1)
                chunk = bytes(data[s:e])
                if chunk.strip():
                    hits.append(tag.decode() + "=" + clip(chunk.decode("utf-8", "ignore"), 30))
                    out[s:e] = b" " * (e - s)
        # 属性形式：<rdf:Description xmp:CreatorTool="WPS Office" ...>
        attr = re.compile(rb"\b" + t + rb'\s*=\s*"([^"]*)"')
        for m in attr.finditer(data):
            s, e = m.start(1), m.end(1)
            if bytes(data[s:e]).strip():
                hits.append(tag.decode() + "=" + clip(bytes(data[s:e]).decode("utf-8", "ignore"), 30))
                out[s:e] = b" " * (e - s)
    return bytes(out), hits


def rewrite_length(buf, head_start, stream_kw, old_len, new_len):
    """把 /Length 改成 new_len，用前导零补齐保持字节数不变。返回是否成功"""
    head = bytes(buf[head_start:stream_kw])
    if old_len == new_len:
        return True
    m = re.search(rb"/Length\s+(\d+)\s+(\d+)\s+R", head)
    if m:  # 间接对象
        objnum, gen = m.group(1).decode(), m.group(2).decode()
        om = re.search(rb"(?<![0-9])" + re.escape(objnum.encode()) + rb"\s+" +
                       re.escape(gen.encode()) + rb"\s+obj\s*(\d+)\s*endobj", bytes(buf))
        if not om:
            return False
        digits = om.group(1)
        txt = str(new_len).encode().zfill(len(digits))
        if len(txt) != len(digits):
            return False
        buf[om.start(1):om.end(1)] = txt
        return True
    m = re.search(rb"/Length\s+(\d+)", head)
    if not m:
        return False
    digits = m.group(1)
    txt = str(new_len).encode().zfill(len(digits))
    if len(txt) != len(digits):
        return False
    pos = head_start + m.start(1)
    buf[pos:pos + len(digits)] = txt
    return True


def process_pdf(path, args, backup_dir):
    with open(path, "rb") as f:
        data = f.read()
    if not data.startswith(b"%PDF-"):
        # 有些文件前面有垃圾字节，仍尝试
        if data.find(b"%PDF-") < 0:
            return "skip", "不是有效的 PDF 文件", [], []
    if re.search(rb"/Encrypt\b", data):
        return "skip", "加密 PDF：字符串已加密，等长覆写会破坏文件，已跳过", [], []
    signed = bool(re.search(rb"/ByteRange\b|/Type\s*/Sig\b|/AcroForm\b", data))

    buf = bytearray(data)
    hits = []
    keys = PDF_KEYS_BASE + (PDF_KEYS_DEEP if args.deep else [])

    # 1) 所有 /Info 字典（含增量更新遗留下来的旧版本）
    for m in re.finditer(rb"(\d+)\s+(\d+)\s+obj\b", bytes(buf)):
        s = m.end()
        e = buf.find(b"endobj", s)
        if e < 0 or e - s > 4000:
            continue
        head = bytes(buf[s:e])
        if b"stream" in head or b"<<" not in head:
            continue
        if not re.search(rb"/(Author|Creator|Producer|CreationDate|ModDate|Keywords)\b", head):
            continue
        hits += blank_info_dict(buf, s, e, keys)

    # 2) XMP 元数据流（明文 或 FlateDecode）
    for m in re.finditer(rb"(\d+)\s+(\d+)\s+obj\b", bytes(buf)):
        s = m.end()
        sidx = buf.find(b"stream", s, s + 2000)
        if sidx < 0:
            continue
        head = bytes(buf[s:sidx])
        if not (re.search(rb"/Subtype\s*/XML\b", head) or
                re.search(rb"/Type\s*/Metadata\b", head)):
            continue
        lm = re.search(rb"/Length\s+(\d+)", head)
        if not lm:
            continue
        old_len = int(lm.group(1))
        p = sidx + 6
        if buf[p:p + 2] == b"\r\n":
            p += 2
        elif buf[p:p + 1] in (b"\n", b"\r"):
            p += 1
        d_end = p + old_len
        if d_end > len(buf):
            continue
        tail = bytes(buf[d_end:d_end + 24]).lstrip(b"\r\n \t\x00\x0c")
        if not tail.startswith(b"endstream"):
            continue
        comp = bytes(buf[p:d_end])
        is_flate = bool(re.search(rb"/FlateDecode\b", head))
        has_other_filter = bool(re.search(rb"/Filter\b", head)) and not is_flate

        if not re.search(rb"/Filter\b", head):
            new_raw, xhits = blank_xmp(comp)
            if xhits:
                buf[p:d_end] = new_raw
                hits += ["XMP/" + h for h in xhits]
        elif is_flate and not has_other_filter:
            try:
                raw = zlib.decompress(comp)
            except Exception:
                continue
            new_raw, xhits = blank_xmp(raw)
            if not xhits:
                continue
            new_comp = zlib.compress(new_raw, 9)
            if len(new_comp) > len(comp):
                hits.append("XMP 压缩后变长，未能清理（建议用 WPS 重新导出 PDF）")
                continue
            if not rewrite_length(buf, s, sidx, old_len, len(new_comp)):
                continue
            buf[p:d_end] = new_comp + b" " * (len(comp) - len(new_comp))
            hits += ["XMP/" + h for h in xhits]
        else:
            hits.append("XMP 使用非 Flate 过滤器，未能清理")

    suspects = []
    if args.scan:
        for pat in SUSPECT_BYTES:
            for sm in pat.finditer(bytes(buf)):
                v = clip(sm.group(0).decode("utf-8", "ignore"))
                if v not in suspects and len(suspects) < 6:
                    suspects.append(v)
    if signed:
        suspects.append("检测到数字签名/表单，清理会使签名失效，请确认后再外发")

    if not hits:
        return "clean", "本来就没有可清的作者/版本信息", suspects, []
    if args.dry_run:
        return "would", "发现 %d 项元数据" % len(hits), suspects, hits
    if len(buf) != len(data):
        return "fail", "内部校验失败：文件长度发生变化，原件未动", suspects, hits

    tmp = path + ".cleantmp"
    try:
        with open(tmp, "wb") as f:
            f.write(buf)
        if backup_dir:
            do_backup(path, backup_dir)
        os.replace(tmp, path)
    except Exception as e:
        return "fail", "写入失败（原件未动）：%s" % e, suspects, hits
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return "done", "已清除 %d 项元数据" % len(hits), suspects, hits


# ---------------------------------------------------------------- 通用流程
def do_backup(path, backup_dir):
    os.makedirs(backup_dir, exist_ok=True)
    base = os.path.basename(path)
    dst = os.path.join(backup_dir, base)
    n = 1
    while os.path.exists(dst):
        dst = os.path.join(backup_dir, "%03d_%s" % (n, base))
        n += 1
    shutil.copy2(path, dst)
    return dst


def collect(inputs, recursive=True):
    files = []
    for a in inputs:
        if os.path.isdir(a):
            if recursive:
                for root, dirs, fs in os.walk(a):
                    dirs[:] = [d for d in dirs
                               if d not in ("_备份", "_日志") and not d.startswith(".")]
                    files += [os.path.join(root, f) for f in fs]
            else:
                files += [os.path.join(a, f) for f in os.listdir(a)]
        else:
            files.append(a)
    out = []
    for f in files:
        if not os.path.isfile(f):
            continue
        name = os.path.basename(f)
        if name.startswith("~$") or f.endswith((".cleantmp", ".bak")):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext in OOXML_EXT or ext in PDF_EXT or ext in LEGACY_EXT:
            out.append(f)
    return out


def popup(title, text):
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)
    except Exception:
        pass


def latest_log():
    if not os.path.isdir(LOG_ROOT):
        return None
    fs = [os.path.join(LOG_ROOT, f) for f in os.listdir(LOG_ROOT) if f.endswith(".log")]
    return max(fs, key=os.path.getmtime) if fs else None


def show_log():
    p = latest_log()
    if not p:
        print("还没有日志。运行一次清理（或体检）后会自动生成。")
        return 0
    print("最新日志：" + p)
    try:
        os.startfile(p)
    except Exception as e:
        print("自动打开失败，请手动打开上面的路径：%s" % e)
    return 0


def write_log(records, args, summary_lines, backup_dir):
    """把每次运行的完整明细落盘，方便事后核对（日志含原始值，勿外发）"""
    os.makedirs(LOG_ROOT, exist_ok=True)
    path = os.path.join(LOG_ROOT, "清理日志_%s.log" % datetime.date.today().isoformat())
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(path, "a", encoding="utf-8") as f:
        f.write("\n" + "=" * 76 + "\n")
        f.write("时间: %s   版本: v%s   模式: %s\n" % (
            now, VERSION, "试运行（未改动文件）" if args.dry_run else "实际清理"))
        f.write("参数: deep=%s scan=%s neutralize=%s backup=%s\n" % (
            args.deep, args.scan, args.neutralize, not args.no_backup))
        f.write("输入: %s\n" % " | ".join(args.paths))
        f.write("-" * 76 + "\n")
        for r in records:
            f.write("[%s] %s\n" % (r["label"], r["path"]))
            f.write("    结果: %s\n" % r["msg"])
            for h in r["hits"]:
                f.write("      项: %s\n" % h)
            for s in r["suspects"]:
                f.write("    注意: %s\n" % s)
        f.write("-" * 76 + "\n")
        for line in summary_lines:
            f.write(line + "\n")
        if backup_dir and os.path.isdir(backup_dir):
            f.write("原件备份: %s\n" % backup_dir)
    return path


def main(argv=None):
    # pythonw.exe 启动时 sys.stdout 为 None，补一个空输出避免 print 报错
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    ap = argparse.ArgumentParser(
        description="清除 Word / Excel / PPT / PDF 的作者与版本信息（v%s）" % VERSION,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("paths", nargs="*", help="文件或目录，可多个；拖放即可")
    ap.add_argument("--dry-run", action="store_true", help="只体检，不改动文件")
    ap.add_argument("--deep", action="store_true", help="连标题、创建/修改时间一起清")
    ap.add_argument("--scan", action="store_true", help="额外扫描可疑残留（本机路径、邮箱等）")
    ap.add_argument("--neutralize", action="store_true", help="把 WPS 主题名/命名空间中性化（编辑器指纹）")
    ap.add_argument("--no-backup", action="store_true", help="不生成备份")
    ap.add_argument("--no-log", action="store_true", help="不写运行日志")
    ap.add_argument("--show-log", action="store_true", help="打开最近一次运行日志后退出")
    ap.add_argument("--no-recursive", action="store_true", help="目录不递归")
    ap.add_argument("--gui", action="store_true", help="结束时弹窗显示结果（右键菜单用）")
    args = ap.parse_args(argv)

    if args.show_log:
        return show_log()
    if not args.paths:
        ap.print_help()
        print("\n把文件拖到 DROP_ON_ME.bat 上，或把路径作为参数传进来。")
        return 1

    files = collect(args.paths, recursive=not args.no_recursive)
    if not files:
        msg = "没有找到可处理的 Word / Excel / PPT / PDF 文件。"
        print(msg)
        if args.gui:
            popup("文档属性清理 v%s" % VERSION, msg)
        return 1

    backup_dir = None
    if not args.dry_run and not args.no_backup:
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_dir = os.path.join(BACKUP_ROOT, stamp)

    labels = {"done": "已清除", "would": "待清除", "clean": "无需处理",
              "skip": "跳过", "fail": "失败"}
    stat, records, total_items = {}, [], 0
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        try:
            if ext in LEGACY_EXT:
                tag, msg = "skip", "旧版二进制格式（%s），请先在 WPS 里另存为 .docx/.xlsx/.pptx" % ext
                suspects, hits = [], []
            elif ext in OOXML_EXT:
                tag, msg, suspects, hits = process_ooxml(f, args, backup_dir)
            else:
                tag, msg, suspects, hits = process_pdf(f, args, backup_dir)
        except Exception as e:
            tag, msg, suspects, hits = "fail", "出错（原件未动）：%s" % e, [], []
        stat[tag] = stat.get(tag, 0) + 1
        total_items += len(hits)
        records.append({"label": labels[tag], "path": os.path.abspath(f),
                        "msg": msg, "hits": hits, "suspects": suspects})
        print("[%s] %s" % (labels[tag], os.path.basename(f)), flush=True)
        print("      %s" % msg, flush=True)
        for h in hits:
            print("        - %s" % h, flush=True)
        for s in suspects:
            print("      注意：%s" % s, flush=True)

    summary_lines = [
        "共 %d 个文件：%s" % (
            len(files), " / ".join("%s %d 个" % (labels[k], v) for k, v in stat.items())),
        "共%s %d 项元数据%s" % (
            "检出" if args.dry_run else "清除", total_items,
            "（试运行，未写入）" if args.dry_run else ""),
    ]
    print()
    for line in summary_lines:
        print(line, flush=True)
    if backup_dir and os.path.isdir(backup_dir):
        print("原件备份：%s" % backup_dir)
    if not args.dry_run:
        print("提示：务必在最后一次保存之后再清理，否则 WPS 保存时会把用户名写回去。")
    if not args.no_log:
        print("运行日志：%s   （含原始值，勿外发）" %
              write_log(records, args, summary_lines, backup_dir))
    if args.gui:
        popup("文档属性清理 v%s" % VERSION, "\n".join(summary_lines)[:900])
    return 0


if __name__ == "__main__":
    sys.exit(main())
