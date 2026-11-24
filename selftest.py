#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selftest.py — 自检：生成带作者/版本信息的合成 docx / xlsx / pptx / pdf，
跑一遍 clean_doc_meta.py，验证清干净且文件结构没被破坏。

用法: python selftest.py
样本在系统临时目录里生成，跑完自动删除，不碰你的真实文件。
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "clean_doc_meta.py")

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        '</Relationships>')
CORE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>季度报告</dc:title><dc:subject>客户报价</dc:subject>'
        '<dc:creator>Administrator</dc:creator><cp:keywords>示例项目A</cp:keywords>'
        '<dc:description>内部</dc:description><cp:lastModifiedBy>测试用户</cp:lastModifiedBy>'
        '<cp:revision>37</cp:revision><cp:lastPrinted>2026-11-20T09:30:00Z</cp:lastPrinted>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-01T08:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-11-21T19:00:00Z</dcterms:modified>'
        '</cp:coreProperties>')
APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
       'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
       '<Application>WPS Office_12.1.0.29158_F1E327BC-9A88-4E11-B3C2D5E6F7A8B9</Application>'
       '<AppVersion>12.1.0.29158</AppVersion><Company>示例公司</Company>'
       '<Manager>测试经理</Manager><HyperlinkBase>C:\\Users\\Administrator\\</HyperlinkBase>'
       '<DocSecurity>0</DocSecurity></Properties>')
CUSTOM = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
          'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
          '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="KSOTemplateDocerSaveRecord">'
          '<vt:lpwstr>eyJoZGlkIjoiYWJjIiwidXNlcklkIjoiNDMzMzc4Njk3In0=</vt:lpwstr></property>'
          '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="KSOProductBuildVer">'
          '<vt:lpwstr>12.1.0.29158</vt:lpwstr></property>'
          '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="部门">'
          '<vt:lpwstr>销售部</vt:lpwstr></property></Properties>')
DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
       '<w:body><w:p><w:r><w:t>正文内容不应被改动</w:t></w:r>'
       '<w:ins w:author="Administrator" w:date="2026-11-21T10:00:00Z"><w:r><w:t>修订</w:t></w:r></w:ins>'
       '</w:p></w:body></w:document>')
COMMENTS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:comment w:author="Administrator" w:initials="Adm" w:id="0">'
            '<w:p><w:r><w:t>批注内容</w:t></w:r></w:p></w:comment></w:comments>')
PEOPLE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">'
          '<w15:person w15:author="测试用户" w15:id="abc123"/></w15:people>')
XL_WORKBOOK = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
               'xmlns:dbsheet="http://web.wps.cn/et/2021/dbsheet">'
               '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
               '<mc:Choice Requires="x15">'
               '<x15ac:absPath url="E:\\WorkFiles\\Sales Management\\Key Account\\'
               'ClientA\\ProjectX-2026\\" '
               'xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>'
               '</mc:Choice></mc:AlternateContent></workbook>')
XL_THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="WPS">'
            '<a:themeElements/></a:theme>')
XL_SHEET = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">'
            '<sheetData/></worksheet>')
XL_COMMENTS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<authors><author>Administrator</author></authors>'
               '<commentList><comment ref="A1" authorId="0"><text><t>批注</t></text></comment></commentList>'
               '</comments>')
XL_PERSONS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<person xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" '
              'displayName="测试用户" id="{AAA-BBB}" userId="433378697"/>')
PPT_AUTHORS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
               '<p:cmAuthor id="1" name="Administrator" initials="Adm" lastIdx="1" clrIdx="0"/>'
               '</p:cmAuthorLst>')

XMP = (
    '<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?>'
    '<x:xmpmeta xmlns:x="adobe:ns:meta/">'
    '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
    '<rdf:Description xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:xmp="http://ns.adobe.com/xap/1.0/" '
    'xmlns:pdf="http://ns.adobe.com/pdf/1.3/" '
    'xmlns:xmpmm="http://ns.adobe.com/xap/1.0/mm/" '
    'xmp:CreatorTool="WPS Office_12.1.0.29158" pdf:Producer="WPS Office">'
    '<dc:creator><rdf:Seq><rdf:li>Administrator</rdf:li></rdf:Seq></dc:creator>'
    '<dc:title><rdf:Alt><rdf:li xml:lang="x-default">季度报告</rdf:li></rdf:Alt></dc:title>'
    '<xmpmm:DocumentID>uuid:Administrator-1234</xmpmm:DocumentID>'
    '<xmp:CreateDate>2026-11-21T19:00:00+08:00</xmp:CreateDate>'
    '</rdf:Description></rdf:RDF></x:xmpmeta>'
    '<?xpacket end="w"?>'
).encode("utf-8")


def make_ooxml(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in parts:
            z.writestr(name, content.encode("utf-8") if isinstance(content, str) else content)


def make_pdf(path, flate):
    xmp_obj = zlib.compress(XMP, 9) if flate else XMP
    filter_part = b" /Filter /FlateDecode" if flate else b""
    page_content = b"BT /F1 12 Tf 20 100 Td (Hi) Tj ET\n"
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
        b"<< /Length " + str(len(page_content)).encode() + b" >>\nstream\n" +
        page_content + b"endstream",
        b"<< /Type /Metadata /Subtype /XML" + filter_part +
        b" /Length " + str(len(xmp_obj)).encode() + b" >>\nstream\n" + xmp_obj + b"\nendstream",
        b"<< /Author (Administrator) /Creator (WPS Writer) "
        b"/Producer (WPS Office_12.1.0.29158_F1E327BC) "
        b"/Title (\xe5\xad\xa3\xe5\xba\xa6\xe6\x8a\xa5\xe5\x91\x8a) "
        b"/Keywords (\xe7\xab\x8b\xe7\xa7\x91\xe4\xba\xa7\xe4\xb8\x9a\xe5\x9b\xad) "
        b"/CreationDate (D:20261121190000) /ModDate (D:20261121190000) >>",
    ]
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(len(out))
        out += ("%d 0 obj\n" % i).encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += ("xref\n0 %d\n" % (len(objs) + 1)).encode()
    out += b"0000000000 65535 f \n"
    for off in offsets:
        out += ("%010d 00000 n \n" % off).encode()
    out += ("trailer\n<< /Size %d /Root 1 0 R /Info 6 0 R >>\nstartxref\n%d\n%%%%EOF\n"
            % (len(objs) + 1, xref_pos)).encode()
    with open(path, "wb") as f:
        f.write(bytes(out))


def run(args):
    return subprocess.run([sys.executable, SCRIPT] + args + ["--no-log"],
                          capture_output=True, text=True, encoding="utf-8", errors="replace")


def check(cond, label):
    print(("  [OK]   " if cond else "  [FAIL] ") + label)
    return bool(cond)


def main():
    tmp = tempfile.mkdtemp(prefix="meta_selftest_")
    ok = True
    try:
        base = {
            "[Content_Types].xml": CT, "_rels/.rels": RELS,
            "docProps/core.xml": CORE, "docProps/app.xml": APP,
            "docProps/custom.xml": CUSTOM,
        }
        samples = {
            "sample.docx": dict(base, **{"word/document.xml": DOC,
                                         "word/comments.xml": COMMENTS,
                                         "word/people.xml": PEOPLE}),
            "sample.xlsx": dict(base, **{"xl/workbook.xml": XL_WORKBOOK,
                                         "xl/theme/theme1.xml": XL_THEME,
                                         "xl/worksheets/sheet1.xml": XL_SHEET,
                                         "xl/comments1.xml": XL_COMMENTS,
                                         "xl/persons/person1.xml": XL_PERSONS}),
            "sample.pptx": dict(base, **{"ppt/presentation.xml": "<p:presentation/>",
                                         "ppt/commentAuthors.xml": PPT_AUTHORS}),
        }
        for name, parts in samples.items():
            make_ooxml(os.path.join(tmp, name), list(parts.items()))
        make_pdf(os.path.join(tmp, "plain_xmp.pdf"), False)
        make_pdf(os.path.join(tmp, "flate_xmp.pdf"), True)

        before = {}
        for name in list(samples) + ["plain_xmp.pdf", "flate_xmp.pdf"]:
            p = os.path.join(tmp, name)
            before[name] = open(p, "rb").read()
            if name.endswith((".docx", ".xlsx", ".pptx")):
                with zipfile.ZipFile(p) as z:
                    before[name + "_names"] = z.namelist()

        print("\n== 1. 试运行 --dry-run --scan ==")
        r = run([tmp, "--dry-run", "--scan"])
        print(r.stdout.strip()[:1200])
        ok &= check(r.returncode == 0 and "待清除 5" in r.stdout, "5 个文件全部检出待清除项")

        print("\n== 1b. --neutralize / --deep 试运行 ==")
        r = run([os.path.join(tmp, "sample.xlsx"), "--dry-run", "--neutralize"])
        ok &= check("主题名 WPS" in r.stdout and "WPS 命名空间声明" in r.stdout,
                    "WPS 主题名与命名空间可被中性化")
        ok &= check("sheet1.xml" in r.stdout and "workbook.xml" in r.stdout,
                    "workbook 与 worksheet 两处 WPS 命名空间都识别到")
        r = run([os.path.join(tmp, "sample.docx"), "--dry-run", "--deep"])
        ok &= check("dcterms:created" in r.stdout, "--deep 能清创建/修改时间")

        print("\n== 2. 实际清理 ==")
        r = run([tmp, "--no-backup"])
        print(r.stdout.strip()[:1600])
        ok &= check(r.returncode == 0 and "已清除 5" in r.stdout, "5 个文件全部清理成功")

        body_marks = {"sample.docx": "正文内容不应被改动",
                      "sample.xlsx": "mc:AlternateContent",
                      "sample.pptx": "<p:presentation/>"}
        print("\n== 3. OOXML 校验 ==")
        for name in samples:
            p = os.path.join(tmp, name)
            with zipfile.ZipFile(p) as z:
                names = z.namelist()
                core = z.read("docProps/core.xml").decode("utf-8")
                app = z.read("docProps/app.xml").decode("utf-8")
                custom = z.read("docProps/custom.xml").decode("utf-8")
                allxml = "".join(z.read(n).decode("utf-8", "ignore")
                                 for n in names if n.endswith(".xml"))
            ok &= check(names == before[name + "_names"], name + " 条目数与顺序不变")
            ok &= check("Administrator" not in core and "dc:creator>" not in core,
                        name + " 作者/最后保存者已移除")
            ok &= check("WPS Office_" not in app and "<Company>" not in app,
                        name + " 版本号/公司/管理者已移除")
            ok &= check("KSOTemplateDocerSaveRecord" not in custom,
                        name + " WPS 账号记录已清空")
            ok &= check("w:author=\"Administrator\"" not in allxml and
                        "w15:author=\"测试用户\"" not in allxml and
                        "name=\"Administrator\"" not in allxml and
                        "<author>Administrator</author>" not in allxml and
                        'displayName="测试用户"' not in allxml,
                        name + " 批注/修订/人员表作者已清空")
            ok &= check(body_marks[name] in allxml, name + " 正文未受影响")
            if name == "sample.xlsx":
                ok &= check("E:\\WorkFiles" not in allxml and
                            "ClientA" not in allxml,
                            name + " 本机绝对路径与客户目录已清空")

        print("\n== 4. PDF 校验 ==")
        for name in ("plain_xmp.pdf", "flate_xmp.pdf"):
            p = os.path.join(tmp, name)
            data = open(p, "rb").read()
            old = before[name]
            ok &= check(len(data) == len(old), name + " 文件长度不变")
            sx_old = re.search(rb"startxref\s+(\d+)", old).group(1)
            sx_new = re.search(rb"startxref\s+(\d+)", data).group(1)
            ok &= check(sx_old == sx_new, name + " startxref 偏移不变")
            ok &= check(old[old.find(b"xref\n0 "):old.find(b"trailer")] ==
                        data[data.find(b"xref\n0 "):data.find(b"trailer")],
                        name + " xref 表字节完全一致")
            ok &= check(b"Administrator" not in data, name + " /Author 已清空")
            ok &= check(b"WPS Office" not in data and b"WPS Writer" not in data,
                        name + " 生成器与版本号已清空")
            ok &= check(b"D:20261121190000" not in data, name + " 创建/修改时间已清空")
            mi = data.find(b"/Type /Metadata")
            m = re.search(rb"/Length (\d+)", data[mi:mi + 300])
            length = int(m.group(1))
            s = data.find(b"stream\n", mi) + 7
            blob = data[s:s + length]
            xmp = zlib.decompress(blob) if name.startswith("flate") else blob
            ok &= check(b"Administrator" not in xmp and b"WPS Office" not in xmp and
                        b"uuid:Administrator" not in xmp,
                        name + " XMP 里的作者/工具/文档ID 已清空")
            ok &= check(xmp.count(b"<rdf:RDF") == 1 and xmp.endswith(b"?>"),
                        name + " XMP 结构完好")

        print("\n== 5. 幂等性（再跑一次） ==")
        r = run([tmp, "--no-backup"])
        print(r.stdout.strip()[:600])
        ok &= check("无需处理 5" in r.stdout, "二次运行不再改动任何文件")

        print("\n== 结论：" + ("全部通过" if ok else "存在失败项，请检查") + " ==")
        return 0 if ok else 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
