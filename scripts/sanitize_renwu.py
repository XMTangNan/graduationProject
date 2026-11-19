#!/usr/bin/env python3
"""清理交付 docx 中继承自他人模板的残留：批注、他人身份元数据、WPS 账号信息与错映射字符。

处理项：
    1. 删除 word/comments.xml 及 document.xml 内全部批注引用（6 条 2021 年学院填写提示批注）；
    2. 删除 docProps/custom.xml（含 KSOProductBuildVer、他人 userId 的
       KSOTemplateDocerSaveRecord、ICV）及其 Content_Types 覆写与关系项；
    3. 重写 docProps/core.xml：去掉 creator=sign、lastModifiedBy=派大茗、lastPrinted=2008、
       created=2021、TotalTime 等他人痕迹与旧校名标题，改为本文件真实信息；
    4. 去掉 docProps/app.xml 的 TotalTime（他人累计编辑时长）；
    5. 修掉"学生自拟"前被错映射成 £ 的空方框：改为华文中宋 36 半磅的 □，与其他未勾选项一致；
    6. 删除 6 个只含 rPr/w:vanish 且无文本的空 run（模板残留，非隐藏水印）。
"""

import datetime
import pathlib
import shutil
import zipfile

from lxml import etree

TARGET = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)
TITLE = "面向大众民事的联盟链电子公证存证系统设计与实现"
AUTHOR = "周健"

CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC = "{http://purl.org/dc/elements/1.1/}"
DCT = "{http://purl.org/dc/terms/}"
XSI = "{http://www.w3.org/2001/XMLSchema-instance}"
EPUB = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"


def rebuild(parts):
    """就地重写 docx：套用全部清理动作后回写压缩包。

    Args:
        parts: dict，成员名 -> bytes，来自原压缩包。
    Returns:
        dict: 清理后的成员名 -> bytes。
    """
    out = {}
    drop = {"word/comments.xml", "docProps/custom.xml"}

    # 1) document.xml：删批注引用、错映射字符、空 vanish run
    doc = etree.fromstring(parts["word/document.xml"])
    removed = 0
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for el in doc.findall(".//" + W + tag):
            el.getparent().remove(el)
            removed += 1
    fixed_box = 0
    gone = 0
    for run in doc.findall(".//" + W + "r"):
        texts = run.findall(W + "t")
        joined = "".join((t.text or "") for t in texts)
        rpr = run.find(W + "rPr")
        if joined == "\u00a3" and rpr is not None and "Wingdings 2" in etree.tostring(rpr, encoding="unicode"):
            fonts = rpr.find(W + "rFonts")
            for attr in ("ascii", "hAnsi", "cs"):
                if fonts is not None and attr in fonts.attrib:
                    del fonts.attrib[{"ascii": W + "ascii", "hAnsi": W + "hAnsi", "cs": W + "cs"}[attr]]
            if fonts is not None:
                fonts.set(W + "eastAsia", "华文中宋")
            ics = rpr.find(W + "iCs")
            if ics is not None:
                rpr.remove(ics)
            for size_tag, val in ((W + "sz", "36"), (W + "szCs", "36")):
                node = rpr.find(size_tag)
                if node is None:
                    node = etree.SubElement(rpr, size_tag)
                node.set(W + "val", val)
            texts[0].text = "\u25a1"
            fixed_box += 1
            continue
        if rpr is not None and rpr.find(W + "vanish") is not None and joined == "":
            run.getparent().remove(run)
            gone += 1
    print("  document.xml: 批注引用清除=%d, 方框字符修正=%d, 空vanish run删除=%d" % (removed, fixed_box, gone))
    out["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 2) 关系与内容类型：去掉 comments / custom 的引用
    ct = etree.fromstring(parts["[Content_Types].xml"])
    for ov in ct.findall(CT + "Override"):
        if ov.get("PartName") in ("/word/comments.xml", "/docProps/custom.xml"):
            ct.remove(ov)
    out["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    rels = etree.fromstring(parts["word/_rels/document.xml.rels"])
    for r in list(rels):
        if "comments" in r.get("Target", "").lower():
            rels.remove(r)
    out["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    root_rels = etree.fromstring(parts["_rels/.rels"])
    for r in list(root_rels):
        if "custom" in r.get("Target", "").lower():
            root_rels.remove(r)
    out["_rels/.rels"] = etree.tostring(root_rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 3) core.xml：真实且自洽的文档属性
    now = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)
    created = datetime.datetime(2026, 11, 17, 10, 5, 0, tzinfo=datetime.timezone.utc)
    def node(tag, text=None, **attrs):
        el = etree.Element(tag)
        if text is not None:
            el.text = text
        for k, v in attrs.items():
            el.set(k, v)
        return el
    core = etree.Element(CP + "coreProperties", nsmap={
        "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "dc": "http://purl.org/dc/elements/1.1/",
        "dcterms": "http://purl.org/dc/terms/",
        "dcmitype": "http://purl.org/dc/dcmitype/",
        "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    })
    core.append(node(DC + "creator", AUTHOR))
    core.append(node(CP + "keywords", ""))
    core.append(node(DC + "description", "本科毕业论文（设计）任务书与开题报告"))
    core.append(node(CP + "lastModifiedBy", AUTHOR))
    core.append(node(DC + "title", TITLE + "（任务书与开题报告）"))
    core.append(node(DC + "subject", "本科毕业论文（设计）"))
    core.append(node(CP + "revision", "1"))
    core.append(node(DCT + "created", created.isoformat().replace("+00:00", "Z"), **{XSI + "type": "dcterms:W3CDTF"}))
    core.append(node(DCT + "modified", now.isoformat().replace("+00:00", "Z"), **{XSI + "type": "dcterms:W3CDTF"}))
    core.append(node(DC + "language", "zh-CN"))
    out["docProps/core.xml"] = etree.tostring(core, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 4) app.xml：去掉他人累计编辑时长
    app = etree.fromstring(parts["docProps/app.xml"])
    for tt in app.findall(EPUB + "TotalTime"):
        app.remove(tt)
    out["docProps/app.xml"] = etree.tostring(app, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 5) 其余成员原样保留（不含被丢弃的 parts）
    for name, data in parts.items():
        if name in out or name in drop or name.startswith("word/_rels/comments"):
            continue
        out[name] = data
    return out


def main():
    """读入 docx，执行清理并回写。"""
    with zipfile.ZipFile(TARGET) as z:
        parts = {n: z.read(n) for n in z.namelist()}
        order = list(z.namelist())
    cleaned = rebuild(parts)
    tmp = TARGET.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            if name in cleaned:
                z.writestr(name, cleaned[name])
    shutil.move(str(tmp), str(TARGET))
    left = [n for n in zipfile.ZipFile(TARGET).namelist() if "comment" in n or "custom" in n]
    print("  清理后仍含 comment/custom 的成员:", left)
    print("  bytes=", len(TARGET.read_bytes()))


if __name__ == "__main__":
    main()
