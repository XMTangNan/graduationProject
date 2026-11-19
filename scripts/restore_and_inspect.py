#!/usr/bin/env python3
"""恢复被清零的毕设源文件，复制模板为交付文件，并检查表单单元格的段落格式。

用途：
    1) 比对 D 盘完好原件与工作区文件的 SHA-256，修复工作区中被写成全 0 字节的文件；
    2) 复制《任务书与开题报告》模板为交付用新文件（不覆盖模板本身）；
    3) 打印目标单元格内段落的 pPr / run 字体信息，供后续按原格式写入正文。
"""

import hashlib
import re
import shutil
import pathlib

from docx import Document
from docx.oxml.ns import qn

D = pathlib.Path("/mnt/d/Study/undergraduate/毕设/doc")
W = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/doc")

PAIRS = [
    (
        D / "申报表/2023131072_周健_SJ0212-2016版-本科毕业设计（论文）选题审批表_面向大众民事的联盟链电子公证存证系统设计与实现.docx",
        W / "申报表/2023131072_周健_SJ0212-2016版-本科毕业设计（论文）选题审批表_面向大众民事的联盟链电子公证存证系统设计与实现.docx",
    ),
    (D / "任务书与开题报告/任务书与开题报告.docx", W / "任务书与开题报告/任务书与开题报告.docx"),
    (D / "人工智能学院2027届本科毕业论文（设计）工作计划表.pdf", W / "人工智能学院2027届本科毕业论文（设计）工作计划表.pdf"),
]

OUT_NAME = "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_面向大众民事的联盟链电子公证存证系统设计与实现.docx"


def sha(path):
    """返回文件内容的 SHA-256 十六进制串。"""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def restore():
    """用 D 盘原件修复工作区中内容不一致（被清零）的文件。"""
    for src, dst in PAIRS:
        if not src.exists():
            print("SRC-MISSING", src)
            continue
        h_src = sha(src)
        h_dst = sha(dst) if dst.exists() else "MISSING"
        if h_src != h_dst:
            shutil.copy2(src, dst)
            print("RESTORED", dst.name, "verified=", sha(dst) == h_src, "bytes=", len(dst.read_bytes()))
        else:
            print("ALREADY-OK", dst.name)


def clone_template():
    """复制模板为交付文件（模板本身保持不变）。"""
    tpl = W / "任务书与开题报告/任务书与开题报告.docx"
    out = W / "任务书与开题报告" / OUT_NAME
    shutil.copy2(tpl, out)
    print("DELIVERABLE", out.name, "bytes=", len(out.read_bytes()))
    return out


def show(cell, tag):
    """打印单元格内各段落样式、pPr 与 run 字体信息。"""
    print("\n----", tag, "paras=", len(cell.paragraphs))
    for i, p in enumerate(cell.paragraphs):
        pr = p._p.find(qn("w:pPr"))
        xml = re.sub(r"\s+", " ", pr.xml if pr is not None else "")
        xml = re.sub(r'xmlns:[a-z0-9]+="[^"]*"\s*', "", xml)
        runs = [
            (
                r.text[:12],
                r.font.name,
                r._element.xpath("string(./w:rPr/w:rFonts/@w:eastAsia)"),
                r.font.size.pt if r.font.size else None,
                r.bold,
            )
            for r in p.runs
        ]
        print("  p%d %r text=%r" % (i, p.style.name, p.text[:24]))
        print("      pPr: %s" % xml[:240])
        if runs:
            print("      runs:", runs[:3])


def main():
    """执行修复、复制与格式检查。"""
    restore()
    out = clone_template()
    d = Document(out)
    show(d.tables[0].rows[0].cells[2], "T1R0 题目")
    show(d.tables[0].rows[3].cells[1], "T1R3 选题背景及目的")
    show(d.tables[0].rows[4].cells[1], "T1R4 工作任务及要求")
    show(d.tables[1].rows[1].cells[0], "T3R1 背景与意义")


if __name__ == "__main__":
    main()
