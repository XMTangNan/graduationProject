#!/usr/bin/env python3
"""把任务书与开题报告 docx 导出为纯文本镜像（.md），供核对与灾后重建使用。

用法：
    python3 scripts/dump_mirror.py [docx路径] [镜像输出路径...]
两个参数都可省略：默认读交付文件，默认写到同目录以及 D 盘同一子目录。
"""

import pathlib
import sys

from docx import Document

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)
NAMES = ["一、本科毕业论文（设计）任务书", "二、本科毕业论文（设计）开题报告"]


def to_text(docx_path):
    """按表格顺序导出表单正文文本。

    Args:
        docx_path: docx 文件路径。
    Returns:
        str: Markdown 文本。
    """
    doc = Document(str(docx_path))
    out = ["# 任务书与开题报告（正文文本镜像）", "", "> 与同名 docx 内容一致，仅作核对与灾后重建用。"]
    for ti, table in enumerate(doc.tables):
        out.append("")
        out.append("## " + NAMES[ti])
        for row in table.rows:
            seen, cells = set(), []
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                cells.append(cell)
            if len(cells) == 1:
                paras = [p.text.strip() for p in cells[0].paragraphs if p.text.strip()]
                if paras:
                    out.append("")
                    out.append("**" + paras[0] + "**")
                    out.extend(paras[1:])
            else:
                label = cells[0].text.strip().replace("\n", "")
                body = " ".join(c.text.strip() for c in cells[1:] if c.text.strip())
                if len(body) <= 60:
                    out.append("")
                    out.append("- **%s**：%s" % (label, body))
                else:
                    out.append("")
                    out.append("- **%s**" % label)
                    for c in cells[1:]:
                        for p in c.paragraphs:
                            if p.text.strip():
                                out.append("  " + p.text.strip())
    return "\n".join(out) + "\n"


def main():
    """导出镜像并写入目标路径。"""
    src = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else DOCX
    text = to_text(src)
    targets = [pathlib.Path(x) for x in sys.argv[2:]] or [src.with_name(src.stem + "-正文镜像.md")]
    for t in targets:
        t.parent.mkdir(parents=True, exist_ok=True)
        t.write_text(text, encoding="utf-8")
        print("MD %6d 字符 -> %s" % (len(text), t))


if __name__ == "__main__":
    main()
