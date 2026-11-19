#!/usr/bin/env python3
"""把 WebCrypto 自校验明确标为可选辅助功能，与主流程区分（幂等）。"""

import pathlib

from docx import Document

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

OLD = ("常规流程中前端不做摘要计算与密钥处理；仅在用户需要自证原文未被替换时，"
       "把原文下载到本机后调用浏览器 WebCrypto 重新计算摘要，与链上登记值对照，原文不需要再次上传。")
NEW = ("主流程中前端不做摘要计算、不接触密钥，摘要一律由后端计算。"
       "另设一个可选的自校验功能，与主流程分开：它不参与存证与核验的必需步骤，不新增后端接口，缺少它系统仍能正常完成存证、授权与出证。"
       "该功能用于用户对平台不完全信任的场景——把已存证的原文下载回本机后，由浏览器 WebCrypto 重算 SHA-256 并与链上登记值对照，"
       "从而在不依赖服务端的前提下自证原文未被替换，原文不再上传；受浏览器安全上下文限制，该功能仅在 localhost 或 https 下生效。")
GUARD = "可选的自校验功能"


def main():
    """替换目标段落并保存。"""
    doc = Document(str(DOCX))
    cell = doc.tables[1].rows[2].cells[0]
    if GUARD in cell.text:
        print("SKIP already present")
        doc.save(str(DOCX))
        return
    hit = next((p for p in cell.paragraphs if OLD in p.text), None)
    if hit is None:
        print("MISS")
        return
    runs = [r for r in hit.runs if r.text]
    runs[0].text = hit.text.replace(OLD, NEW)
    for r in runs[1:]:
        r.text = ""
    doc.save(str(DOCX))
    print("APPLIED", DOCX.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
