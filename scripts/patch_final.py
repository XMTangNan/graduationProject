#!/usr/bin/env python3
"""最后一轮精确化：法律表述、技术风险、预期成果与里程碑（幂等）。"""

import copy
import pathlib

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

EDITS = [
    (1, 1, 0,
     "《民事诉讼证据的若干规定》第94条对不同存证主体的审查要求不同，本课题不主张自身属于中立第三方平台。",
     "《民事诉讼证据的若干规定》第94条列举了可以径行确认电子数据真实性的若干情形（含由中立第三方平台提供或确认），本课题不主张自身属于该类平台。",
     "第94条列举了可以径行确认"),
    (0, 4, 1,
     "删除请求通过销毁该凭证的数据密钥使原文不可恢复；",
     "删除请求通过销毁该凭证的数据密钥使原文不可恢复（密文与包裹后的数据密钥一并删除，主密钥不随业务数据备份）；",
     "主密钥不随业务数据备份"),
    (1, 2, 0,
     "整数溢出依赖Solidity 0.8的内置检查，重入由nonReentrant修饰符防范",
     "整数溢出依赖Solidity 0.8的内置检查（编译器锁定 0.8.x，FISCO BCOS 3.x 支持上限为 0.8.11），重入由自实现的 nonReentrant 修饰符防范（不依赖 OpenZeppelin 版本适配）",
     "自实现的 nonReentrant"),
    (1, 2, 0,
     "回执 PDF 由平台 secp256k1 密钥签名",
     "回执 PDF 由平台 secp256k1 密钥签名（JDK 经 BouncyCastle 提供该曲线）",
     "BouncyCastle"),
    (1, 2, 0,
     "区块时间与交易哈希证明“该存证记录自写入起未被修改”",
     "区块时间与交易哈希证明“该存证记录自写入起未被修改”（区块时间取自节点本地时钟，仅作排序与相对时间参考）",
     "节点本地时钟"),
]

TAIL = [
    ("head", "6. 预期成果与验收节点"),
    ("body", "预期成果三项：① 可运行的原型系统一套，含完整源代码与单机部署脚本（一键启动链、后端与前端）；"
             "② 毕业论文；③ 答辩PPT。"),
    ("body", "自检节点：2027年1月底 四节点链与两个合约跑通，完成一笔登记与一次核验；2月底 后端 18–20 个接口通过接口测试；"
             "3月中旬 前端 7 个页面、回执导出与在线核验完成，三类测试跑完；3月下旬 提交中期材料。"),
    ("body", "技术选型版本在概要设计说明书中固化（JDK、Spring Boot、FISCO BCOS、Solidity ≤0.8.11、OpenPDF、BouncyCastle、ZXing），"
             "后续论文与中期报告直接引用该版本基线。"),
]


def set_para_text(par, text):
    """整段替换文本，保留首个非空 run 的格式。"""
    runs = [r for r in par.runs if r.text]
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def append_like(cell, proto_par, text):
    """按原型段落格式在单元格末尾追加一段。"""
    el = copy.deepcopy(proto_par._p)
    runs = el.findall(qn("w:r"))
    keep = next((r for r in runs if r.find(qn("w:t")) is not None), runs[0] if runs else None)
    if keep is None:
        cell.add_paragraph(text)
        return
    for child in list(el):
        if child.tag == qn("w:pPr") or child is keep:
            continue
        el.remove(child)
    for junk in list(keep):
        if junk.tag not in (qn("w:rPr"), qn("w:t")):
            keep.remove(junk)
    ts = keep.findall(qn("w:t"))
    if ts:
        node = ts[0]
        for extra in ts[1:]:
            keep.remove(extra)
    else:
        node = OxmlElement("w:t")
        rpr = keep.find(qn("w:rPr"))
        if rpr is not None:
            rpr.addnext(node)
        else:
            keep.append(node)
    node.text = text
    node.set(qn("xml:space"), "preserve")
    cell._tc.append(el)


def main():
    """应用修订并保存。"""
    doc = Document(str(DOCX))
    stats = {}
    for ti, ri, ci, old, new, guard in EDITS:
        cell = doc.tables[ti].rows[ri].cells[ci]
        if guard in cell.text:
            stats["SKIP"] = stats.get("SKIP", 0) + 1
            continue
        hit = next((p for p in cell.paragraphs if old in p.text), None)
        if hit is None:
            stats["MISS"] = stats.get("MISS", 0) + 1
            print("MISS  ", old[:24])
            continue
        set_para_text(hit, hit.text.replace(old, new))
        stats["APPLIED"] = stats.get("APPLIED", 0) + 1
    cell = doc.tables[1].rows[2].cells[0]
    if "6. 预期成果与验收节点" in cell.text:
        stats["TAIL_SKIP"] = 1
    else:
        protos = [p for p in cell.paragraphs if p.text.strip()]
        head_proto = next(p for p in protos if p.text.strip().startswith("5."))
        body_proto = protos[-1]
        for kind, text in TAIL:
            append_like(cell, head_proto if kind == "head" else body_proto, text)
        stats["TAIL_ADDED"] = len(TAIL)
    doc.save(str(DOCX))
    print("saved", DOCX.stat().st_size, "bytes", stats)


if __name__ == "__main__":
    main()
