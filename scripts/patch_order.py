#!/usr/bin/env python3
"""A 方案（最小改动）：清单聚合摘要只由逐文件内容摘要构成，文件名与上传顺序不再影响摘要。

实测依据（本机 python3，sha256 + 大端长度前缀）：
  - 现行"文件名‖摘要"编码：同一批文件重命名后聚合摘要改变（内容未变）；顺序本身不影响。
  - 改为"仅摘要升序 + 长度前缀"：6 种排列同值、改名同值、替换/删减/重复均变、可按逐文件摘要差集定位单张。
连带修正：AES-GCM 的 AAD 由"凭证编号‖序号‖文件名"改为"凭证编号‖序号"，避免上传前后命名差异导致解密失败。

就地修改三份文件：交付 docx、doc/方案设计/*.md、rule.md。幂等，不生成备份。
"""

import datetime
import pathlib
import re
import sys

from docx import Document

ROOT = pathlib.Path("/home/xczc/code/undergraduate/graduationProject")
DOCX = ROOT / "doc/任务书与开题报告/2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_面向大众民事的联盟链电子公证存证系统设计与实现.docx"
MDF = ROOT / "doc/方案设计/角色与数据流-架构梳理-20261119.md"
RUL = ROOT / "rule.md"

# ---------------- docx：段内替换 ----------------
DOCX_SUBS = [
    ("一份凭证包含多张截图时，按文件名排序，每个字段以4字节大端长度前缀定长编码（文件名长度‖文件名‖摘要长度‖文件摘要）后依次喂入同一摘要算法，得到一条清单聚合摘要，该编码保证清单到字节流的映射是单射、字段边界无歧义；",
     "一份凭证包含多张截图时，先把每个文件各自算出SHA-256摘要，再把这些摘要按值升序排列，每项以4字节大端长度前缀定长编码（摘要长度‖摘要）后依次喂入同一摘要算法，得到一条清单聚合摘要；文件名与上传先后均不参与该编码，同一批材料不会因命名或提交顺序不同而固化成不同记录，而替换、删减或重复提交任一张都会改变聚合摘要；", "基本思路-编码定义"),
    ("聚合输入采用带长度前缀的定长编码，使不同清单不会因字段边界含糊而落到同一字节序列。",
     "聚合输入只由逐文件摘要按升序加长度前缀构成，映射单射，且不随文件重命名或上传先后而变化。", "工作重点-单射句"),
    ("前端：Vue 3 + Vite 配合组件库，完成场景选择、凭证上传、场景化字段填写、授权签发与撤销、回执预览与下载。",
     "前端：Vue 3 + Vite 配合组件库，完成场景选择、凭证上传、场景化字段填写、授权签发与撤销、回执预览与下载。上传页面须向用户说明：上传先后不影响存证结果，文件名只用于阅读与留档，聚合摘要只由文件内容决定；每个材料的发生时间作为场景字段逐条填写。", "前端-提示语"),
    ("另设一组篡改对照实验：改动原文一个字节后验证比对必然失败，并检验多文件清单中替换单张截图能否被定位。",
     "另设一组篡改对照实验：改动原文一个字节后验证比对必然失败；检验多文件清单中替换单张截图能被逐文件摘要差集定位；并验证仅重命名文件或调整上传先后时聚合摘要保持不变。", "测试-对照实验"),
]

# ---------------- md / rule：文本替换 ----------------
TEXT_SUBS = [
    (MDF, "按文件名排序（同序按 seq）", "收集 4 个逐文件摘要，按摘要值升序", "md-步骤3-a"),
    (MDF, "len‖name‖32‖digest", "len(4B 大端)‖摘要", "md-步骤3-b"),
    (MDF, "（映射单射，边界无歧义）", "（映射单射；文件名与上传顺序不参与）", "md-步骤3-c"),
    (MDF, "| 4 | — | 每文件生成随机 DEK | AES-GCM 加密，**每次新 12 字节 IV**，AAD = 凭证编号‖seq‖文件名 | — | — |",
     "| 4 | — | 每文件生成随机 DEK | AES-GCM 加密，**每次新 12 字节 IV**，AAD = 凭证编号‖seq（不含文件名，避免命名差异导致解密失败） | — | — |", "md-AAD"),
    (MDF, "③ 仅改文件名 → 聚合摘要变化（因文件名进编码，需说明是设计选择）",
     "③ 仅改文件名或调整上传先后 → 聚合摘要**不变**（文件名不进编码）", "md-对照实验"),
    (MDF, "| 改了文件名算篡改吗？ | 聚合编码包含文件名（带长度前缀），改名会改变聚合摘要；这是把\"提交时可见的文件标签\"一并绑定的设计选择，对照实验会单独说明 |",
     "| 改了文件名算篡改吗？ | 不算。聚合摘要只由逐文件内容摘要升序加长度前缀构成，文件名不参与，改名与上传先后都不改变固化结果；文件名原样记录在链下元数据与回执中供阅读，替换、删减或重复提交任一张才会改变聚合摘要 |", "md-答辩问答"),
    (RUL, "| 摘要计算 | **后端流式计算 SHA-256**；前端不做密码学。一份凭证含多文件时，按文件名排序依次喂入同一摘要算法得到**清单聚合摘要**上链，链下保留逐文件摘要 |",
     "| 摘要计算 | **后端流式计算 SHA-256**；前端不做密码学。一份凭证含多文件时，把逐文件摘要**按值升序**、每项以 4 字节大端长度前缀编码后喂入同一摘要算法得到**清单聚合摘要**上链；**文件名与上传顺序不参与编码**，链下保留逐文件摘要用于定位单张 |", "rule-摘要计算"),
    (RUL, "- 聚合摘要编码：文件按文件名排序，每个文件按\"文件名长度 ‖ 文件名 ‖ 摘要长度 ‖ 文件摘要\"编码后依次喂入同一个 SHA-256，避免不同文件组合",
     "- 聚合摘要编码（2026-11-19 改为 A 方案）：逐文件摘要按值升序，每项编码为\"摘要长度(4B 大端) ‖ 摘要\"后依次喂入同一个 SHA-256；文件名与上传顺序不参与。实测：6 种排列同值、改名同值、替换/删减/重复均变、差集可定位单张。字段边界仍无歧义（定长前缀保证映射单射）。", "rule-编码决策"),
    (RUL, "（密文与包裹后的数据密钥一并删除，主密钥不随业务数据备份）",
     "（密文与包裹后的数据密钥一并删除，主密钥不随业务数据备份）", "rule-占位"),
    (RUL, "  5e. scripts/patch_comments.py 按用户批注修订",
     "  5e. scripts/patch_comments.py 按用户批注修订", "rule-占位2"),
]

AAD_RULE_OLD = "| 原文存储 | **逐文件 AES-GCM 加密落盘**，目录与数据库分离；数据密钥经主密钥包裹后与路径、MIME、单文件摘要存 **MySQL**；主密钥独立保管 |"
AAD_RULE_NEW = ("| 原文存储 | **逐文件 AES-GCM 加密落盘**，目录与数据库分离；**每次加密新生成 12 字节 IV**，AAD = 凭证编号‖文件序号（不含文件名）；"
                "数据密钥经主密钥包裹后与路径、MIME、单文件摘要存 **MySQL**；主密钥独立保管 |")


def first_text_run(p_el):
    runs = p_el.findall.__self__ if False else None
    return None


def paragraphs(doc):
    out = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            keep, seen = [], set()
            for cell in row.cells:
                keep.append(cell)
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                out.extend(cell.paragraphs)
    return out


def replace_in_paragraph(p, old, new):
    runs = list(p.runs)
    texts = [r.text or "" for r in runs]
    joined = "".join(texts)
    i = joined.find(old)
    if i < 0:
        return False
    end = i + len(old)
    pos = 0
    for r, t in zip(runs, texts):
        s, e = pos, pos + len(t)
        pos = e
        if e <= i or s >= end:
            continue
        r.text = t[:max(0, i - s)] + (new if s <= i < e else "") + t[max(0, min(len(t), end - s)):]
    return True


def patch_docx(dry):
    doc = Document(str(DOCX))
    ps = paragraphs(doc)
    log = []
    for old, new, tag in DOCX_SUBS:
        hit = next((p for p in ps if old in p.text), None)
        if hit is None:
            log.append("MISS %s" % tag)
            continue
        if new.startswith(old):
            if new[len(old):][:24] in hit.text:
                log.append("SKIP %s" % tag)
                continue
        elif new[:24] in hit.text and old not in hit.text:
            log.append("SKIP %s" % tag)
            continue
        log.append(("OK   " if replace_in_paragraph(hit, old, new) else "MISS ") + tag)
    if not dry and any(x.startswith("OK") for x in log):
        doc.core_properties.modified = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)
        doc.save(str(DOCX))
    return log


def patch_text(path, pairs, dry):
    txt = path.read_text(encoding="utf8")
    log = []
    for old, new, tag in pairs:
        if old == new:
            log.append("SKIP %s (%s)" % (tag, path.name[:12]))
            continue
        if new.startswith(old) and new[len(old):][:20] in txt:
            log.append("SKIP %s (%s)" % (tag, path.name[:12]))
            continue
        if old not in txt:
            log.append("MISS %s (%s)" % (tag, path.name[:12]))
            continue
        txt = txt.replace(old, new, 1)
        log.append("OK   %s (%s)" % (tag, path.name[:12]))
    if not dry and any(l.startswith("OK") for l in log):
        path.write_text(txt, encoding="utf8")
    return log


def main():
    dry = "--dry-run" in sys.argv
    out = patch_docx(dry)
    md_pairs = [(o, n, t) for (pth, o, n, t) in TEXT_SUBS if pth is MDF]
    rul_pairs = [(o, n, t) for (pth, o, n, t) in TEXT_SUBS if pth is RUL and o != n]
    rul_pairs.append((AAD_RULE_OLD, AAD_RULE_NEW, "rule-AAD"))
    out += patch_text(MDF, md_pairs, dry)
    out += patch_text(RUL, rul_pairs, dry)
    print("\n".join("  " + x for x in out))
    print("dry-run" if dry else "已就地保存（不生成备份）")


if __name__ == "__main__":
    main()
