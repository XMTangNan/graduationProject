#!/usr/bin/env python3
"""在现有交付文件上就地做措辞修订（消除"自我否定"式表述），保留 Word 生成的元数据。

与 fill_renwu_kaiti.py 的区别：本脚本不重建任何段落，只替换命中句子所在的 run 文本，
因此用户在 Word 中的另存信息（app.xml 统计、rsid、样式等）全部保留。

用法：python3 scripts/patch_tone.py [--dry-run]
"""

import datetime
import pathlib
import shutil
import sys

from docx import Document

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

# (定位锚点, 原文片段, 替换后片段, 说明)
EDITS = [
    ("开题·背景", "3.上链不等于当然被采信", "3.区块链存证的效力边界",
     "小标题去否定式"),
    ("开题·背景", "存证系统若不交代这条边界，就会把技术上的不可篡改误当成法律上的证明力。",
     "本系统据此把目标限定在提供可复核的完整性证明与授权核验；数据是否真实、具有多大证明力，仍由人民法院依法释〔2019〕19号第93条、第94条审查认定。",
     "把泛指批评改成本系统定位+正面回收"),
    ("任务书·背景", "需要界定的是，本课题所称电子公证指以技术手段固定电子数据并供第三方核验，不属《公证法》第2条的公证行为。",
     "",
     "免责界定从背景段尾移出"),
    ("任务书·目的", "工作量与难度适合本科毕业设计。",
     "工作量与难度适合本科毕业设计。需要说明，本课题所称电子公证指以技术手段固定电子数据并供第三方核验，不属于《公证法》第2条规定的公证行为，其价值在于为民事协商、调解与诉讼提供可核验的技术凭证。",
     "免责界定挪到目的段末并补正面价值句"),
    ("开题·现状", "相关讨论涉及零知识证明、门限签名等方向，实现门槛较高，超出本科原型范围，本课题不声称覆盖。",
     "相关讨论还涉及零知识证明、门限签名等方向，实现门槛较高，本课题原型不予实现，在论文中列为后续工作。",
     "去掉“不声称覆盖”的自我削弱"),
    ("开题·重难点", "上链前的真实性是第一难点，链本身解决不了。",
     "上链前的真实性是第一难点，链自身无法保证，构成系统设计的边界条件。",
     "措辞中性化"),
    ("开题·重难点", "本系统既存原文又出回执，与《民事诉讼证据的若干规定》第94条所称中立第三方平台仍有距离。拟采取两项缓解措施",
     "本系统既存原文又出回执，尚不满足《民事诉讼证据的若干规定》第94条对中立第三方平台的要求。为此设置两项措施",
     "“仍有距离”改为“尚不满足+措施”"),
    ("开题·重难点", "这一处理在论文中作为局限性如实说明，不回避。",
     "该处理的技术前提与法律后果在论文中如实说明。",
     "删元话语“不回避”"),
    ("开题·测试", "实验设计与结果在论文中完整给出，不预设结论。",
     "实验设计与结果在论文中完整给出。",
     "删元话语“不预设结论”"),
    ("开题·意义", "形成“法律要件—数据字段—合约接口”的对应写法，比笼统做一个存证系统更能说明技术如何贴合具体法律规则。",
     "形成“法律要件—数据字段—合约接口”的对应写法，使每项技术选择都能落实到具体法律要件上。",
     "删贬低同类的比较句"),
    ("开题·范围", "5. 本课题不做的事", "5. 研究范围与边界",
     "标题不再自我弱化"),
    # ---- 第 2 轮：把"指控式缺口"改写为"缺口 + 本课题的设计要求" ----
    ("开题·背景1", "就可能让举证方拿不出可检验的原始载体。",
     "就可能让举证方拿不出可检验的原始载体。本课题据此把保全动作放在纠纷发生之前，用链上摘要固定其后的完整状态。",
     "第1条补设计要求尾句"),
    ("开题·背景2", "缺少当场可用、花费很低的手段。",
     "缺少当场可用、花费很低的手段。本课题据此把使用门槛设定为个人可自助完成，注册后即可发起存证，不必先经过机构。",
     "第2条补设计要求尾句"),
    ("开题·背景4", "4.存证字段与民事要件对不上", "4.通用存证字段与民事要件之间存在落差",
     "第4条标题去口语化与全称判断"),
    ("开题·背景4", "同类系统多用一套通用元数据，不区分场景。借贷要看借款合意、交付事实、金额与还款时间，租赁要看租期、押金与交割状态，一句笼统的“已存证”支撑不了任何要件证明。",
     "以文件为单位的常见做法只记录文件名、摘要与存证时间，不区分民事行为类型。而借贷要看借款合意、交付事实、金额与还款时间，租赁要看租期、押金与交割状态，仅“某文件在某时刻已存证”这一层信息，还不能直接对应到要件事实。本课题据此为四类场景分别设计存证字段模板。",
     "第4条正文：批评对象改为做法，去掉“任何”，末尾接对策"),
    ("开题·背景5", "5.个人信息的控制权被交给平台", "5.谁能核验我的凭证，本人说不上话",
     "第5条标题去指控式"),
    ("开题·背景5", "凭证访问权限通常只写在后端接口里，用户既不知道谁查过自己的借条，也难以撤回。《个人信息保护法》第45条、第47条规定的查阅复制与删除权利，在链上数据不可删除这一特性面前需要专门的技术安排。",
     "访问控制通常写在服务端：用户能看到自己的存证列表，却难以指定“某人可在某时间段内核验某一条凭证”，事后也不易撤回或查知核验记录。本课题据此把授权关系与核验留痕放进智能合约，使授权可见、可撤销、可追溯。",
     "第5条正文：删去自伤式合规冲突（移入重难点），只留缺口+对策"),
    ("开题·重难点", "个人敏感信息与链上不可删除的冲突是第二难点。",
     "个人敏感信息与链上不可删除的冲突是第二难点，《个人信息保护法》第45条、第47条规定的查阅复制与删除权利在这里最需要正面处理。",
     "承接第5条移出的法条引用，避免全篇丢失出处"),
]


def iter_paragraphs(doc):
    """依次产出正文段落与两张表单表格内的段落。"""
    for p in doc.paragraphs:
        yield "正文", p
    for ti, table in enumerate(doc.tables):
        for row in table.rows:
            keep, seen = [], set()
            for cell in row.cells:
                keep.append(cell)
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                for p in cell.paragraphs:
                    yield "表%d" % ti, p


def cross_replace(paragraph, old, new):
    """在段落内跨 run 替换文本，保留各 run 原有字符格式。

    Args:
        paragraph: docx 段落。
        old: 待替换片段（须完整落在该段落的 run 文本内）。
        new: 替换后的片段，可为空串。
    Returns:
        bool: 是否命中并替换。
    """
    runs = list(paragraph.runs)
    texts = [r.text or "" for r in runs]
    joined = "".join(texts)
    idx = joined.find(old)
    if idx < 0:
        return False
    end = idx + len(old)
    pos = 0
    for run, text in zip(runs, texts):
        start, stop = pos, pos + len(text)
        pos = stop
        if stop <= idx or start >= end:
            continue
        pre = text[:max(0, idx - start)]
        post = text[max(0, min(len(text), end - start)):]
        mid = new if start <= idx < stop else ""
        run.text = pre + mid + post
    return True


def main():
    """执行全部替换并保存（默认先备份原文件）。"""
    dry = "--dry-run" in sys.argv
    doc = Document(str(DOCX))
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    backup = DOCX.with_name(DOCX.stem + "-改前备份" + stamp + ".docx")
    done, missed, skipped = [], [], []
    snapshot = [p.text for _, p in iter_paragraphs(doc)]
    for tag, old, new, why in EDITS:
        if new and any(new in t for t in snapshot):
            skipped.append((tag, why))
            continue
        hit = False
        for where, paragraph in iter_paragraphs(doc):
            if old in paragraph.text and cross_replace(paragraph, old, new):
                hit = True
                done.append((tag, where, why))
                break
        if not hit:
            missed.append((tag, old[:24], why))
    for row in done:
        print("  OK   %-12s %s" % (row[0], row[2]))
    for row in skipped:
        print("  SKIP %-12s %s" % (row[0], row[1]))
    for row in missed:
        print("  MISS %-12s 未命中(可能已应用): %s" % (row[0], row[1]))
    if dry:
        print("dry-run，未写入")
        return
    if not done and not skipped:
        print("无任何替换，退出")
        return
    shutil.copy2(DOCX, backup)
    doc.save(str(DOCX))
    print("备份:", backup.name)
    print("已保存:", DOCX.name, "bytes=", len(DOCX.read_bytes()))
    print("成功 %d 处，未命中 %d 处" % (len(done), len(missed)))


if __name__ == "__main__":
    main()
