#!/usr/bin/env python3
"""按用户批注（2026-11-19 15:0x，共 7 条）修订文档表述；就地替换，不生成备份，幂等。

对应关系：
  批注0/1/5  存储与密钥职责不清、密钥来源与管理者未写  -> 明确"密文落盘 / MySQL 存元数据 / DEK 包裹入库 / KEK 由部署者启动时注入"
  批注2      核验是否只读（与"必须发交易"矛盾）        -> 区分"使用授权的核验=发交易扣次留痕"与"回执事后复核=只读、不消费授权"
  批注3      "比对结果回写链上"与技术路径不一致        -> 比对结果作为核验交易的合约事件字段上链；回执摘要另发一笔独立登记
  批注4      三个派生问题需解释                        -> 就地给出定义（幂等令牌、以合约串行扣减为准、本地消息表+事件回灌）
  批注6      选型口水话不该出现在开题报告            -> 删除 JDK 实测/曲线备选讨论，只留实现结论，理由移入 rule.md
"""

import datetime
import pathlib
import re
import sys

from docx import Document

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

# 段内替换：(定位锚点, 旧片段, 新片段, 关联批注)
SUBS = [
    # ---- 批注0：任务书·要求段，存储职责写清 ----
    ("基于FISCO BCOS、Spring Boot与Vue完成系统实现",
     "原文资料由MySQL登记管理并逐文件以AES-GCM加密落盘",
     "原文文件逐文件以AES-GCM加密后写入与数据库分离的磁盘目录，MySQL只登记其路径、摘要等元数据", "批注0"),
    # ---- 批注1/2/3：基本思路段 ----
    ("系统按“自己保全—摘要上链—本人授权—出证可复核”",
     "原始文件逐个用随机数据密钥加密后落盘，MySQL存放路径、文件元数据与凭证字段，主密钥单独保管。",
     "原始文件逐个用随机生成的数据密钥以AES-GCM加密，密文写入磁盘目录，MySQL只存放路径、逐文件摘要与凭证字段等元数据；"
     "数据密钥本身经主密钥加密后随元数据入库，主密钥由平台部署者在服务启动时注入、不入库不入代码仓库。", "批注1"),
    ("系统按“自己保全—摘要上链—本人授权—出证可复核”",
     "核验方凭用户签发的授权发起核验，核验本身作为一笔链上交易留痕；比对结果连同回执摘要回写链上，使回执也可以复核。",
     "核验方凭用户签发的授权发起核验：该动作必须作为一笔链上交易写入合约，合约事件同时记录核验时刻、verifierId 与本次比对结论"
     "（一致或不一致），并消费一次授权配额；回执内容的摘要另作一条独立存证记录上链，供事后以只读方式复核，不再消费授权。", "批注2/3"),
    # ---- 批注4：难点三 ----
    ("授权的可验证性是第三难点",
     "由此派生三个要处理的问题：同一核验方重复提交的幂等性、并发核验时剩余次数的原子扣减、核验交易失败后前后端状态的一致性。",
     "由此派生三个要处理的问题：同一核验方重复提交的幂等性，用一次性核验令牌作为判重键，重复请求不再扣次；"
     "并发核验时剩余次数的原子扣减，扣减只在合约内完成、不在后端做“先查后改”；"
     "核验交易失败后前后端状态的一致性，用本地消息表重试并以链上事件为准回灌记录。", "批注4"),
    # ---- 批注5：后端技术路线 ----
    ("后端：Spring Boot",
     "主密钥单独保管，不与业务数据同处备份。",
     "主密钥由平台部署者（本课题中由开发者以管理员角色担任）在服务启动时经环境变量或权限受限的密钥文件注入，仅驻留内存，"
     "不写入数据库、不进入代码仓库、不与业务数据同一份备份。", "批注5"),
    # ---- 批注3：核验页只读说明 ----
    ("出证与验证：回执",
     "核验页凭回执上的存证编号比对链上登记值是否一致。",
     "核验页凭回执上的存证编号以只读方式比对链上登记值是否一致，该复核步骤不消费授权次数、不产生交易。", "批注3"),
]

# 批注6：整句删除选型口水话（正则锚定，容许用户微调）
CUT6 = (re.compile(r"回执PDF用ECDSA.*?(?=系统不引入CA/PKI)"),
        "回执PDF由平台使用ECDSA对回执文件的SHA-256签名，")


def paragraphs(doc):
    """返回正文与表格内全部段落。"""
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


def apply_para(p, old, new):
    """在段落内跨 run 替换文本，格式随首个 run 保留。"""
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
        pre = t[:max(0, i - s)]
        post = t[max(0, min(len(t), end - s)):]
        r.text = pre + (new if s <= i < e else "") + post
    return True


def main():
    """执行批注修订，幂等；就地保存，不生成备份。"""
    dry = "--dry-run" in sys.argv
    doc = Document(str(DOCX))
    ps = paragraphs(doc)
    log = []
    for anchor, old, new, tag in SUBS:
        hit = None
        for p in ps:
            if anchor in p.text:
                hit = p
                break
        if hit is None:
            log.append("MISS 锚点未找到 %s (%s)" % (tag, anchor[:14]))
            continue
        if new[:18] in hit.text and old not in hit.text:
            log.append("SKIP 已达标 %s" % tag)
            continue
        if apply_para(hit, old, new):
            log.append("OK   %s" % tag)
        else:
            log.append("MISS 片段未找到 %s: %s" % (tag, old[:24]))
    # 批注6
    for p in ps:
        if "回执PDF用ECDSA" in p.text:
            joined = "".join(r.text or "" for r in p.runs)
            cut, repl = CUT6
            new_text = cut.sub(repl, joined, count=1)
            if new_text != joined:
                runs = list(p.runs)
                runs[0].text = new_text
                for r in runs[1:]:
                    r.text = ""
                log.append("OK   批注6（删除选型口水话）")
            break
    else:
        log.append("SKIP 批注6 已无该表述")
    for x in log:
        print(" ", x)
    if dry:
        print("dry-run 未写入")
        return
    doc.core_properties.modified = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)
    doc.save(str(DOCX))
    print("已就地保存（不生成备份）bytes=", len(DOCX.read_bytes()))


if __name__ == "__main__":
    main()
