#!/usr/bin/env python3
"""补齐：后端段（包裹算法/大小/状态机）、石冠彬文献、架构文档与 rule 增补。幂等。"""
import pathlib
from docx import Document

ROOT = pathlib.Path("/home/xczc/code/undergraduate/graduationProject")
P = ROOT / "doc/任务书与开题报告/2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_面向大众民事的联盟链电子公证存证系统设计与实现.docx"


def set_par(par, text):
    runs = [r for r in par.runs if r.text]
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def main():
    doc = Document(str(P))
    t = doc.tables
    c = t[1].rows[2].cells[0]
    # 1) 后端段
    OLD = "数据密钥经主密钥包裹（每次包裹同样使用独立随机IV）后与MIME类型、单文件摘要同库存储，"
    NEW = ("数据密钥以主密钥按AES-256-GCM包裹（每次包裹使用独立随机IV，密文、IV与认证标签同库存储），与MIME类型、单文件摘要同库存储；"
           "上传大小默认单文件不超过10MB、单份凭证合计不超过30MB，前后端同时校验，页面可调上限不得超过部署硬上限；"
           "链上登记与本地写库不使用分布式事务：先落库标记待登记，提交交易后按回执标记已确认，超时或响应丢失先按交易哈希与凭证编号查询原交易状态再决定重发，"
           "未获链上确认不视为存证完成，合约禁止重复登记同一编号作为幂等兜底；")
    done1 = done2 = False
    for p in c.paragraphs:
        if "不使用分布式事务" in p.text:
            done1 = True
            break
        if OLD in p.text:
            set_par(p, p.text.replace(OLD, NEW))
            done1 = True
            break
    if not done1:
        print("!! 后端段仍未命中，打印候选段开头：")
        for p in c.paragraphs:
            if "后端" in p.text[:4] or "LONGBLOB" in p.text:
                print("   >", p.text[:120])
    # 2) 文献（石冠彬/陈全真）
    cbg = t[1].rows[1].cells[0]
    if "石冠彬" in cbg.text:
        done2 = True
    else:
        for p in cbg.paragraphs:
            if "杭州互联网法院" in p.text and "和版权、供应链存证链已实际运行" in p.text:
                set_par(p, p.text.replace(
                    "和版权、供应链存证链已实际运行",
                    "和版权、供应链存证链已实际运行；司法侧的采信与审查路径已有专门研究（石冠彬、陈全真《论区块链存证电子数据的优势及司法审查路径》，"
                    "西南民族大学学报（人文社会科学版）2021年第1期），其共识是链上完整性不等于证据真实性，须区分入链前后两个阶段审查"))
                done2 = True
                break
    doc.save(str(P))
    print("docx: 后端段", done1, "文献", done2, "bytes", P.stat().st_size)

    # 3) 架构文档
    A = ROOT / "doc/方案设计/角色与数据流-架构梳理-20261119.md"
    s = A.read_text(encoding="utf-8")
    AREP = [
        ("6 status==0 才返回：链上摘要、本次比对结论、核验时刻",
         "6 合约事件按 actionType 区分：VERIFY（附 MATCH/MISMATCH 结论）/ VIEW / DOWNLOAD（仅记准许）；三类均消耗“授权使用次数”，无效授权与交易未成立不扣次\n    7 status==0 才返回：链上摘要、本次比对结论（或文件流）、核验时刻"),
        ("| 核验 | recordVerify(no, verifierId, requestId, result) | !consumed[requestId]；有效；used<quota |",
         "| 核验/访问 | recordVerify(no, verifierId, requestId, actionType, result) | !consumed[requestId]；有效；used<quota |"),
        ("| 存证登记 | register(no, digest, scenario, algo) | state==NONE；白名单 |",
         "| 存证登记 | register(no, digest, scenario, algo) | state==NONE；清单非空；白名单 |"),
        ("| 删除 | 密文行 + DEK 行 + 链下关联在同一业务事务删除；链上不动；audit_log 记过程 |",
         "| 删除 | 密文行 + DEK 行 + 链下关联在同一业务事务删除；链上不动；audit_log 记过程 |\n| 待验临时件 | 核验方上传的待验文件只在内存/临时目录瞬时处理，算完摘要即删，不入库 |"),
    ]
    m = 0
    for old, new in AREP:
        if old in s:
            s = s.replace(old, new)
        else:
            print("ARCH-MISS:", old[:26]); m += 1
    A.write_text(s, encoding="utf-8")
    print("arch done, miss", m)

    # 4) rule.md
    R = ROOT / "rule.md"
    r = R.read_text(encoding="utf-8")
    if "actionType" not in r:
        r = r.replace("合约写接口限定平台服务账户 |", "合约写接口限定平台服务账户；事件按 actionType 分 VERIFY/VIEW/DOWNLOAD，仅核验类写比对结论，"
                       "三类共耗“授权使用次数”，无效授权与交易未成立不扣次 |")
    if "空清单拒绝登记" not in r:
        r = r.replace("上链聚合摘要，库内保留逐文件摘要 |", "上链聚合摘要，库内保留逐文件摘要；空清单拒绝登记；核验方临时上传件算完摘要即删、不入库 |")
    if "石冠彬" not in r:
        r = r.replace("其余候选论文阶段逐条核实后补充。",
                       "已入文：石冠彬、陈全真《论区块链存证电子数据的优势及司法审查路径》西南民族大学学报（人文社会科学版）2021年第1期。"
                       "候选（出处未齐暂不入文）：崔世群《区块链证据真实性问题研究》。")
    R.write_text(r, encoding="utf-8")
    print("rule done")


if __name__ == "__main__":
    main()
