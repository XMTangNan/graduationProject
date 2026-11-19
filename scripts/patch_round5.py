#!/usr/bin/env python3
"""第五轮审查落地（幂等）：链上数据范围一致化+删除撤权、敏感字段加密范围、回执原始字节留存。"""
import pathlib
from docx import Document

P = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_面向大众民事的联盟链电子公证存证系统设计与实现.docx")

SUBS = [
(1, 2, 0, "链上只出现聚合摘要、随机凭证编号、场景类型与区块信息；借贷相对方、银行账号、金额等可识别到个人的内容留在链下加密；", "链上存证登记交易仅保存文件清单聚合摘要、随机凭证编号、场景类型等必要凭证信息；授权管理另需在合约中记录凭证标识、被授权方的假名化标识、有效期、使用次数与撤销状态；核验、查看、下载事件另记 requestId、actionType、操作时间与处理结果。借贷相对方、银行账号、金额等可识别到个人的内容不在链上；假名化标识在可经链下映射关联到个人时，仍按个人信息要求管理；", "假名化标识在可经链下映射关联到个人时"),
(1, 2, 0, "删除请求符合法定情形后，系统删除密文、包裹后的数据密钥与链下身份关联，使原文无法经系统正常业务流程恢复；", "删除请求符合法定情形后，系统删除密文、包裹后的数据密钥与链下身份关联，使原文无法经系统正常业务流程恢复，同时将该凭证在合约中标记为已失效，撤销或阻断其全部未结束的授权，此后任何核验、查看、下载请求一律拒绝；", "标记为已失效"),
(0, 4, 1, "包裹后的数据密钥、文件摘要与凭证字段同库管理，链上只写摘要与凭证信息；", "包裹后的数据密钥、文件摘要与凭证字段同库管理，其中涉及个人身份、金融账户等敏感个人信息的业务字段单独加密存储，链上只写摘要与凭证信息；", "敏感个人信息的业务字段单独加密存储"),
(1, 2, 0, "第二阶段，对最终PDF直接计算SHA-256得到回执摘要，经合约中“存证编号到回执摘要”的映射登记上链；", "第二阶段，对最终PDF直接计算SHA-256得到回执摘要，经合约中“存证编号到回执摘要”的映射登记上链；PDF原生字节与其摘要、存证编号同时留存于数据库，用户下载直接返回该份原件、不重新渲染；若上链交易超时或响应丢失，沿用同一份PDF与同一摘要重试或查证，不重新生成；", "PDF原生字节与其摘要"),
(1, 2, 0, "合约部署与登记地址由管理员白名单控制。", "合约部署与登记地址由管理员白名单控制；合约维护凭证失效标志，删除凭证时置位并阻断其未结束的授权。", "合约维护凭证失效标志"),
]


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
    st = {"A": 0, "S": 0, "M": 0}
    for ti, ri, ci, old, new, guard in SUBS:
        cell = doc.tables[ti].rows[ri].cells[ci]
        if guard in cell.text:
            st["S"] += 1
            continue
        hit = next((p for p in cell.paragraphs if old in p.text), None)
        if hit is None:
            st["M"] += 1
            print("MISS:", old[:26])
            continue
        set_par(hit, hit.text.replace(old, new))
        st["A"] += 1
        print("OK:  ", old[:26])
    doc.save(str(P))
    print("saved", P.stat().st_size, st)


if __name__ == "__main__":
    main()
