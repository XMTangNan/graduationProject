#!/usr/bin/env python3
"""按“一台电脑、尽量简单”的口径最小化收口（幂等）。不改申报表要求项。"""

import pathlib

from docx import Document

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

EDITS = [
    (1, 2, 0,
     "管理端另设两个以只读为主的页面：区块链管理展示当前区块高度、最近存证交易与状态、节点与群组列表；系统设置以枚举维护场景字典，并提供上传上限、算法标识与合约地址的配置，白名单增删复用合约 Owner 方法。",
     "管理端在同一页面下用两个只读为主的标签页组织：区块链管理标签展示当前区块高度与最近存证交易及状态；系统设置标签以枚举展示场景字典，并提供上传上限、算法标识与合约地址三项配置；登记地址白名单只在合约层配置，不单独开发界面。",
     "两个只读为主的标签页"),
    (1, 2, 0,
     "采用自签名公钥分发，可复核性来自链上摘要与交易哈希；",
     "采用自签名公钥分发，前端不单独提供验签页面，可复核性来自链上摘要与交易哈希；",
     "前端不单独提供验签页面"),
    (1, 2, 0,
     "合约部署与登记地址由管理员白名单控制。",
     "合约部署与登记地址由管理员白名单控制；单机部署时每个节点限 512 MB 堆内存，其余资源留给后端与数据库。",
     "每个节点限 512 MB 堆内存"),
    (1, 2, 0,
     "跨主机部署、真实第三方节点接入与国密算法适配列为后续工作。",
     "不做前端验签页面与白名单管理界面；跨主机部署、真实第三方节点接入与国密算法适配列为后续工作。",
     "不做前端验签页面与白名单管理界面"),
    (0, 4, 1,
     "（4）完成软件编码、测试等工作，实现系统功能，提交作品；",
     "（4）完成软件编码、测试等工作，实现系统功能，提交作品（含完整源代码与单机部署脚本）；",
     "含完整源代码与单机部署脚本"),
]


def set_para_text(par, text):
    """整段替换文本，保留首个非空 run 的字体格式。"""
    runs = [r for r in par.runs if r.text]
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def main():
    """执行修订并保存。"""
    doc = Document(str(DOCX))
    stats = {}
    for ti, ri, ci, old, new, guard in EDITS:
        cell = doc.tables[ti].rows[ri].cells[ci]
        if guard in cell.text:
            stats["SKIP"] = stats.get("SKIP", 0) + 1
            print("SKIP   ", old[:26])
            continue
        hit = next((p for p in cell.paragraphs if old in p.text), None)
        if hit is None:
            stats["MISS"] = stats.get("MISS", 0) + 1
            print("MISS   ", old[:26])
            continue
        set_para_text(hit, hit.text.replace(old, new))
        stats["APPLIED"] = stats.get("APPLIED", 0) + 1
        print("APPLIED", old[:26])
    doc.save(str(DOCX))
    print("saved", DOCX.name, DOCX.stat().st_size, "bytes", stats)


if __name__ == "__main__":
    main()
