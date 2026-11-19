#!/usr/bin/env python3
"""按审查结论就地修订《任务书与开题报告》正文（幂等，可重复运行）。

每项为 (表序号, 行序号, 列序号, 旧片段, 新片段, 幂等守卫片段)：
单元格内已出现守卫片段即跳过；否则在段落内命中旧片段后整段重写（保留首个非空 run 的格式）。
"""

import pathlib
import sys

from docx import Document

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

EDITS = [
    (0, 4, 1, "注册登录、轻量实名认证与角色权限", "注册登录、角色与权限控制", "注册登录、角色与权限控制"),
    (1, 2, 0, "按文件名排序依次喂入同一摘要算法得到一条清单聚合摘要",
     "按文件名排序，把每个文件按“文件名长度‖文件名‖摘要长度‖文件摘要”编码后依次喂入同一摘要算法，得到一条清单聚合摘要",
     "摘要长度‖文件摘要"),
    (1, 2, 0, "核验时可以定位到被替换的具体某一张。",
     "核验时可以定位到被替换的具体某一张；聚合输入对文件名与摘要均带长度前缀，避免不同文件组合产生同一摘要输入。",
     "聚合输入对文件名与摘要均带长度前缀"),
    (1, 2, 0, "用户可以查见何人在何时核验过自己的凭证。实现上有一个绕不开的取舍",
     "用户可以查见何人在何时核验过自己的凭证。核验方无须持有链上账户：交易由平台托管的核验账户代发，"
     "合约方法显式接收 verifierId，链上事件同时记录 verifierId 与交易发起地址，"
     "由平台对核验人身份与 verifierId 的对应关系背书。实现上有一个绕不开的取舍", "verifierId"),
    (1, 2, 0, "完成场景选择、凭证上传、场景化字段填写、授权签发与撤销、回执预览与下载。",
     "完成场景选择、凭证上传、场景化字段填写、授权签发与撤销、回执预览与下载。管理端另设两个以只读为主的页面："
     "区块链管理展示当前区块高度、最近存证交易与状态、节点与群组列表；系统设置以枚举维护场景字典，"
     "并提供上传上限、算法标识与合约地址的配置，白名单增删复用合约 Owner 方法。", "管理端另设两个"),
    (1, 2, 0, "系统提供的可复核性来自链上摘要与交易哈希，不引入数字证书体系；",
     "回执 PDF 由平台 secp256k1 密钥签名，公钥与指纹在核验页公布并登记上链；系统不引入 CA/PKI 证书链，"
     "采用自签名公钥分发，可复核性来自链上摘要与交易哈希；", "secp256k1"),
    (1, 2, 0, "前端不承担摘要计算与密钥处理；仅在用户需要自证原文未被替换时，把原文下载到本机后在浏览器内重新计算摘要并与链上登记值对照，原文不需要再次上传。",
     "常规流程中前端不做摘要计算与密钥处理；仅在用户需要自证原文未被替换时，把原文下载到本机后调用浏览器 "
     "WebCrypto 重新计算摘要，与链上登记值对照，原文不需要再次上传。", "WebCrypto"),
    (1, 1, 0, "依《公证法》第2条，证据公证由公证机构依申请办理",
     "依《公证法》第2条与第11条第（九）项，保全证据类公证由公证机构依申请办理", "第11条第（九）项"),
    (1, 2, 0, "需求分析（场景访谈、民事要件清单、用例与角色权限矩阵",
     "需求分析（场景调研、民事要件清单、用例与角色权限矩阵", "场景调研"),
    (1, 1, 0, "《民事诉讼证据的若干规定》第94条对中立第三方平台记录的证据另有审查要求。本系统据此把目标限定在",
     "《民事诉讼证据的若干规定》第94条对不同存证主体的审查要求不同，本课题不主张自身属于中立第三方平台。"
     "本系统把目标限定在", "不主张自身属于中立第三方平台"),
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


def apply_edit(doc, ti, ri, ci, old, new, guard):
    """在指定单元格内应用一条修订，带幂等守卫。"""
    cell = doc.tables[ti].rows[ri].cells[ci]
    if guard and guard in cell.text:
        return "SKIP"
    for par in cell.paragraphs:
        if old in par.text:
            set_para_text(par, par.text.replace(old, new))
            return "APPLIED"
    return "MISS"


def main():
    """执行全部修订并保存。"""
    src = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else DOCX
    doc = Document(str(src))
    stats = {}
    for ti, ri, ci, old, new, guard in EDITS:
        res = apply_edit(doc, ti, ri, ci, old, new, guard)
        stats[res] = stats.get(res, 0) + 1
        print("%-8s T%dR%dC%d  %s" % (res, ti, ri, ci, old[:30]))
    doc.save(str(src))
    print("saved:", src.name, "bytes:", src.stat().st_size, "stats:", stats)


if __name__ == "__main__":
    main()
