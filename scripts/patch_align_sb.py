#!/usr/bin/env python3
"""把任务书·工作任务及要求重建为与《选题审批表》一一对应的三段式。

申报表原文要求的功能模块为：用户管理；四大民事场景存证管理；区块链管理；智能合约授权核验；
系统设置；存证报告导出。过程与产出为固定的 6 条。上一轮单机口径改写时把模块名与 6 条产出
压掉了，会造成任务书与申报表要素不一致，本脚本补回并保留已定的单机实现口径。

做法与 fill_renwu_kaiti.py 相同：克隆单元格内已有段落 XML 继承字体与缩进，整段重写目标单元格。
"""

import copy
import datetime
import pathlib
import shutil

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

BLOCKS = [
    ("head", "功能"),
    ("body", "主要工作包括：用户管理（注册登录、轻量实名认证与角色权限）；"
              "四大民事场景存证管理（民间借贷、房屋/物品租赁、普通民间简易民事协议、个人原创内容版权，支持图片与PDF等多媒体凭证的上传、字段填写与预览下载）；"
              "区块链管理；智能合约授权核验（链上签发带有效期与次数的授权，可撤销，核验留痕）；"
              "系统设置；存证报告导出（含交易哈希与摘要的回执及在线核验）。"),
    ("blank", ""),
    ("head", "过程与产出"),
    ("item", "（1）查阅课题相关资料，对课题涉及的相关技能进行学习，提交可行性分析报告、英文文献翻译、文献综述；"),
    ("item", "（2）深度调研，完成系统的需求分析，提交需求规格说明书；"),
    ("item", "（3）完成概要设计和详细设计，提交概要说明书、数据库设计说明书，或研究型论文进展报告；"),
    ("item", "（4）完成软件编码、测试等工作，实现系统功能，提交作品；"),
    ("item", "（5）论文撰写，提交毕业论文；"),
    ("item", "（6）答辩准备，提交答辩PPT。"),
    ("blank", ""),
    ("head", "要求"),
    ("body", "基于FISCO BCOS、Spring Boot与Vue完成系统实现；原文资料由MySQL登记管理并逐文件以AES-GCM加密落盘，链上只写摘要与凭证信息；"
             "个人信息依《个人信息保护法》履行告知同意，删除请求通过销毁该凭证的数据密钥使原文不可恢复；合约以修饰符控制登记与核验权限，"
             "接口须通过越权、篡改与并发测试。产品规划及功能设计合理可靠，系统具备易用性、可靠性与稳定性。"),
    ("body", "按人工智能学院2027届工作计划表节点推进，论文符合学校规范化要求；开发与测试在单台计算机上完成。"
             "系统须在登录页与回执中标注：本系统为本科教学科研原型，不属于《公证法》第2条规定的公证行为，不具备法定公证效力。"),
]


def first_text_run(p_el):
    """返回段落中优先含 w:t 的 run，空段落退化为第一个 run。"""
    runs = p_el.findall(qn("w:r"))
    for r in runs:
        if r.find(qn("w:t")) is not None:
            return r
    return runs[0] if runs else None


def make_para(proto_el, text):
    """克隆原型段落并整段写入文本，保留 pPr 与 run 字符格式。"""
    el = copy.deepcopy(proto_el)
    keep = first_text_run(el)
    if keep is None:
        raise RuntimeError("原型段落没有 run")
    for child in list(el):
        if child.tag != qn("w:pPr") and child is not keep:
            el.remove(child)
    for junk in list(keep):
        if junk.tag not in (qn("w:rPr"), qn("w:t")):
            keep.remove(junk)
    ts = keep.findall(qn("w:t"))
    for extra in ts[1:]:
        keep.remove(extra)
    if not ts:
        node = OxmlElement("w:t")
        rpr = keep.find(qn("w:rPr"))
        if rpr is not None:
            rpr.addnext(node)
        else:
            keep.append(node)
        ts = [node]
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")
    return el


def main():
    """重建任务书·工作任务及要求单元格。"""
    doc = Document(str(DOCX))
    book = doc.tables[0]
    c_task = book.rows[4].cells[1]
    c_bg = book.rows[3].cells[1]
    body_protos = [p._p for p in c_task.paragraphs if p.text.strip()]
    if not body_protos:
        raise RuntimeError("任务及要求单元格无正文段可作为格式原型")
    heads = [p._p for p in c_bg.paragraphs if p.text.strip() in ("选题背景", "主要目的")]
    protos = {
        "head": copy.deepcopy(heads[0] if heads else body_protos[0]),
        "body": copy.deepcopy(body_protos[0]),
        "item": copy.deepcopy(body_protos[0]),
        "blank": copy.deepcopy(body_protos[0]),
    }
    tc = c_task._tc
    for p in list(tc.findall(qn("w:p"))):
        tc.remove(p)
    for kind, text in BLOCKS:
        tc.append(make_para(protos[kind], text))
    doc.core_properties.modified = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)
    doc.save(str(DOCX))
    print("段落写入:", len(BLOCKS), "| 已就地保存（不生成备份）:", DOCX.name, "bytes=", len(DOCX.read_bytes()))


if __name__ == "__main__":
    main()
