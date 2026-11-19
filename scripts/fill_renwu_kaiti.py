#!/usr/bin/env python3
"""按申报表填写《本科毕业论文（设计）任务书与开题报告》，正文已按评审意见修订。

修订要点：
    1. 字数压回学院批注要求（任务书·背景及目的≈300字，任务及要求≈250字，开题两节≥300字）；
    2. 删除《民法典》第501条误引，改按第469条（数据电文视为书面形式）+《民事诉讼法》第66条
       （电子数据为独立证据种类）+《民事诉讼证据的若干规定》第14/15/93/94条 + 《人民法院在线
       诉讼规则》第16条（仅推定上链后未篡改）表述；
    3. 补《公证法》第2条、第36条的概念界定，明确系统不是法定公证；
    4. 摘要一律为用户浏览器本地计算，服务端不代算；删除"离线保存"等错误措辞；
    5. 授权机制从 bearer 授权码改为带有效期、限次、可撤销、核验留痕的链上授权关系；
    6. 回执加平台电子签名、回执摘要上链与在线核验入口；补可信时间戳（RFC 3161）与国密选型理由；
    7. 交代联盟节点构成与平台"既存原文又出证"的中立性局限；
    8. 四类场景统一为：民间借贷、租赁、简易民事协议、原创内容版权；
    9. 删除"闭环/落地/同质化/平民化/可复用/基础设施/重复风险低"等套话，任务书与开题不重复表述。

段落格式通过 deepcopy 克隆模板内已有段落 XML 继承，不使用 paragraph.text 赋值。
"""

import copy
import re
import pathlib

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

TITLE = "面向大众民事的联盟链电子公证存证系统设计与实现"
SCENES = "民间借贷、租赁、简易民事协议、原创内容版权"

BOOK_BG = [
    ("head", "选题背景"),
    ("body", "本课题来源于社会（行业）实践。《民法典》第469条把可以随时调取查用的数据电文视为书面形式，《民事诉讼法》第66条将电子数据列为独立证据种类，微信聊天、转账截图、电子借条因此具备进入诉讼的资格。难的是保住它们：这类材料只由当事人自己持有，删改、丢失后难以自证，民间借贷、租房押金、口头约定一旦起纠纷，普通人缺少低成本的中立保全手段——公证保全收费高、流程长，面向机构的存证平台又接不上个人场景。"),
    ("head", "主要目的"),
    ("body", "针对" + SCENES + "四类场景梳理各自的民事要件，转成不同的链上存证字段；实现用户浏览器本地计算摘要、FISCO BCOS固化摘要与凭证信息、经本人授权方可由第三方核验、导出带签名回执这一流程，并完成原型系统与论文。所需规范文本、开源联盟链和教学算力均可获得，工作量与难度适合本科毕业设计。需要说明，本课题所称电子公证指以技术手段固定电子数据并供第三方核验，不属于《公证法》第2条规定的公证行为，其价值在于为民事协商、调解与诉讼提供可核验的技术凭证。"),
]

BOOK_TASK = [
    ("body", "角色与功能：设存证用户、被授权的核验方、平台管理员三类角色。用户上传场景凭证后由浏览器本地计算SHA-256摘要（预留SM3接口），原文存MySQL，链上只写摘要、去标识化的凭证编号、场景类型与时间戳；用户可签发带有效期和次数限制的核验授权，随时撤销，核验行为写链留痕；系统输出加盖平台电子签名、回执摘要同样上链的民事存证回执，并提供在线核验入口。"),
    ("body", "非功能要求：个人信息依《个人信息保护法》履行告知同意并加密存储，删除请求通过销毁链下密钥使原文不可恢复；合约须防范越权登记、重入与整数溢出，接口应通过越权、篡改与并发测试。"),
    ("body", "预期成果：可运行演示系统一套（源码、合约、数据库脚本）与毕业论文一份。技术栈为FISCO BCOS 3.x（Solidity合约、Java SDK）、Spring Boot、Vue、MySQL。"),
]

KT_BG = [
    ("head", "选题背景"),
    ("body", "民事活动的证据形态已经在近十几年里整体转为电子数据。与普通人关系最密切的" + SCENES + "四类事情，几乎全部依靠微信聊天、转账截图与线上文档留存。围绕规范文本、公开报道和现有存证产品的调研显示，这类数据在法律上的地位是清楚的，麻烦在于怎么把它保住："),
    ("blank", ""),
    ("sub", "1.有证据资格，却难以自证完整"),
    ("body", "《民法典》第469条承认数据电文的书面形式，《民事诉讼法》第66条第1款第（五）项把电子数据单列为证据种类；《最高人民法院关于民事诉讼证据的若干规定》（法释〔2019〕19号）第93条则要求结合生成、存储、传输的全过程审查其真实性。聊天记录握在当事人一方手里，一次误删、一次换机或一次账号注销，就可能让举证方拿不出可检验的原始载体。本课题据此把保全动作放在纠纷发生之前，用链上摘要固定其后的完整状态。"),
    ("blank", ""),
    ("sub", "2.面向个人的中立保全渠道基本空缺"),
    ("body", "依《公证法》第2条，证据公证由公证机构依申请办理，其收费、材料与时限对个人偏重；行业内的区块链存证产品多对接法院、仲裁与版权机构，服务对象是机构。普通人在借钱给同事、退租前交割家具、与朋友口头约定合作时，缺少当场可用、花费很低的手段。本课题据此把使用门槛设定为个人可自助完成，注册后即可发起存证，不必先经过机构。"),
    ("blank", ""),
    ("sub", "3.区块链存证的效力边界"),
    ("body", "《人民法院在线诉讼规则》（法释〔2021〕12号）第16条只推定区块链存储的数据在技术上链后未经篡改，上链之前是否真实仍要结合取证方式与原始载体判断；《民事诉讼证据的若干规定》第94条对中立第三方平台记录的证据另有审查要求。本系统据此把目标限定在提供可复核的完整性证明与授权核验；数据是否真实、具有多大证明力，仍由人民法院依法释〔2019〕19号第93条、第94条审查认定。"),
    ("blank", ""),
    ("sub", "4.通用存证字段与民事要件之间存在落差"),
    ("body", "以文件为单位的常见做法只记录文件名、摘要与存证时间，不区分民事行为类型。而借贷要看借款合意、交付事实、金额与还款时间，租赁要看租期、押金与交割状态，仅“某文件在某时刻已存证”这一层信息，还不能直接对应到要件事实。本课题据此为四类场景分别设计存证字段模板。"),
    ("blank", ""),
    ("sub", "5.谁能核验我的凭证，本人说不上话"),
    ("body", "访问控制通常写在服务端：用户能看到自己的存证列表，却难以指定“某人可在某时间段内核验某一条凭证”，事后也不易撤回或查知核验记录。本课题据此把授权关系与核验留痕放进智能合约，使授权可见、可撤销、可追溯。"),
    ("blank", ""),
    ("head", "选题意义"),
    ("body", "实用价值在于把保全做在纠纷之前，降低普通人在协商、调解与起诉时的举证门槛，也让法释〔2021〕12号第16条的技术推定有明确的适用对象。"),
    ("body", "方法价值在于把民法典与证据规则中的要件逐条翻译成链上字段和合约权限，形成“法律要件—数据字段—合约接口”的对应写法，使每项技术选择都能落实到具体法律要件上。"),
    ("body", "权限价值在于把自主授权、限次核验与核验留痕写进智能合约，使个人对自己民事凭证的控制成为可验证的技术事实，而不只是一句隐私政策承诺。"),
    ("blank", ""),
    ("head", "相关研究与产品现状"),
    ("body", "国外工作大致沿三条线：以哈希锚定与时间戳证明文件在某时刻前已经存在；用可验证凭证一类标准支持跨机构出示与校验；在核验的同时保护内容隐私。相关讨论还涉及零知识证明、门限签名等方向，实现门槛较高，本课题原型不予实现，在论文中列为后续工作。"),
    ("body", "国内以司法与行业需求为主要牵引，互联网法院的区块链存证平台和版权、供应链存证链已实际运行，底层多选用国产开源联盟链，在国密算法支持方面积累较多。公开可见的产品普遍服务企业端与诉讼环节，面向个人日常民事行为、又把授权核验做进合约层的轻量方案少见。"),
    ("body", "两相对照，遗留问题集中在三处：按民事行为类型细分的存证字段缺少设计；用户授权与撤销多在服务端实现，缺少链上可验证的表达；上链前真实性、个人敏感信息与链上不可删除之间的张力缺少正面处理。本课题围绕这三处展开。"),
]

KT_IDEA = [
    ("head", "1. 基本思路"),
    ("body", "系统按“自己保全—摘要上链—本人授权—出证可复核”四个动作组织。原始文件不进链：用户在浏览器里用Web Crypto对文件计算SHA-256，把摘要连同场景类型、去标识化凭证编号和业务时间组成一条存证记录，经智能合约写入FISCO BCOS；原文与个人信息加密后存MySQL，服务端只登记、不代替用户重算摘要，以免平台单方面替换原文。核验方凭用户签发的授权读取链上摘要并与原文比对，比对结果连同回执摘要回写链上，使回执本身也可以复核。"),
    ("blank", ""),
    ("head", "2. 开发过程"),
    ("body", "按软件工程标准过程推进：需求分析（场景访谈、民事要件清单、用例与角色权限矩阵，产出需求规格说明书）；概要与详细设计（链上链下数据边界、四类场景字段模板、合约接口、数据库表结构，产出概要说明书与数据库设计说明书）；编码实现（联盟链环境与合约、后端、前端分模块推进）；测试（功能、越权与篡改、并发与稳定性）；文档与答辩材料。可行性分析报告、英文文献翻译与文献综述在开题前后完成。"),
    ("blank", ""),
    ("head", "3. 工作重点与重难点"),
    ("body", "场景字段模板是业务侧重点。借贷记借款合意、交付事实、金额与还款约定；租赁记租期、押金与交割清单；简易民事协议记约定事项、参与方与履行状态；原创内容版权记作品摘要、创作完成时间与权属声明。字段既要能对应要件证明，又不能把敏感明文写进链上。"),
    ("body", "上链前的真实性是第一难点，链自身无法保证，构成系统设计的边界条件。可行的处理是把它降级为可核查的声明：系统记录提交存证的终端信息、文件原始元数据和用户声明的业务时间，并对接遵循RFC 3161的可信时间戳服务机构对存证记录加盖时间戳，使“何时提交”有据可查；业务发生时间本身标注为当事人声明，由法院依法释〔2019〕19号第93条综合审查。"),
    ("body", "个人敏感信息与链上不可删除的冲突是第二难点，《个人信息保护法》第45条、第47条规定的查阅复制与删除权利在这里最需要正面处理。链上只出现摘要、随机凭证编号、场景类型与时间戳；借贷相对方、金额等留在链下并加密，需要举证时由用户解密提供。用户行使删除权时销毁链下密钥，原文随即不可恢复，链上残留的是一条无法反推身份的摘要。该处理的技术前提与法律后果在论文中如实说明。"),
    ("body", "授权的可验证性是第三难点。授权关系写在合约里：凭证编号对应一组被授权身份，附带有效期、可用次数与撤销标志，每次核验产生一条链上事件记录，用户可以查见何人在何时核验过自己的凭证。"),
    ("body", "平台中立性是第四难点。本系统既存原文又出回执，尚不满足《民事诉讼证据的若干规定》第94条对中立第三方平台的要求。为此设置两项措施：回执由平台与联盟内见证节点分别签名，回执摘要上链并支持在线复核；论文中如实写明本系统是技术保全工具，证明力由人民法院依第93条、第94条审查认定。"),
    ("blank", ""),
    ("head", "4. 技术路线"),
    ("body", "区块链层：FISCO BCOS 3.x，PBFT共识的本地联盟链，节点由学院实验室、指导教师课题组与两处合作单位各承担一个（合作节点以实际可落实为限，落实不了则在学院内以多实例模拟并在论文中说明）；合约部署地址用白名单控制。合约用Solidity编写并部署在EVM兼容链上，Java只作为java-sdk侧的业务调用语言，不表述为用Java编写合约。摘要算法默认SHA-256，同时预留SM3（GB/T 32905-2016）与SM2（GB/T 32918）接口：原型阶段用SHA-256与ECDSA便于和开源工具链对接，若需对接司法或公证行业链的国密要求可以切换，合约以hashAlgo字段记录所用算法。"),
    ("body", "后端：Spring Boot，按controller、service、repository、dto分层，MySQL存储用户、原文与凭证元数据，RESTful接口以JWT鉴权；文件落盘目录与数据库分离，敏感字段以对称密钥加密，密钥独立管理以支持密钥销毁。"),
    ("body", "前端：Vue 3 + Vite，完成场景选择、凭证上传、本地摘要计算（Web Crypto API）、场景化字段填写、授权签发与撤销、在线核验与回执下载。"),
    ("body", "出证与验证：回执由后端模板生成，含存证编号、链上交易哈希、场景字段、时间戳与免责说明；对回执内容做平台电子签名，参照《电子签名法》第13条、第14条对可靠电子签名的要求；回执摘要另发一笔交易上链；核验页凭回执上的存证编号比对链上摘要与链下原文摘要是否一致。"),
    ("body", "测试：合约层在本地链上用脚本化用例覆盖越权登记、重复存证、授权过期与撤销、重入与溢出；系统层用JUnit与接口测试覆盖功能与并发；另设一组篡改对照实验，改动原文一个字节后验证比对必然失败，实验设计与结果在论文中完整给出。"),
    ("blank", ""),
    ("head", "5. 研究范围与边界"),
    ("body", "不自研共识机制与密码算法，不涉及代币与金融业务，不宣称具备法定公证效力；论文中涉及证明力的表述以现行条文和在线诉讼规则为限，其余列为后续方向。"),
]

KT_PLAN = [
    "2026年11月 课题调研与资料收集，撰写可行性分析报告，接收毕业论文（设计）任务书",
    "2026年12月 撰写开题报告，12月23日前交指导教师审核；完成英文文献翻译",
    "2027年01月 完成开题；提交文献综述；完成需求分析与四类场景字段模板设计",
    "2027年01月-2027年03月 总体设计与数据库设计，搭建联盟链，编写并部署存证与授权合约，实现后端框架与关键接口，论文初期撰写",
    "2027年03月 中期检查，汇报进度与质量并按意见整改",
    "2027年03月-2027年04月 完成前端、授权撤销与回执签名、在线核验，开展功能与安全测试及篡改对照实验，论文撰写并通过学术行为检测",
    "2027年04月-2027年05月 指导教师评阅与评阅教师评阅，配合学院论文抽检并整改",
    "2027年05月 参加分组答辩，完成成绩评定",
    "2027年06月 成绩上报；2027年6月至9月按学院要求整理归档材料",
]


def first_text_run(p_el):
    """返回段落中优先含 w:t 的 run；空段落退化为第一个 run，均无则返回 None。"""
    runs = p_el.findall(qn("w:r"))
    for r in runs:
        if r.find(qn("w:t")) is not None:
            return r
    return runs[0] if runs else None


def ensure_text_node(run_el):
    """确保 run 内含一个 w:t 并返回它（缺失时在 rPr 之后插入）。"""
    ts = run_el.findall(qn("w:t"))
    for extra in ts[1:]:
        run_el.remove(extra)
    if ts:
        return ts[0]
    node = OxmlElement("w:t")
    rpr = run_el.find(qn("w:rPr"))
    if rpr is not None:
        rpr.addnext(node)
    else:
        run_el.append(node)
    return node


def make_para(proto_el, text):
    """克隆原型段落并写入文本，保留 pPr 与 run 的字体格式。"""
    el = copy.deepcopy(proto_el)
    keep = first_text_run(el)
    if keep is None:
        raise RuntimeError("原型段落没有可写入的 run")
    for child in list(el):
        if child.tag == qn("w:pPr") or child is keep:
            continue
        el.remove(child)
    for junk in list(keep):
        if junk.tag not in (qn("w:rPr"), qn("w:t")):
            keep.remove(junk)
    node = ensure_text_node(keep)
    node.text = text
    node.set(qn("xml:space"), "preserve")
    return el


def set_title_cell(cell, text):
    """把题目单元格文本替换为 text，保留原 run 格式。"""
    p = cell.paragraphs[0]._p
    keep = first_text_run(p)
    for r in list(p.findall(qn("w:r"))):
        if r is not keep:
            p.remove(r)
    node = ensure_text_node(keep)
    node.text = text
    node.set(qn("xml:space"), "preserve")


def rewrite(cell, protos, blocks, keep_label):
    """用 blocks 重写单元格正文。

    Args:
        cell: 目标单元格。
        protos: 段落原型字典，键为 head/body/item/sub/blank/line。
        blocks: [(kind, text), ...]；元素为纯字符串时使用 line 原型。
        keep_label: 是否保留原第一段（表单标题行）。
    Returns:
        int: 新写入的段落数。
    """
    tc = cell._tc
    label_el = copy.deepcopy(cell.paragraphs[0]._p) if keep_label else None
    for p in list(tc.findall(qn("w:p"))):
        tc.remove(p)
    if label_el is not None:
        tc.append(label_el)
    for item in blocks:
        kind, text = ("line", item) if isinstance(item, str) else item
        tc.append(make_para(protos[kind], text))
    return len(blocks)


def cjk(text):
    """统计中文字符数，用于对照学院批注的字数要求。"""
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def main():
    """填写任务书与开题报告的全部正文单元格，并打印各节字数。"""
    doc = Document(str(OUT))
    t_book, t_kaiti = doc.tables[0], doc.tables[1]

    c_title_book = t_book.rows[0].cells[2]
    c_bg = t_book.rows[3].cells[1]
    c_task = t_book.rows[4].cells[1]
    c_title_kt = t_kaiti.rows[0].cells[1]
    c_kt_bg = t_kaiti.rows[1].cells[0]
    c_kt_idea = t_kaiti.rows[2].cells[0]
    c_kt_plan = t_kaiti.rows[3].cells[0]

    def ps(cell):
        return [p._p for p in cell.paragraphs]

    bp = ps(c_bg)
    protos_bg = {"head": copy.deepcopy(bp[0]), "body": copy.deepcopy(bp[1]), "blank": copy.deepcopy(bp[4]), "item": copy.deepcopy(bp[6])}
    tp = ps(c_task)
    protos_task = {"head": copy.deepcopy(tp[0]), "body": copy.deepcopy(tp[1]), "blank": copy.deepcopy(tp[2]), "item": copy.deepcopy(tp[4])}
    kp = ps(c_kt_bg)
    protos_ktbg = {"blank": copy.deepcopy(kp[1]), "head": copy.deepcopy(kp[2]), "body": copy.deepcopy(kp[3]), "sub": copy.deepcopy(kp[5])}
    ip = ps(c_kt_idea)
    protos_ktid = dict(protos_ktbg)
    protos_ktid["head"] = copy.deepcopy(ip[2])
    protos_ktid["body"] = copy.deepcopy(ip[3])
    protos_ktid["blank"] = copy.deepcopy(ip[1])
    protos_plan = {"line": copy.deepcopy(ps(c_kt_plan)[1])}

    set_title_cell(c_title_book, TITLE)
    set_title_cell(c_title_kt, TITLE)
    rewrite(c_bg, protos_bg, BOOK_BG, keep_label=False)
    rewrite(c_task, protos_task, BOOK_TASK, keep_label=False)
    rewrite(c_kt_bg, protos_ktbg, KT_BG, keep_label=True)
    rewrite(c_kt_idea, protos_ktid, KT_IDEA, keep_label=True)
    rewrite(c_kt_plan, protos_plan, KT_PLAN, keep_label=True)
    doc.save(str(OUT))

    print("saved:", OUT.name, "bytes=", len(OUT.read_bytes()))
    for tag, blocks, need in (("任务书·背景及目的", BOOK_BG, "≈300"), ("任务书·任务及要求", BOOK_TASK, "≈250"),
                              ("开题·背景与意义", KT_BG, "≥300"), ("开题·思路与方法", KT_IDEA, "≥300")):
        text = "".join(t for _, t in blocks)
        print("  %-16s 中文字数=%4d 要求%s" % (tag, cjk(text), need))


if __name__ == "__main__":
    main()
