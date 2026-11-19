#!/usr/bin/env python3
"""按实测与官方文档核实结果修订技术方案表述（第 3 轮，锚点式整段改写，幂等）。

核实来源：
  - 本机 JDK 21.0.12.1 实测：SunEC 不支持 secp256k1（InvalidAlgorithmParameterException），
    支持 secp256r1；AES-GCM 同 key 同 IV 重用导致密钥流重用；AAD 不符或密文改 1 bit 均抛
    AEADBadTagException；DataOutputStream.writeInt 为大端 4 字节；流式摘要与一次性摘要一致。
  - FISCO BCOS 官方"硬件和系统要求"：内存最低 8GB/推荐 32GB、4核/8核、OS 写 Ubuntu 18.04+
    （只给下限），并明确"一台机器上不推荐配置过多节点"。
  - PBFT 节点数下限与 solc 版本上限本次未能取到官方页面，按"未核实、装好后确认"处理，不写死数字。
"""

import copy
import datetime
import pathlib
import shutil
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DOCX = pathlib.Path(
    "/home/xczc/code/undergraduate/graduationProject/doc/任务书与开题报告/"
    "2023131072_周健_SJ0205-2016版-本科毕业论文（设计）任务书与开题报告_"
    "面向大众民事的联盟链电子公证存证系统设计与实现.docx"
)

REWRITES = [
    # 基本思路：把"长度前缀"的说法改为可证的单射编码
    ("系统按“自己保全—摘要上链—本人授权—出证可复核”",
     "系统按“自己保全—摘要上链—本人授权—出证可复核”四个动作组织。用户在页面上传凭证并填写场景要素，"
     "后端以流式方式计算文件的SHA-256摘要；一份凭证包含多张截图时，按文件名排序，每个字段以4字节大端长度前缀定长编码"
     "（文件名长度‖文件名‖摘要长度‖文件摘要）后依次喂入同一摘要算法，得到一条清单聚合摘要，该编码保证清单到字节流的映射是单射、"
     "字段边界无歧义；上链登记的正是这条摘要，链下同时保留逐文件摘要。原始文件逐个用随机数据密钥加密后落盘，"
     "MySQL存放路径、文件元数据与凭证字段，主密钥单独保管。核验方凭用户签发的授权发起核验，核验本身作为一笔链上交易留痕；"
     "比对结果连同回执摘要回写链上，使回执也可以复核。"),
    # 重点：同上，删去未经证实的"避免不同组合产生同一摘要输入"
    ("场景字段模板是业务侧重点",
     "场景字段模板是业务侧重点。借贷记借款合意、交付事实、金额与还款约定；租赁记租期、押金与交割清单；"
     "简易民事协议记约定事项、参与方与履行状态；原创内容版权记作品摘要、创作完成时间与权属声明。"
     "字段既要能对应要件证明，又不能把敏感明文写进链上。一份凭证可能由多张截图组成，链上登记其按序聚合的清单摘要、"
     "链下保留逐文件摘要，核验时可以定位到被替换的具体某一张；聚合输入采用带长度前缀的定长编码，"
     "使不同清单不会因字段边界含糊而落到同一字节序列。"),
    # 区块链层：节点是 C++ 进程，无 JVM 堆；引用官方单机节点数提示与最低内存
    ("区块链层：FISCO BCOS 3.x",
     "区块链层：FISCO BCOS 3.x，在同一台主机上部署4个PBFT共识节点以模拟4个联盟成员，节点证书与配置由官方建链脚本一次生成，"
     "节点之间经本机端口通信；合约部署与登记地址由管理员白名单控制。官方对硬件的要求是内存最低8 GB、推荐32 GB，"
     "并提示同一台机器上不推荐配置过多节点，故本课题把节点数固定在4并留出资源给后端与数据库；"
     "fisco-bcos 节点是 C++ 进程、不存在 JVM 堆，如需限制其内存用 systemd 的 MemoryMax 或 cgroup v2 实现。"
     "合约用Solidity编写并部署在EVM兼容链上，Java只作为java-sdk侧的业务调用语言。"
     "摘要算法固定使用SHA-256，合约保留算法标识字段以便后续接入SM3（GB/T 32905-2016）等国密算法，本课题不实现双算法切换。"
     "跨主机、跨组织的节点分布与真实成员接入属于部署扩展，不在本科原型范围内。"),
    # 出证：secp256k1 在 JDK 默认 provider 不可用，改用原生 secp256r1；补离线验签
    ("出证与验证：回执",
     "出证与验证：回执报告由后端模板生成，内容含存证编号、链上交易哈希与区块高度、逐文件摘要与清单聚合摘要、"
     "场景字段、时间与免责说明；回执内容本身的摘要作为一条独立存证记录上链，核验页凭回执上的存证编号比对链上登记值是否一致。"
     "回执PDF用ECDSA-secp256k1对回执文件的SHA-256签名：实测本机JDK 21的默认provider（SunEC）拒绝secp256k1"
     "（InvalidAlgorithmParameterException），故签名走BouncyCastle提供的曲线实现；选用该曲线是为了与链上账户同一套密钥体系，"
     "如希望少一个依赖，可退化为JDK原生的secp256r1（P-256），两者在系统中互不影响。系统不引入CA/PKI证书链，采用自签名公钥分发，"
     "公钥与指纹在核验页公布并登记上链；"
     "前端不单独提供验签页面，验签以离线方式给出说明（用 openssl 对回执文件的SHA-256做ECDSA验证，并与链上登记的公钥指纹比对），"
     "可复核性主要来自链上摘要与交易哈希；页面部署在本机或局域网环境，答辩以现场演示为准。"),
    # 测试：把未经核实的编译器上限改为"装好后核实并锁定"，并把密钥复用/挪用的实测结论写进用例
    ("测试：合约层在本地链上",
     "测试：合约层在本地链上用脚本化用例覆盖越权登记与越权核验、重复存证、重复核验请求、授权过期、次数用尽与撤销，"
     "并发核验下断言配额不超卖；整数溢出依赖Solidity 0.8的内置检查（编译器版本在安装后以所用FISCO BCOS小版本实际支持的"
     "solc列表为准并锁定），重入由自实现的nonReentrant修饰符防范（不依赖第三方库版本适配），两者各配一条回归用例。"
     "后端密码学部分以单元测试固化实测结论：AES-GCM在同一密钥下重复使用IV会导致密钥流重用（断言必须每次生成新IV）、"
     "AAD与凭证编号和文件序号绑定后，密文被挪用到别的凭证或被改动1比特均解密失败、流式计算与一次性计算的摘要一致。"
     "系统层用JUnit与接口测试覆盖功能与并发核验，并验证重复提交与并发扣次的正确性。"
     "另设一组篡改对照实验：改动原文一个字节后验证比对必然失败，并检验多文件清单中替换单张截图能否被定位。"),
    # 范围与边界：补 verifierId 的语义限定
    ("不自研共识机制与密码算法",
     "不自研共识机制与密码算法，不涉及代币与金融业务，不宣称具备法定公证效力；"
     "链上核验留痕只证明该verifierId在该时刻被使用，不主张该次访问由核验人本人发起，该对应关系由平台账号体系担保；"
     "全部开发与测试在单台计算机上完成，系统的性能与可用性结论仅适用于该单机环境；"
     "不做前端验签页面与白名单管理界面；跨主机部署、真实第三方节点接入与国密算法适配列为后续工作。"),
    # 版本基线：去掉 BouncyCastle 依赖，加入 libssl/OpenSSL 需实测的口径
    ("技术选型版本在概要设计说明书中固化",
     "技术选型版本在概要设计说明书中固化（JDK、Spring Boot、FISCO BCOS、Solidity以安装后核实的可用版本为准、OpenPDF、BouncyCastle、ZXing），"
     "后续论文与中期报告直接引用该版本基线；操作系统按官方Ubuntu 18.04及以上口径选用，"
     "但FISCO BCOS二进制对具体libssl/OpenSSL版本的依赖需在装好后实测确认，不一致时改用官方对应构建或降级系统版本。"),
]


def first_text_run(p_el):
    """返回段落中优先含 w:t 的 run；空段落退化为第一个 run。"""
    runs = p_el.findall(qn("w:r"))
    for r in runs:
        if r.find(qn("w:t")) is not None:
            return r
    return runs[0] if runs else None


def set_paragraph_text(paragraph, text):
    """整段替换为 text，保留首个 run 的字符格式。"""
    runs = list(paragraph.runs)
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def main():
    """按锚点整段改写，幂等。"""
    dry = "--dry-run" in sys.argv
    doc = Document(str(DOCX))
    paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            keep, seen = [], set()
            for cell in row.cells:
                keep.append(cell)
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                paras.extend(cell.paragraphs)
    applied, skipped, missed = [], [], []
    for anchor, new in REWRITES:
        target = None
        for p in paras:
            if anchor in p.text:
                target = p
                break
        if target is None:
            missed.append(anchor[:18])
            continue
        if target.text.strip() == new.strip():
            skipped.append(anchor[:18])
            continue
        set_paragraph_text(target, new)
        applied.append(anchor[:18])
    for x in applied:
        print("  OK  改写", x)
    for x in skipped:
        print("  SKIP 已达标", x)
    for x in missed:
        print("  MISS 未命中", x)
    if dry or not applied:
        print("未写入")
        return
    doc.core_properties.modified = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)
    doc.save(str(DOCX))
    print("已就地保存（不生成备份）bytes=", len(DOCX.read_bytes()))


if __name__ == "__main__":
    main()
