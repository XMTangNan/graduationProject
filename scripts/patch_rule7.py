#!/usr/bin/env python3
"""rule.md 同步 GPT 审查采纳项（整行锚点，幂等）。"""

import pathlib

RULE = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/rule.md")

LINES = [
    ("| 链 | FISCO BCOS 3.x",
     "| 链 | FISCO BCOS 3.x，**同机 4 个 PBFT 节点模拟 4 个联盟成员**（N=3f+1 为最小容错配置，与官方单机快速入门一致）；"
     "节点为 C++ 进程，不做 JVM 限制类设计，单机内存占用实测后写入论文；合约部署与登记地址白名单 |"),
    ("| 摘要计算 |",
     "| 摘要计算 | **后端流式 SHA-256**，前端不做密码学。多文件凭证：逐文件摘要按值升序、每项 4 字节大端长度前缀定长编码后喂入同一摘要，"
     "得**清单聚合摘要**（与文件名、上传顺序无关；重命名/换序不变，替换/删减/重复在抗碰撞假设下会变）；上链聚合摘要，库内保留逐文件摘要 |"),
    ("| 原文存储 |",
     "| 原文存储 | **MySQL 存储密文（对齐申报表“MySQL 存储原文资料”）**：逐文件 AES-256-GCM 加密后入文件表 LONGBLOB；"
     "数据密钥经主密钥包裹后与 MIME、单文件摘要同库；每次加密与每次包裹均用独立随机 IV；主密钥部署注入、不入库、不与业务数据同备 |"),
    ("| 删除权实现 |",
     "| 删除权实现 | 删除密文行、包裹后 DEK 与链下身份关联 → 原文无法经正常流程恢复；链上记录无法删除，按 PIPL 47 条 2 款"
     "以链上最小化+删链下关联+停止处理应对；**不写“密钥销毁=完全删除”、不主张匿名化** |"),
    ("| 合约 |",
     "| 合约 | Solidity（EVM 兼容，编译器**锁定 0.8.11**），Java 仅作 java-sdk 调用侧；**单个合约**集中四类功能区："
     "存证登记、授权签发/撤销、核验扣次与留痕、回执摘要登记；白名单与配额扣减在合约内原子完成；"
     "合约无外部调用与价值转移，**不实现重入锁** |"),
    ("| 授权与核验 |",
     "| 授权与核验 | 授权关系入合约（verifierId 哈希、有效期、次数、撤销位）；核验动作发交易留痕并扣次；一次性 requestId 判重、"
     "以交易回执与链上事件为准对账，**不建本地消息表/消息队列**；群内交易按共识顺序执行，天然无读改写竞态 |"),
    ("| 出证 |",
     "| 出证 | **两阶段回执**：第一阶段取原始存证交易哈希与区块高度生成最终 PDF；第二阶段把 PDF 的 SHA-256 聚合登记上链；"
     "回执**不含自身登记交易信息**（防递归包含），该信息由核验页查询展示；回执不做数字签名、无二维码、无离线验签，"
     "不引入 BouncyCastle/ZXing；核验页只读比对、不消费次数 |"),
    ("| 前端 |",
     "| 前端 | Vue 3 + Vite + Element Plus，页面 7 个（管理端两 Tab 合并一页，只读为主；仅上传上限可改，主密钥不经页面配置）；"
     "浏览器 WebCrypto 自查仅列后续扩展，不写进承诺与验收 |"),
    ("| 后端 |",
     "| 后端 | Spring Boot 单体，controller/service/repository/dto 分层，RESTful + JWT + MyBatis-Plus；不引入微服务、Redis、消息队列 |"),
    ("| 算法范围 |",
     "| 算法范围 | 摘要仅 SHA-256；文件加密 AES-256-GCM（JDK 自带）；链上交易签名由 java-sdk 完成，不新增密码库；"
     "合约留算法标识字段，国密（SM3/SM2）列为后续工作 |"),
    ("| 版本基线 |",
     "| 版本基线 | 首选 JDK 17 + Spring Boot 3.3.x + java-sdk 3.6.x + MySQL 8 + MyBatis-Plus + OpenPDF + Solidity 0.8.11；"
     "联调不通回退官方组合 JDK 11 + Spring Boot 2.7.18；第一周最小联调（建链→部署合约→SDK 读写）定版并锁定，之后不升级 |"),
    ("| 工作量口径 |",
     "| 工作量口径 | 合约 1 个、接口 18–20 个、页面 7 个（管理端两 Tab 合并为一页）、测试四组"
     "（功能、密码学单元、链上状态、篡改对照） |"),
    ("- 不做回执数字签名/验签、不做二维码",
     ["- 不主张《电子签名法》第13、14条的“可靠电子签名”，不引入证书体系（PKI）。",
      "- 不做回执数字签名/验签、不做二维码；演示一律用脱敏或合成材料，不收集真实用户个人信息。"]),
    ("《个人信息保护法》第13、45、47条",
     None),
    ("《人民法院在线诉讼规则》（法释〔2021〕12号）第16条",
     None),
    ("- 字数：任务书·选题背景及目的 350–450 字，任务及要求 450–550 字",
     "- 字数：任务书·选题背景及目的 350–450 字，任务及要求 500–700 字（含存储、个人信息保护与测试口径的合规细节）。"),
    ("  6. scripts/dump_mirror.py",
     ["  5d. scripts/patch_gpt.py GPT 审查修订（密文入 MySQL、法条补全收紧、去签名/重入锁/消息表、合约合一、两阶段回执、WebCrypto 降为后续；幂等）",
      "  6. scripts/dump_mirror.py 刷新正文文本镜像"]),
]


def main():
    """执行整行替换。"""
    lines = RULE.read_text(encoding="utf-8").splitlines()
    out, done = [], set()
    for ln in lines:
        hit = None
        for i, (anchor, new) in enumerate(LINES):
            if new is None or i in done:
                continue
            if anchor in ln:
                hit = i
                break
        if hit is not None:
            nv = LINES[hit][1]
            if isinstance(nv, list) and nv[-1] in lines:
                done.add(hit)
                out.append(ln)
                continue
            done.add(hit)
            new = LINES[hit][1]
            out.extend(new if isinstance(new, list) else [new])
        else:
            out.append(ln)
    # 法条清单两处小扩写（子串替换）
    text = "\n".join(out)
    text = text.replace("《个人信息保护法》第13、45、47条", "《个人信息保护法》第13、14、17、28、29、45、47条")
    text = text.replace("（法释〔2021〕12号）第16条；", "（法释〔2021〕12号）第16、17条；")
    RULE.write_text(text + "\n", encoding="utf-8")
    print("replaced", len(done), "lines; total", len((text + "\n").splitlines()))


if __name__ == "__main__":
    main()

