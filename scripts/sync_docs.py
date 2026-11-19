#!/usr/bin/env python3
"""第二轮审查的 rule.md 与架构文档同步（幂等；反引号锚点用 BT 拼接）。"""
import pathlib

BT = chr(96)
R = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/rule.md")
A = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/doc/方案设计/角色与数据流-架构梳理-20261119.md")

RP = [
    ("| 版本基线 |", "| 版本基线 | **具体版本号不写进正式文档**（任务书/开题不预设）；实现候选=官方快速入门验证组合："
     "FISCO BCOS 3.16.4（当前 Latest；3.17.x 为 2026-08 起安全加固版，联调二选一）+ java-sdk 3.6.0 + JDK 11（官方口径 8–14 均支持）+ Spring Boot 2.7.x；"
     "换装 JDK 零成本（用户确认），不为版本审美整体升栈；联调后锁定并在概要设计说明书登记完整清单 |"),
    ("| 摘要计算 |", "| 摘要计算 | **后端流式 SHA-256**，前端不做密码学。多文件凭证：逐文件摘要按值升序**直接拼接**（每项固定 32 字节，无边界歧义）喂入同一摘要，"
     "得**清单聚合摘要**（与文件名、上传顺序无关；内容变化/数量增减/重复次数变化在抗碰撞假设下可检测）；上链聚合摘要，库内保留逐文件摘要 |"),
    ("| 授权与核验 |", "| 授权与核验 | 授权关系入合约（verifierId 哈希、有效期、次数、撤销位）；**核验=方案 A：核验方上传自己持有的材料**，后端算摘要与链上值比对"
     "（比对平台留存副本只能证明平台侧一致，不作为主流程）；经授权查看/下载原文同走记账；一次性 requestId 判重、以交易回执与链上事件为准，"
     "**不建本地消息表/消息队列**；群内交易按共识顺序执行，天然无读改写竞态；合约写接口限定平台服务账户 |"),
    ("| 出证 |", "| 出证 | **两阶段回执**：第一阶段取原始存证交易哈希与区块高度生成最终 PDF；第二阶段对最终 PDF 直接计算 SHA-256 得**回执摘要**（不称聚合），"
     "经合约“存证编号→回执摘要”映射（receiptHash）登记；回执**不含自身登记交易信息**；在线复核=**核验方上传自己手里的回执文件**算摘要比对映射；"
     "不签名、无二维码、无离线验签；核验页只读，不消费次数 |"),
    ("| 删除权实现 |", "| 删除权实现 | 删除密文行、包裹后 DEK 与链下身份关联 → 原文无法经正常流程恢复；链上记录无法删除：源头按 PIPL 第6条最小必要压减链上可识别信息，"
     "断链下关联，技术上难以删除的部分按 47 条 2 款**停止处理**；**不写“密钥销毁=完全删除”**、不主张匿名化；"
     "**备份边界：原型不做生产级备份副本清理，论文按备份保留策略说明，不宣称消除所有副本** |"),
    ("| 原文存储 |", "| 原文存储 | **MySQL 存储密文（对齐申报表）**：逐文件 AES-256-GCM 加密后入文件表 LONGBLOB；每次加密与每次包裹独立随机 IV；"
     "主密钥部署注入、不入库不同备；**大小闭环：单文件 10MB / 单凭证 30MB，前后端同时校验，max_allowed_packet=64MB、multipart 35MB，列表查询不取 BLOB 列** |"),
]

lines = R.read_text(encoding="utf-8").splitlines()
for anchor, newline in RP:
    if newline[:14] in "\n".join(lines):
        continue
    idx = next((i for i, l in enumerate(lines) if anchor in l), None)
    if idx is None:
        print("RULE-MISS", anchor)
        continue
    lines[idx:idx + 1] = [newline]
# 法条与文献两行做子串增强
text = "\n".join(lines)
text = text.replace("《个人信息保护法》第13、14、17、28、29、45、47条", "《个人信息保护法》第6、13、14、17、28、29、30、45、47、51条")
if "袁勇" not in text:
    text = text.replace("凡无法核实的一律不写，宁缺勿造；规范文本与标准按上表引用。",
        "凡无法核实的一律不写，宁缺勿造；规范文本与标准按上表引用。已核实可用：袁勇、王飞跃《区块链技术发展现状与展望》自动化学报 2016,42(4):481-494"
        "（doi:10.16383/j.aas.2016.c160158）；杭州互联网法院司法区块链 2018-09-18 上线（法制日报）。其余候选论文阶段逐条核实后补充。")
if "5e. scripts/patch_round2.py" not in text:
    text = text.replace("  5d.", "  5d.").replace("  6. scripts/dump_mirror.py",
        "  5e. scripts/patch_round2.py 第二轮审查修订（16/17/18 归位、PIPL 精确归口、核验方案 A、回执映射与上传比对、去终端信息、版本不入正式文档；幂等）\n  6. scripts/dump_mirror.py")
R.write_text(text + "\n", encoding="utf-8")
print("rule.md ok, lines", len((text + "\n").splitlines()))

s = A.read_text(encoding="utf-8")
REP = [
    ("每项 4 字节大端长度前缀定长编码后喂入同一摘要", "按值升序直接拼接（每项固定 32 字节，无边界歧义）喂入同一摘要"),
    ("升序+定长前缀编码", "升序+定长直拼"),
    ("定长前缀编码：", "定长 32 字节直拼："),
    ("升序加长度前缀", "按值升序直接拼接"),
    ("反例用例（无长度前缀时的边界混淆）", "并说明定长 32 字节直拼无边界歧义"),
    ("5 预检通过 → 后端读取密文行、解 DEK 解密、重算逐文件与聚合摘要，与链上登记值比对",
     "5 预检通过 → 核验方上传自己持有的待核验材料（方案 A：验的是核验方手中的件，不是平台留存副本）；后端计算其逐文件与聚合摘要，与链上登记值比对"),
    ("（合约不读文件不算摘要、比对在后端、事件证明平台提交过经授权检查的核验结果）",
     "（合约不读文件不算摘要、比对在后端、事件证明平台提交过经授权检查的核验结果）；经授权后查看/下载原文同样走 recordVerify 记账"),
    ("→ Registry.query(no) 取链上聚合摘要与回执摘要两条登记", "→ query(no) 取链上聚合摘要与“编号→回执摘要”映射（receiptHash）两条登记"),
    ("→ 与回执 PDF 上印的值比对，返回 tx_hash/区块高度供自查",
     "→ 接收核验方上传的回执 PDF 计算 SHA-256 与映射值比对（只读库内副本不算复核完成），返回 tx_hash/区块高度供自查"),
    ("发 " + BT + "recordReceipt" + BT + " 上链", "发 " + BT + "recordReceipt" + BT + " 写入“编号→回执摘要”映射上链"),
    ("| 回执登记 | recordReceipt(no, receiptDigest) | 该 no 已登记；白名单 | ReceiptRecorded(no, receiptDigest) |",
     "| 回执登记 | recordReceipt(no, receiptDigest)：写 mapping[no]→receiptDigest | 该 no 已登记；白名单 | ReceiptRecorded(no, receiptDigest) |\n| 回执查询 | receiptDigestOf(no) | — | —（view，供复核比对） |"),
    ("合约**无外部调用、无价值转移**：不设重入锁", "全部写函数仅限白名单服务账户调用（防外部伪造 verifierId 记账）。合约**无外部调用、无价值转移**：不设重入锁"),
    ("上传大小上限 |", "上传大小上限（单文件 10MB/单凭证 30MB 前后端同校验；max_allowed_packet=64MB、multipart 35MB；列表不取 cipher 列） |"),
    ("| 删除 | 密文行 + DEK 行 + 链下关联在同一业务事务删除；链上不动；audit_log 记过程 |",
     "| 删除 | 密文行 + DEK 行 + 链下关联在同一业务事务删除；链上不动；audit_log 记过程 |\n| 备份 vs 删除 | 原型不做生产级备份副本清理；恢复旧备份可能带回已删记录，论文按备份保留策略说明边界 |"),
    ("系统另存提交终端信息与文件原始元数据供综合审查", "系统另存平台受理时间与文件原始元数据（名称、大小、格式）供综合审查"),
    ("对最终 PDF 计算 SHA-256 聚合，", "对最终 PDF 直接计算 SHA-256 得回执摘要，"),
    ("（编译器锁定0.8.11，在FISCO BCOS 3.x支持范围内）", "（编译器取 0.4.25–0.8.11 支持区间内的 0.8 系列，具体随联调锁定）"),
    ("编译器锁定 0.8.11", "取 0.8 系列，具体随联调锁定"),
    ("（编译器版本需装好后核实）", ""),
    ("仅改文件名或调整上传先后 → 聚合摘要**不变**（文件名不进编码）；④ 伪造",
     "仅改文件名或调整上传先后 → 聚合摘要**不变**（内容相同仅改名的替换不视为篡改）；④ 伪造"),
]
miss = 0
for old, new in REP:
    if old in s:
        s = s.replace(old, new)
    else:
        print("ARCH-MISS:", old[:30]); miss += 1
# 版本基线段（§16 末尾）与 §13 对应表更新
old16 = "版本基线（第一周最小联调后锁定，之后不升级）：JDK 17、Spring Boot 3.3.x、FISCO BCOS 3.x 最新稳定、java-sdk 3.6.x、Solidity 0.8.11、MySQL 8、MyBatis-Plus、OpenPDF、Vue 3 + Vite + Element Plus；联调不通回退官方组合 JDK 11 + Spring Boot 2.7.18 并记录于概要设计说明书。"
new16 = "版本基线（第一周最小联调后锁定，之后不升级；**只登记于本文与概要设计说明书，不写进任务书/开题**）：候选=官方验证组合 JDK 11 + java-sdk 3.6.0 + Spring Boot 2.7.18 + FISCO BCOS（3.16.4 或 3.17.x 二选一）+ Solidity 0.8 系列（≤0.8.11）+ MySQL 8 + MyBatis-Plus + OpenPDF + Vue 3/Vite/Element Plus；换装 JDK 零成本，不为版本审美升栈。"
if old16 in s:
    s = s.replace(old16, new16)
elif "候选=官方验证组合" not in s:
    print("ARCH-MISS: 16基线段"); miss += 1
s = s.replace("| **两阶段回执、不做数字签名/二维码** | 开题·出证与验证、边界 |",
              "| **两阶段回执、回执摘要直算 SHA-256、上传式复核、receiptHash 映射、不做签名** | 开题·出证与验证、边界 |")
s = s.replace("| 授权核验（发交易扣次）与回执只读复核的区分；**合约不算摘要** | 开题·难点三、出证与验证 |",
              "| 授权核验=方案A（上传待验材料、发交易扣次）与回执只读复核的区分；**合约不算摘要**；授权亦平台代发 | 开题·难点三、出证与验证 |")
A.write_text(s, encoding="utf-8")
print("arch ok, miss", miss, "len", len(s))
