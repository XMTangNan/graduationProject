#!/usr/bin/env python3
"""rule.md / 架构文档同步第四轮口径（幂等）。"""
import pathlib

ROOT = pathlib.Path("/home/xczc/code/undergraduate/graduationProject")
R = ROOT / "rule.md"
A = ROOT / "doc/方案设计/角色与数据流-架构梳理-20261119.md"

r = R.read_text(encoding="utf-8")
if "先验链再差比" not in r:
    r = r.replace("上链聚合摘要，库内保留逐文件摘要；空清单拒绝登记；",
                  "上链聚合摘要，库内保留逐文件摘要；空清单拒绝登记；**差比前先验链**：以库内逐文件摘要重算聚合与链上比对确认清单可信，再做多重集合差异比较（勿称集合/差集）；")
if "回执摘要仅登记一次" not in r:
    r = r.replace("回执**不含自身登记交易信息**；在线复核=**核验方上传自己手里的回执文件**算摘要比对映射；",
                  "回执**不含自身登记交易信息**，每编号**回执摘要仅登记一次**（require receiptHash[no]==0）；在线复核=**核验方上传自己手里的回执文件**算摘要比对映射，第二笔交易的哈希与高度不对外展示（方案 B）；")
if "requestId（授权操作唯一请求编号" not in r:
    r = r.replace("一次性 requestId 判重", "requestId（授权操作唯一请求编号，网络重试沿用原号）判重")
if "访问留痕证明范围" not in r:
    r = r.replace("三类共耗“授权使用次数”，无效授权与交易未成立不扣次 |",
                  "三类共耗“授权使用次数”，无效授权与交易未成立不扣次；**访问留痕证明“授权检查通过并记账”，不证明实际阅读/完整接收** |")
if "服务端密钥托管" not in r:
    r = r.replace("- 不主张《电子签名法》", "- **密钥模式如实声明**：服务端密钥托管（KEK 平台管理、用户不持钥、平台技术上可解密业务文件），非端到端加密；文档与系统都如实写明。\n- 不主张《电子签名法》")
R.write_text(r, encoding="utf-8")
print("rule synced")

s = A.read_text(encoding="utf-8")
pairs = [
    ("3 POST /api/verify {no, verifierId, requestId=一次性令牌}", "3 POST /api/verify {no, verifierId, requestId}   // requestId=授权操作唯一请求编号，网络重试沿用原号"),
    ("（合约不读文件、不算摘要：比对发生在 Java 后端；链上事件证明的是\"平台提交了一次经授权检查的核验结果\"）",
     "（合约不读文件、不算摘要：比对发生在 Java 后端；链上事件证明\"平台提交了一次经授权检查的记账\"，不证明用户实际阅读/完整接收文件）"),
    ("5 预检通过 → 核验方上传自己持有的待核验材料（方案 A：验的是核验方手中的件，不是平台留存副本）；后端计算其逐文件与聚合摘要，与链上登记值比对",
     "5 预检通过 → 先验链再差比：以 MySQL 逐文件摘要重算清单聚合并与链上登记值比对（确认库内摘要清单可信），再与核验方上传材料的摘要做多重集合差异比较，定位新增/缺失/内容不同（方案 A：验的是核验方手中的件）"),
    ("| 回执登记 | recordReceipt(no, receiptDigest)：写 mapping[no]→receiptDigest | 该 no 已登记；白名单 |",
     "| 回执登记 | recordReceipt(no, receiptDigest)：写 mapping[no]→receiptDigest | 该 no 已登记；receiptHash[no]==0 仅一次；白名单 |"),
    ("- 用户：login、POST /evidences(multipart)", "- 密钥模式声明：服务端密钥托管（KEK 平台管理），非端到端加密；平台技术上具备解密能力，作为信任前提如实写入论文"),
]
m = 0
for old, new in pairs:
    if old in s:
        s = s.replace(old, new)
    else:
        print("ARCH-MISS:", old[:26]); m += 1
A.write_text(s, encoding="utf-8")
print("arch synced miss", m)
