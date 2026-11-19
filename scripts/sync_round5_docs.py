#!/usr/bin/env python3
"""rule.md / 架构文档 同步第五轮口径（幂等）。"""
import pathlib

ROOT = pathlib.Path("/home/xczc/code/undergraduate/graduationProject")
R = ROOT / "rule.md"
A = ROOT / "doc/方案设计/角色与数据流-架构梳理-20261119.md"

r = R.read_text(encoding="utf-8")

# 1) 链上数据范围：把“只有摘要/编号/场景/时间”的绝对说法改准，并说明假名化仍属个人信息
old = "**链上出现的只有**：聚合摘要、随机凭证编号、场景类型、算法标识、配额与核验事件、回执摘要。"
new = ("**链上出现的只有**：聚合摘要、随机凭证编号、场景类型、算法标识、授权元数据（假名化 verifierId、有效期、配额、撤销位）、"
       "核验/查看/下载事件（requestId、actionType、时间、结果）、回执摘要。**不写姓名/手机号/身份证号/银行账号等直接身份信息**；"
       "但假名化标识若可经链下映射关联到个人，仍按个人信息要求管理，不主张其已匿名化。")
if new in r:
    pass
elif old in r:
    r = r.replace(old, new)
else:
    print("RULE-MISS: 链上边界行")

# 2) 删除权：加“合约置失效标志、撤销未结束授权”
if "失效标志" not in r:
    r = r.replace("**备份边界：原型不做生产级备份副本清理，论文按备份保留策略说明，不宣称消除所有副本** |",
                  "**备份边界：原型不做生产级备份副本清理，论文按备份保留策略说明，不宣称消除所有副本**；删除时同时在合约置凭证失效标志（evidenceDisabled），撤销或阻断该凭证全部未结束授权，此后核验/查看/下载一律拒绝 |")

# 3) 原文存储行：补敏感业务字段单独加密
if "敏感业务字段单独加密" not in r:
    r = r.replace("主密钥部署注入、不入库不同备；",
                  "主密钥部署注入、不入库不同备；**敏感业务字段单独加密**（个人身份、金融账户等合并为一个 JSON 加密列，用 KEK 直接加密，不再逐字段生成密钥），其余业务字段明文+权限控制；")

# 4) 出证行：回执 PDF 原始字节必须留存
if "PDF 原生字节留存" not in r:
    r = r.replace("不签名、无二维码、无离线验签；核验页只读，不消费次数 |",
                  "不签名、无二维码、无离线验签；核验页只读，不消费次数；**PDF 原生字节留存**（含摘要，下载直接返回该份、不重渲染；上链超时沿用同一 PDF 与摘要重试） |")

R.write_text(r, encoding="utf-8")
print("rule synced; 失效标志:", "失效标志" in r, "| 敏感业务字段:", "敏感业务字段单独加密" in r, "| PDF字节:", "PDF 原生字节留存" in r)

s = A.read_text(encoding="utf-8")

# 架构：链上边界表补授权元数据、事件、失效标志；删除流程补撤权；敏感字段加密列；回执字节
pairs = [
    ("| 授权关系、配额、核验结论 | — | 缓存/回灌 | ✓ 权威 |", "| 授权关系、配额、核验结论 | — | 缓存/回灌 | ✓ 权威（假名化 verifierId、有效期、配额、撤销位、事件） |"),
    ("**链上出现的只有**：聚合摘要、随机凭证编号、场景类型、算法标识、配额与核验事件、回执摘要。", "**链上出现的只有**：聚合摘要、随机凭证编号、场景类型、算法标识、授权元数据（假名化 verifierId/有效期/配额/撤销位）、核验与访问事件（requestId/actionType/时间/结果）、回执摘要；不写直接身份信息，假名化标识仍按个人信息管理。"),
    ("| 删除 | 密文行 + DEK 行 + 链下关联在同一业务事务删除；链上不动；audit_log 记过程 |", "| 删除 | 密文行 + DEK 行 + 链下关联 + 敏感字段密文在同一业务事务删除；同时在合约置 evidenceDisabled 并撤销未结束授权（此后核验/查看/下载全拒）；链上历史记录不动；audit_log 记过程 |"),
    ("| 授权 | grant(no, verifierIdHash, expireAt, quota) | 凭证归属；quota>0；expireAt>now | Granted |", "| 授权 | grant(no, verifierIdHash, expireAt, quota) | 凭证归属；quota>0；expireAt>now；未失效 | Granted |"),
    ("六功能模块与六项过程产出", "敏感字段加密范围（个人身份/金融账户合为一个 JSON 加密列，其余明文+权限）、回执 PDF 原始字节留存、删除置失效标志 —— 对应开题·难点二与技术路线；\n| 六功能模块与六项过程产出"),
]
m = 0
for old2, new2 in pairs:
    if old2 in s:
        s = s.replace(old2, new2)
    else:
        print("ARCH-MISS:", old2[:26]); m += 1
if "evidenceDisabled" not in s:
    s = s.replace("| 存证登记 | register(no, digest, scenario, algo) | state==NONE；清单非空；白名单 |",
                  "| 存证登记 | register(no, digest, scenario, algo) | state==NONE；清单非空；白名单 |\n| 失效 | disable(no) | 凭证归属/管理员；置 evidenceDisabled | EvidenceDisabled(no) |")
A.write_text(s, encoding="utf-8")
print("arch synced miss", m, "| evidenceDisabled:", "evidenceDisabled" in s)
