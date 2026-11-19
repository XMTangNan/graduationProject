import pathlib, re
RULE = pathlib.Path("rule.md")
lines = RULE.read_text(encoding="utf-8").splitlines()

# 1) 合并被多行拆断的表格行（续行不以 | 开头且上一行以 | 开头）
merged = []
for ln in lines:
    if merged and merged[-1].startswith("|") and ln and not ln.startswith("|") and not ln.startswith("#") and ln.strip():
        merged[-1] = merged[-1].rstrip() + ln.strip()
    else:
        merged.append(ln)

# 2) 重写 §9 为固定 4 条
out, i = [], 0
while i < len(merged):
    if merged[i].startswith("## 9. 未核实事项"):
        out.append(merged[i])
        out.append("")
        out.extend([
            "- 环境（Ubuntu 24.04 / JDK 21 / 4 节点单机内存）用户已确认不是问题，不列为风险项；实现前仍按 §4 顺序先跑通官方 demo（建链 → 部署合约 → java-sdk 调用）再进业务。",
            "- FISCO BCOS 3.x 支持 Solidity 0.4.25–0.8.11（已核实），合约编译器锁定 0.8.x（建议 0.8.11），避免“0.8 内置溢出检查”落空。",
            "- 核验交易由平台托管账户代发、合约参数带 verifierId 的留痕方案，合约事件须同时记录 verifierId 与 msg.sender（已在开题写明）。",
            "- 单机 4 节点 + 后端 + MySQL 的内存占用仍按 §4 参数实测一次，写进论文性能小节（结论仅适用单机）。",
        ])
        i += 1
        while i < len(merged) and not merged[i].startswith("## 10."):
            i += 1
        continue
    out.append(merged[i]); i += 1

RULE.write_text(chr(10).join(out) + chr(10), encoding="utf-8")
print("lines", len(out))
