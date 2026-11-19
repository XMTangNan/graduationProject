#!/usr/bin/env python3
"""rule.md 最后一轮：环境项降级 + 版本基线与合约防护决策（行锚点，幂等）。"""

import pathlib

RULE = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/rule.md")

EDITS = [
    ("- FISCO BCOS 3.x 对 Ubuntu 24.04 的支持与 libssl 依赖", (
        ["- 环境（Ubuntu 24.04 / JDK 21 / 4 节点单机内存）用户已确认不是问题，不再列为风险项；",
         "  实现前仍按 §4 顺序先跑通官方 demo（建链 → 部署合约 → java-sdk 调用）再进业务。",
         "- java-sdk 与 JDK 21 的兼容性：若报错优先切 JDK 17（环境问题已确认可解，不占风险预算）。"],
        ["- java-sdk 与 JDK 21 的兼容性"],
    )),
    ("- 单机同时跑 4 节点 + 后端 + MySQL 的实际内存占用", (
        ["- FISCO BCOS 3.x 支持 Solidity 0.4.25–0.8.11（已核实），合约编译器锁定 0.8.x（建议 0.8.11），",
         "  避免“0.8 内置溢出检查”落空。"],
        ["- FISCO BCOS 3.x 支持 Solidity 0.4.25–0.8.11"],
    )),
    ("| 合约 | Solidity（EVM 兼容）", (
        ["| 合约 | Solidity（EVM 兼容，编译器 0.8.x ≤0.8.11），Java 仅作 java-sdk 调用侧；两个合约：存证登记 Registry、",
         "  授权与核验留痕 Grant；重入防护用**自实现的 nonReentrant**（不依赖 OpenZeppelin，避开 OZ 5.x 的 transient storage",
         "  与链上 EVM 不兼容的风险） |"],
        ["自实现的 nonReentrant"],
    )),
    ("| 算法范围 | 只做 SHA-256 + ECDSA", (
        ["| 算法范围 | 只做 SHA-256 + ECDSA（secp256k1，签名库用 BouncyCastle——JDK 原生不支持该曲线）；",
         "  合约留算法标识字段，国密（SM3/SM2）适配列为后续工作，不实现双算法切换 |",
         "| 版本基线 | JDK 21（备 17）、Spring Boot 3.3.x、FISCO BCOS 3.x、Solidity ≤0.8.11、OpenPDF、BouncyCastle、ZXing；",
         "  在概要设计说明书中固化，论文与中期报告直接引用 |"],
        ["BouncyCastle"],
    )),
]


def main():
    """执行替换并保存。"""
    lines = RULE.read_text(encoding="utf-8").splitlines()
    applied = miss = 0
    for anchor, (newlines, guards) in EDITS:
        if any(g in ln for ln in lines for g in guards):
            miss += 1
            continue
        idx = next((i for i, ln in enumerate(lines) if anchor in ln), None)
        if idx is None:
            miss += 1
            print("MISS", anchor)
            continue
        end = idx + 1
        if anchor.startswith("- java-sdk"):
            end = idx + 1
        lines[idx:end] = newlines
        applied += 1
    # 合并被替换掉的两条环境风险行（原三行 -> 一行），删除残留的旧 java-sdk 行
    cleaned = [ln for ln in lines if not ln.startswith("- java-sdk 与 JDK 21 的兼容性（常见为 8/11/17）")]
    RULE.write_text(chr(10).join(cleaned) + chr(10), encoding="utf-8")
    print("applied", applied, "skip", miss, "lines", len(cleaned))


if __name__ == "__main__":
    main()
