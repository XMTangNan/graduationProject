#!/usr/bin/env python3
"""按审查结论修订 rule.md（行锚点匹配，幂等）。

每项为 (锚点子串, [替换后的行...])：命中则用新行替换该行。
"""

import pathlib

RULE = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/rule.md")

EDITS = [
    ("自我指控式表述", [
        "- **文档只写“我不主张什么”，不写“我不够格什么”**：免责边界要写清；法律边界句允许写"
        "（如“不主张达到匿名化标准”“不具备公证效力”），能力自我否定句禁止（如“不满足中立第三方平台要求”）。"
    ]),
    ("已发生两次", [
        "- ⚠️ 已发生三次“大小正确、内容全 0”：WSL 内 python-docx / WPS 保存后遇 WSL 非正常关闭，"
        "ext4 日志回放导致数据块丢失（inode 保留 size、blocks=0）。校验：stat -c %b 非 0、"
        "zipfile.is_zipfile 为 True、非空字节数非 0。防护：写 .tmp → os.replace → 回读校验 → 带时间戳备份，"
        "重要产物两侧比 sha256。"
    ]),
    ("页面约 7 个", [
        "| 前端 | Vue 3 + Vite + 组件库，页面 8 个（含管理端“区块链管理”“系统设置”两个只读为主的 Tab）；"
        "下载原文后可在浏览器内重算摘要自校验（原文不再上传） |"
    ]),
    ("把综述写成带出处的引用", [
        "**禁止**：引用《民法典》第501条（与数据电文书面形式无关；申报表里有这一旧误引，本方文档不复现）；"
        "编造文献、调研数据、百分比、案例。",
        "**文献**：综述可引用真实可核的学术文献（作者、题名、期刊/会议、年份、DOI 或可访问链接），"
        "3–5 篇即可；凡无法核实的一律不写，宁缺勿造；规范文本与标准按上表引用。",
    ]),
    ("任务书·选题背景及目的 ≈300字", [
        "- 字数：任务书·选题背景及目的 350–450 字，任务及要求 450–550 字"
        "（含申报表六模块与六条过程产出，与申报表逐项对齐优先于空表字数注释）；开题两节各 ≥300 字。"
    ]),
    ("每次改完复制一份到 D 盘", [
        "- 每次改完复制一份到 D 盘同一子目录，并核对两侧 sha256 一致；注意本机 /mnt/d 是 9p **只读**挂载，"
        "脚本无法写入，需由 Windows 侧手动取走。"
    ]),
    ("改文档一律就地替换", [
        "- 改文档一律就地替换（保留 Word 元数据与手工修改），不要整篇重建；仅灾后恢复才按 §8 全链重跑。"
    ]),
    ("单机同时跑 4 节点 + 后端 + MySQL 的实际内存占用", [
        "- 单机同时跑 4 节点 + 后端 + MySQL 的实际内存占用：本机实测总内存仅 5 GB（可用约 3 GB），"
        "为当前最大风险，须先实测再写论文性能小节。",
        "- FISCO BCOS 3.x 支持 Solidity 0.4.25–0.8.11（已核实），实现时锁定编译器版本，"
        "避免“0.8 内置溢出检查”落空。",
        "- 核验交易由平台托管账户代发、合约参数带 verifierId 的留痕方案，合约事件须同时记录 "
        "verifierId 与 msg.sender（已在开题写明）。",
    ]),
    ("5b. scripts/patch_align_sb.py", [
        "  5b. scripts/patch_align_sb.py 任务书·工作任务及要求与申报表逐项对齐（功能六模块 + 过程与产出六条 + 要求；整格重写）",
        "  5c. scripts/patch_review.py 审查修订（删实名认证、聚合摘要加长度前缀、核验留痕主体、回执签名口径、开题补两模块；幂等）",
    ]),
]


def main():
    """执行全部规则修订并保存。"""
    lines = RULE.read_text(encoding="utf-8").splitlines()
    stats = {"APPLIED": 0, "MISS": 0}
    for anchor, newlines in EDITS:
        idx = next((i for i, ln in enumerate(lines) if anchor in ln), None)
        if idx is None:
            stats["MISS"] += 1
            print("MISS", anchor)
            continue
        lines[idx:idx + 1] = newlines
        stats["APPLIED"] += 1
        print("APPLIED  ", anchor)
    RULE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("saved", RULE.name, "lines", len(lines), stats)


if __name__ == "__main__":
    main()
