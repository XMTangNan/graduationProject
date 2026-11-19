#!/usr/bin/env python3
"""rule.md 同步“最简单”口径（行锚点，幂等）。"""

import pathlib

RULE = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/rule.md")

EDITS = [
    ("| 前端 | Vue 3 + Vite + 组件库，页面 8 个", [
        "| 前端 | Vue 3 + Vite + 组件库，页面 7 个（管理端“区块链管理”“系统设置”合并为一页两个只读 Tab）；"
        "下载原文后可在浏览器内重算摘要自校验（原文不再上传） |"]),
    ("| 工作量口径 | 合约 2 个、接口约 22–24 个", [
        "| 工作量口径 | 合约 2 个、接口 18–20 个、页面 7 个（管理端两 Tab 合并为一页）、"
        "测试三类（合约权限与状态、接口与并发、篡改对照实验） |"]),
    ("| 链 | FISCO BCOS 3.x", [
        "| 链 | FISCO BCOS 3.x，**同机 4 个 PBFT 节点模拟 4 个联盟成员**（官方建链脚本一次生成证书与配置）；"
        "每节点 JVM 限 -Xmx512m/-Xms512m、小出块上限；合约部署与登记地址白名单 |"]),
    ("| 出证 | 回执打印存证编号", [
        "| 出证 | 回执打印存证编号、交易哈希、区块高度、逐文件与聚合摘要、时间与免责；回执内容摘要独立上链；"
        "回执 PDF 由平台私钥签名、公钥与指纹在核验页公布；**前端不做验签页**；核验页比对链上登记值 |"]),
    ("| 申报表新增的两模块按最小实现 |", [
        "| 申报表新增的两模块按最小实现 | 区块链管理＝2 个只读接口（当前区块高度、最近 N 条存证交易与状态）+ 1 个表格 Tab；"
        "系统设置＝场景字典用枚举、页面展示并修改上传上限/算法标识/合约地址；白名单只在合约层配置，不开发界面。"
        "管理端两 Tab 合并为同一页面。 |"]),
    ("本机实测总内存仅 5 GB", [
        "- 单机同时跑 4 节点 + 后端 + MySQL 的实际内存占用：WSL 限 6 GB 内存 + 8 GB swap、12 核；"
        "每节点限 512 MB 堆后预计约占 3.5 GB，须实测后写论文性能小节。"]),
    ("## 8. 产物与流水线", [
        "## 8. 产物与流水线",
        "",
        "- **交付物（对齐申报表“预期成果”与任务书“过程与产出”）**：① 软件一套，含完整源代码与单机部署脚本；"
        "② 毕业论文；③ 答辩PPT。三者缺一不可。"]),
]


def main():
    """执行替换并保存。"""
    lines = RULE.read_text(encoding="utf-8").splitlines()
    applied = miss = 0
    for anchor, newlines in EDITS:
        idx = next((i for i, ln in enumerate(lines) if anchor in ln), None)
        if idx is not None and len(newlines) == 1 and lines[idx] == newlines[0]:
            idx = None
        if idx is not None and len(newlines) > 1:
            ahead = [x for x in lines[idx + 1: idx + 1 + len(newlines)] if x is not None]
            if ahead[:1] == newlines[1:2]:
                idx = None
        if idx is None:
            miss += 1
            print("MISS", anchor)
            continue
        lines[idx:idx + 1] = newlines
        applied += 1
    RULE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("rule.md applied", applied, "miss", miss, "lines", len(lines))


if __name__ == "__main__":
    main()
