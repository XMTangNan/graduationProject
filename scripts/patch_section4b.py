#!/usr/bin/env python3
"""补齐第 4 节剩余的 4.2 段：原"拒绝分支一览"改为"路径 A 的失败分支"，并清掉文档里的反引号转义残留。

用正则从"### 4.2"起始处一直替换到"### 4.3"之前，避免逐字匹配因代码反引号而失配。
幂等：新内容特征串存在即跳过。
"""

import pathlib
import re
import sys

MD = pathlib.Path("/home/xczc/code/undergraduate/graduationProject/doc/方案设计/角色与数据流-架构梳理-20261119.md")

NEW_42 = """### 4.2 路径 A 的失败分支（是 4.1 的分支，不是第三条路）

表中"步"对应 4.1 的编号；"判定层"说明谁说了算——**后端预检只负责"不值得花钱就别发交易"，权限的最终判定永远在合约**。

| 步 | 情形 | 判定层 | 行为 | 他看到 | 扣配额 |
|---|---|---|---|---|---|
| 4 | 从未授权 / 已撤销 / 已过期 | 后端预检 | 直接 403，**不发交易** | "未被授权"或"授权已失效" | 否（省 gas） |
| 6 | 预检通过但链上状态已变（如刚刚被撤销） | **合约 require** | 交易 revert，不落任何记录 | "授权已失效"（以链为准） | 否（revert 不扣） |
| 6 | 配额用尽 | 合约 require：used < quota 不成立 | 交易 revert | "配额已用完" | 否 |
| 3/6 | 同一 requestId 重发（超时重试、手抖点两次） | 合约 consumed[requestId] | 幂等：返回上次结果，不重复扣 | "本次为重复请求" | 不重复扣 |
| 5→6 | 后端重算摘要与链上登记值不符 | 后端计算，结论随交易上链 | 交易**成功**，result=MISMATCH | "内容被改动过"，可定位到具体 file_id | **是** |

三条要点（答辩就按这个说）：

1. **预检不是判定**。数据库里的授权状态可以被平台单方面改写，链上的不能；所以预检只做快速失败，真判定看合约 revert。
2. **两者不一致时以链为准**。并发撤销会造成"预检说行、合约说不行"，这是正常路径，按 revert 处理并如实记录。
3. **MISMATCH 必须扣次并留痕**。因为访问确实发生了——用户的权利是"知道谁在什么时候查过我的凭证"，哪怕结论是不一致；把它记成"失败所以不扣次"反而让越权试探变得免费。

"""


def main():
    """替换 4.2 段并清理转义残留。"""
    dry = "--dry-run" in sys.argv
    txt = MD.read_text(encoding="utf8")
    before = txt
    esc = txt.count(chr(92) + chr(96))
    txt = txt.replace(chr(92) + chr(96), chr(96))
    if "4.2 路径 A 的失败分支" in txt:
        print("  SKIP 4.2 已是新内容")
    else:
        pat = re.compile(r"### 4\.2 .*?(?=### 4\.3)", re.S)
        if not pat.search(txt):
            print("  MISS 未找到 4.2 段")
        else:
            txt = pat.sub(NEW_42, txt, count=1)
            print("  OK   4.2 段已替换")
    if esc:
        print("  FIX  清除转义反引号 %d 处" % esc)
    if not dry and txt != before:
        MD.write_text(txt, encoding="utf8")
        print("已就地保存（不生成备份）行数=", txt.count(chr(10)) + 1)
    elif dry:
        print("dry-run 未写入")
    else:
        print("无变化")


if __name__ == "__main__":
    main()
