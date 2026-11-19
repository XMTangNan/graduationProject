import pathlib

p = pathlib.Path("rule.md")
lines = p.read_text(encoding="utf-8").splitlines()
anchor = "题目来源：社会（行业）实践；题目类型：设计开发。"
block = [
    "",
    "### 学生背景与讲解口径（决定怎么讲、怎么做）",
    "",
    "- **Java / Spring 熟练**：Java 21 + Spring Boot 是主战场。所有技术讲解、架构类比、代码示例默认以 Spring 体系为参照"
    "（MVC 分层、Filter/Interceptor、事务与 @Transactional、Bean 生命周期、JdbcTemplate/JPA、MockMvc 测试），"
    "不拿 Python / Node 生态并行展开；区块链概念也用 Spring 术语类比（合约≈带状态的服务端组件，交易≈一次带签名与幂等控制的请求）。",
    "- **区块链只有书本知识、无实操**：共识、Merkle 树、账户模型、智能合约这些概念知道“是什么”，"
    "但没部署过节点、没写过合约、没调过 SDK。凡涉及链上操作一律按“第一次做”讲：先给一条能跑通的最小验证"
    "（建链 → 部署一个 Hello 合约 → java-sdk 调用），再讲原理与坑；不假设已有的运维经验。",
    "- **由此定的实施顺序**：编码第一阶段（约 1 周）只做官方 demo 打通，跑通 4 节点建链、合约部署、SDK 调用之后再进业务开发；"
    "卡住时先给可复现的最小命令与报错，不空转重试。",
]
if any("学生背景与讲解口径" in ln for ln in lines):
    print("SKIP already present")
else:
    idx = next((i for i, l in enumerate(lines) if anchor in l), None)
    if idx is None:
        print("MISS anchor")
    else:
        lines[idx + 1:idx + 1] = block
        p.write_text(chr(10).join(lines) + chr(10), encoding="utf-8")
        print("inserted after line", idx + 1, "total", len(lines))
