# 智能会话派单与运营监控 — 专利交底书

面向企业即时通信（企业微信）客户群的**智能会话派单与运营监控方法、系统、设备及存储介质**技术交底书撰写项目。

## 项目结构

```
├── references/          # 参考文档（产品手册、测试报告、创新点分析等）
├── tools/               # 工具脚本
│   └── make_native_equation_docx.py   # 生成原生公式版 Word 交底书
├── outputs/             # 产出物
│   └── 智能会话派单运营监控/
│       ├── *.md         # Markdown 版交底书（多版本迭代）
│       ├── *.docx       # Word 版交底书（多版本迭代）
│       ├── math_figures/   # 公式图表
│       └── mermaid_figures/ # 流程图
├── .claude/             # Claude Code 配置
│   └── skills/patent-disclosure-skill/  # 专利交底书生成 Skill
└── README.md
```

## 协作流程

1. **基于功能分支开发**：从 `main` 检出独立分支进行撰写/修改
   ```bash
   git checkout -b feat/your-name
   ```
2. **提交变更**：写明改了什么
3. **推送分支并创建 Pull Request**
4. **同事 Review 讨论** → 通过后合并到 `main`

### 提交信息规范

| 前缀 | 用途 |
|------|------|
| `feat:` | 新增内容（章节、权利要求等） |
| `fix:` | 修正错误（公式、表述等） |
| `review:` | 审核修改 |
| `refactor:` | 结构调整 |
| `chore:` | 杂项（配置、依赖等） |

## 快速开始

```bash
# 克隆仓库
git clone <repo-url>
cd patent-intelligent-session-dispatch

# 安装 Skill 依赖（如需使用）
pip install -r .claude/skills/patent-disclosure-skill/requirements.txt
```
