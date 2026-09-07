# campus-career-coach · 大学生求职教练

一个面向**大学生 / 应届生**的求职全流程教练，专注于把"我不会写简历 / 不知道选什么岗 / 面试没底气"这类**不会想、没做过、没思路**的问题，通过**苏格拉底式提问**逼出你自己脑子里的答案，并配合脚本与参考资料完成从定位、简历、面试到 offer 的整个流程。

> 教练的核心职责是**让用户自己想明白**，而不是替用户想明白——但用户真的想不出来时，必须直接给答案。

## 适用场景

- 简历没东西可写、写得没底气、写完不敢投
- 不知道自己想做什么 / 适合什么 / 该不该考研考公
- 拿到 JD 不会拆，不知道怎么针对性改简历
- 模拟面试、面试复盘、群面 / 压力面 / 行为面试
- 多个 offer 比较、薪资谈判、三方协议 / 违约金
- 网申填写、性格测评、行测 / 笔试
- 职场新人：试用期、入职 90 天、向上管理

## 特性

- **苏格拉底式提问**：一次只问 1 个主问题、每轮走"问 → 镜像复述 → 戳破 → 给最小下一步"
- **兜底条款**：用户连续两轮空白、要求直说、有时限、情绪低落时，立即停止追问直接给单一方案
- **三件可执行产物**：`求职档案/素材库.md`、`求职档案/用户画像.md`、`求职档案/错题本.md` —— 跨会话持续累积
- **证据链校验**：每一条简历 bullet 都要能追问三层，避免面试露馅
- **PDF 导出 / 简历体检 / docx 生成** 全部用**免费**工具（LibreOffice / Word / WPS 内置 / PDF24）
- **零经历急救**：专门为"大学混过来的 / 什么都没学到"这类用户设计

## 包含什么

```
campus-career-coach/
├── SKILL.md                # 入口：苏格拉底铁律 + 七阶段路由 + 完整工作流
├── references/             # 11 个分领域参考
│   ├── 00-socratic-playbook.md  # 通用教练技术
│   ├── 01-self-discovery.md     # 自我探索与优势挖掘
│   ├── 02-jd-decoding.md        # JD 拆解
│   ├── 03-company-research.md   # 公司背调
│   ├── 04-resume-writing.md     # 简历写作规范
│   ├── 05-resume-review.md      # 诊断评分卡
│   ├── 06-interview-simulation.md # 模拟面试
│   ├── 07-offer-and-workplace.md  # offer 与职场
│   ├── 08-gap-closing.md        # 能力缺口补救
│   ├── 09-网申与笔试.md         # 网申 / 测评 / 笔试
│   ├── 10-evidence-chain.md     # 证据链与包装红线
│   └── 11-零经历急救.md         # 零实习 / 零校园经历急救
├── scripts/                # 5 个 Python 脚本（简历生命周期）
│   ├── read_resume.py      # 解析 .docx 为纯文本（纯标准库）
│   ├── resume_lint.py      # 简历规则体检
│   ├── build_resume_docx.py # JSON → .docx（含照片、中文字体）
│   ├── export_pdf.py       # docx → pdf（LibreOffice / Word / WPS，免费）
│   └── evidence_check.py   # 简历证据链校验
├── assets/                 # 简历数据 JSON schema 与示例
└── 求职档案/                # （使用本技能时由用户在各自工作区创建）
```

## 安装

### 方式一：作为 WorkBuddy 技能安装（推荐）

将整个 `campus-career-coach/` 文件夹复制到：

- 用户级：`~/.workbuddy/skills/campus-career-coach/`
- 项目级：`<你的项目>/.workbuddy/skills/campus-career-coach/`

重启 WorkBuddy 即可在对话中通过关键词触发：

> "帮我写简历""模拟面试""我不知道该投什么岗""我大学混过来的"……

### 方式二：作为参考材料直接阅读

把 `references/` 下的 md 文件当文章读，按需查阅。

### 方式三：用脚本

```bash
# 体检一份 .txt 简历
python scripts/resume_lint.py 你的简历.txt --jd 关键词.txt

# 把 .docx 转纯文本
python scripts/read_resume.py 你的简历.docx --outline

# 由 JSON 生成 .docx
python scripts/build_resume_docx.py --data 简历.json --out 输出.docx

# docx → pdf
python scripts/export_pdf.py 输出.docx

# 证据链校验
python scripts/evidence_check.py --data 简历.json --evidence 素材库.md --out 证据链.md
```

脚本依赖：

- `python-docx`（脚本会自动安装到当前 Python 环境）
- `export_pdf.py` 依次尝试 LibreOffice → Microsoft Word COM → WPS COM，全失败时给出免费替代方案
- 其余脚本零依赖（纯标准库）

## 快速开始

1. 复制本仓库的 `campus-career-coach/` 到 `~/.workbuddy/skills/`
2. 在 WorkBuddy 对话里说"帮我做一份简历"或"我不知道该找什么工作"
3. 教练会按你当下卡点进入 S1 ~ S7 之一，并在你的工作区创建 `求职档案/` 目录存放素材库、用户画像、错题本、证据链等

## 七阶段路由

| 阶段 | 触发信号 | 必读参考 |
|---|---|---|
| S1 定位 | 迷茫、不知道适合什么、要不要考研/考公 | `01-self-discovery.md`；自述"混过来的"→ `11-零经历急救.md` |
| S2 岗位与 JD | 有目标岗位、拿到 JD | `02-jd-decoding.md` |
| S3 公司背调 | 想了解某公司 | `03-company-research.md` |
| S4 简历 | 写 / 改简历、模板、照片 | `04-resume-writing.md` + `05-resume-review.md` + `10-evidence-chain.md`；零经历→`11-零经历急救.md` |
| S5 面试 | 模拟面试、复盘、群面 | `06-interview-simulation.md` |
| S6 offer / 职场 | offer 比较、薪资谈判、试用期 | `07-offer-and-workplace.md` |
| S7 网申 / 笔试 | 网申填写、性格测评、行测 | `09-网申与笔试.md` |

## 重要原则

- **不代写虚假经历**：所有写进简历的内容必须来自用户口述或已有文件
- **不代投 / 代填网申**：可生成内容让用户自己提交
- **不透露第三方隐私**：背调只用公开信息
- **免费优先**：推荐工具默认给免费 / 开源 / 自带方案，不引导付费
- **红线下绝不妥协**：涉及签约 / 违约 / 劳动纠纷，给框架并提示用户咨询专业人士

## 许可证

本项目基于 [MIT License](./LICENSE) 开源，可自由使用、修改、商用。

## 免责声明

本项目为**求职方法论 + 工具集**，不提供职业规划、心理咨询或法律咨询。涉及签约、违约、劳动纠纷等事项，请咨询学校就业指导中心或专业法律人士。

## 贡献

欢迎在 issue 里提场景与改进点。提交 PR 前请：

1. 阅读 `references/00-socratic-playbook.md` 保持教练语气一致
2. 新增 reference 时更新 SKILL.md 路由表
3. 新增脚本时保持"零外部依赖优先"和"免费方案"原则
