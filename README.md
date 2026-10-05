# FDE 信息预收集 skill · fde-pre-visit-intake

[中文](#中文) | [English](#english)

<a id="中文"></a>

## 中文

一个给 AI 编程助手（如 Claude Code）用的 skill：FDE（Forward Deployed Engineer，驻场工程师）上门前，客户通过和自己的 AI 对话，把《上门前信息表》Excel 填好。

### 为什么做这个

FDE 上门的时间很贵，应该花在只有现场才能做的事上：看员工实际怎么操作、拿原始文件、当面取得同意。软件版本、电脑配置、业务量、现在的做法（SOP），本可以提前拿到，但实际上常常拿不到：

- 客户不愿意填长表格，填了也常常是“应该怎么做”，不是“实际怎么做”；
- 客户不知道怎么查电脑配置和软件版本；
- 客户对 AI 能做什么了解有限，说不清自己的需求。

这个 skill 让客户用自己的 AI 工具，通过聊天把表填完。AI 一次问一个问题，帮客户把“我们平时这样做”拆成一步步的 SOP，经客户同意后只读查看电脑，最后生成 Excel。

### 客户这边怎么用

1. 安装，见 [docs/客户安装说明.md](docs/客户安装说明.md)，或者直接下载 [Releases](https://github.com/mixx993/fde-pre-visit-intake/releases) 里的 zip。
2. 对 AI 说：**帮我填上门前信息表**。
3. 大约 20 分钟，可以随时停，下次接着填；任何问题都可以说“跳过”。结果保存在桌面的“上门前信息表”文件夹里。
4. 客户自己把 Excel、录屏和截图通过网盘或 U 盘发给 FDE。

### 表格包含什么

| 表 | 内容 |
| --- | --- |
| 填写说明 | 怎么填、文件怎么发、注意事项 |
| 1 基本情况 | 团队和对接人、业务量、期望、限制（15 个问题） |
| 2 需求与流程 | 最多 3 个需求；每个需求按步骤写谁做、用什么软件、多久、每天几次、最麻烦的地方 |
| 3 电脑与软件 | 电脑配置、相关软件及版本、网络和存储 |
| 4 文件清单 | 录屏、输入和产出样例、设置截图等，带提交状态 |
| 5 特殊设备 | 工控一体机、触摸屏、收银机、平板、PLC 等（没有可跳过）：型号、系统、厂商和保修、联网、接口、数据怎么进出、何时能停机；靠拍照和询问，不在设备上运行任何东西 |

### 安全和隐私原则

这些规则写在 [SKILL.md](fde-pre-visit-intake/SKILL.md) 里，AI 必须遵守：

- 任何问题、任何一整部分都可以跳过，不需要理由；AI 不追问、不劝说，也不替客户猜答案。表里分开标记“不清楚”和“（跳过）”，FDE 能看出格子为什么是空的。
- 不问、不记任何密码；建议账号和商品用代号。
- 查看电脑**只读**，每次查看前先说明要看什么，客户同意后才运行；不安装、不修改、不删除任何东西。
- 不读文档、聊天记录、浏览器数据；统计成片大小时只看文件大小，不看文件名。
- 不联网发送任何内容，结果只保存在本机，由客户自己决定发给谁。
- 只记录客户现在的做法，不替客户设计新流程。

### 目录结构

```
fde-pre-visit-intake/          # skill 本体，整个文件夹交给客户
  SKILL.md                     # AI 的对话规则和流程
  template/intake_form.xlsx    # 空白表格模板
  template/fields.json         # 字段 → 单元格对照表
  scripts/collect_env.ps1      # Windows 只读环境检查（PowerShell 5.1，无需安装）
  scripts/fill_form.ps1        # Windows 按 answers.json 填表（无需安装）
  scripts/fill_form.py         # Mac/Linux/国产系统填表（只用 Python 标准库）
tools/build_template.py        # 修改表格后重新生成模板和 fields.json（需要 openpyxl）
examples/answers.example.json  # answers.json 示例（虚构数据）
docs/客户安装说明.md
```

### 自己改表格

1. 修改 `tools/build_template.py` 里的问题、示例或列。
2. 运行 `python3 tools/build_template.py`，会重新生成 `template/intake_form.xlsx` 和 `template/fields.json`。
3. 如果改了字段名，同步更新 SKILL.md 里的问题表和 answers.json 格式说明。
4. 用示例检查填表脚本：

```bash
python3 fde-pre-visit-intake/scripts/fill_form.py --answers examples/answers.example.json --out /tmp/filled.xlsx
```

### 当前状态（v0.1）

- 已验证：Mac 上 `fill_form.py` 能把示例答案（含特殊字符、换行、特殊设备）填进正确的格子，并保留格式和下拉框；`build_template.py` 能复现模板和对照表。
- **未验证**：
  - `collect_env.ps1` 和 `fill_form.ps1` 从未在 Windows 真机上运行过；
  - Linux 和国产系统（统信 UOS、麒麟等）只写了查看命令，没有实测；
  - 整套对话流程还没有在真实的 AI 工具里、和真实客户走过一遍。
- 安装说明只写了 Claude Code 的目录。其他支持 skill 的 AI 工具，请按该工具的文档放置；不支持的，可以让 AI 直接读取 SKILL.md。
- 欢迎在 Issues 里反馈问题。

### 背景

这个 skill 来自一次真实的 FDE 上门实施复盘：现场时间被大量花在本可以提前获取的信息上。

---

<a id="english"></a>

## English

An agent skill (for Claude Code and other tools that support the SKILL.md format) that lets a client fill in an FDE (Forward Deployed Engineer) pre-visit intake form by chatting with their own AI assistant before the engineer arrives on site.

The skill, the form and the conversation are in **Chinese**, written for small and medium businesses in China. The structure works for any language: translate `SKILL.md` and `tools/build_template.py`, then regenerate the template.

### Why

On-site time is the most expensive part of an FDE engagement. It should go to things only the site can give you: watching how staff actually work, collecting original files, getting consent face to face. Software versions, PC specs, volumes and the current workflow (SOP) could be collected in advance, but in practice they rarely are:

- Clients put off long forms, and when they fill them in they describe how work *should* be done, not how it *is* done.
- Clients don't know how to look up their PC specs or software versions.
- Clients have a limited sense of what AI can do, so their stated needs are often vague.

With this skill the client fills in the form by chatting with their own AI tool. The AI asks one question at a time, breaks "this is how we usually do it" into concrete steps, takes a read-only look at the PC once the client agrees, and writes the answers into the Excel form.

### How the client uses it

1. Install it: see [docs/客户安装说明.md](docs/客户安装说明.md) (Chinese), or download the zip from [Releases](https://github.com/mixx993/fde-pre-visit-intake/releases).
2. Tell the AI: **帮我填上门前信息表** ("help me fill in the pre-visit form").
3. About 20 minutes. The client can skip any question, stop at any point and resume later; answers are saved to a folder on the desktop.
4. The client sends the Excel file, screen recordings and screenshots to the FDE through cloud storage or a USB drive.

### What the form covers

| Sheet | Contents |
| --- | --- |
| Instructions | How to fill it in, how to send files, privacy notes |
| 1 Basics | Team and contacts, volumes, expectations, constraints (15 questions) |
| 2 Needs and workflow | Up to 3 needs; for each, step by step: who, which software, how long, how often, the most painful part |
| 3 PCs and software | PC specs, relevant software and versions, network and storage |
| 4 Files | Screen recordings, input and output samples, settings screenshots, with submission status |
| 5 Special devices | Industrial PCs, touch panels, POS terminals, tablets, PLCs and the like (optional): model, OS, vendor and warranty, network, interfaces, how data gets in and out, when it can be taken offline. Collected through photos and questions only; nothing is run on the device |

### Safety and privacy rules

These rules are written into [SKILL.md](fde-pre-visit-intake/SKILL.md) and the AI must follow them:

- Any question or whole section can be skipped without giving a reason. The AI does not press, persuade or guess the answer, and the form marks "不清楚" (don't know) and "（跳过）" (skipped) separately, so the FDE can tell why a cell is empty.
- Never ask for or record any password; suggest code names for accounts and products.
- PC checks are **read-only**. Before each check the AI explains what it will look at and runs it only after the client agrees. Nothing is installed, changed or deleted.
- No reading of documents, chat history or browser data. When sizing video files, only file sizes are read, not file names.
- Nothing is sent over the network. Results stay on the local machine, and the client decides who receives them.
- The AI records the current workflow only; it does not redesign the client's process.

### Layout

```
fde-pre-visit-intake/          # the skill itself; hand this whole folder to the client
  SKILL.md                     # conversation rules and flow for the AI
  template/intake_form.xlsx    # blank form template
  template/fields.json         # field-to-cell map
  scripts/collect_env.ps1      # Windows read-only PC snapshot (PowerShell 5.1, no installs)
  scripts/fill_form.ps1        # Windows form filler from answers.json (no installs)
  scripts/fill_form.py         # Mac/Linux form filler (Python standard library only)
tools/build_template.py        # regenerate the template and fields.json after edits (needs openpyxl)
examples/answers.example.json  # sample answers.json (fictional data)
docs/客户安装说明.md            # client install guide (Chinese)
```

### Customizing the form

1. Edit the questions, examples or columns in `tools/build_template.py`.
2. Run `python3 tools/build_template.py` to regenerate `template/intake_form.xlsx` and `template/fields.json`.
3. If you rename fields, update the question tables and the answers.json schema in SKILL.md.
4. Check the filler with the sample:

```bash
python3 fde-pre-visit-intake/scripts/fill_form.py --answers examples/answers.example.json --out /tmp/filled.xlsx
```

### Status (v0.1)

- Verified: on macOS, `fill_form.py` writes the sample answers (including special characters, line breaks and special devices) into the right cells and keeps formatting and dropdowns; `build_template.py` reproduces the template and field map.
- **Not yet verified**:
  - `collect_env.ps1` and `fill_form.ps1` have never been run on a real Windows machine;
  - Linux and Chinese domestic distributions (UOS, Kylin) only have commands listed, not tested;
  - The full conversation has not yet been run end to end in a real AI tool with a real client.
- The install guide only gives the folder for Claude Code. For other tools that support skills, follow their docs; for tools that don't, ask the AI to read `SKILL.md` directly.
- Issues and feedback are welcome.

### Background

This skill came out of a retrospective on a real FDE site visit, where much of the on-site time went to information that could have been collected beforehand.

## License

[MIT](LICENSE)
