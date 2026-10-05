# fde-pre-visit-intake · 上门前信息表 AI 填表助手

一个给 AI 编程助手（如 Claude Code）用的 skill：FDE（Forward Deployed Engineer，驻场工程师）上门前，让客户通过和 AI 对话，自动填好《上门前信息表》Excel。

> An agent skill that interviews a client in plain Chinese before an FDE site visit, takes a read-only snapshot of their PC (with consent), and produces a filled-in Excel intake form. Windows-first, no extra installs on the client machine.

## 为什么做这个

FDE 上门的时间很贵，应该花在只有现场才能做的事上：看员工实际怎么操作、拿原始文件、当面取得同意。软件版本、电脑配置、业务量、现在的做法（SOP），本可以提前拿到，但实际上常常拿不到：

- 客户不愿意填长表格，填了也常常是“应该怎么做”，不是“实际怎么做”；
- 客户不知道怎么查电脑配置和软件版本；
- 客户对 AI 能做什么了解有限，说不清自己的需求。

这个 skill 让客户用自己的 AI 工具，通过聊天把表填完。AI 一次问一个问题，帮客户把“我们平时这样做”拆成一步步的 SOP，经客户同意后只读查看电脑，最后生成 Excel。

## 客户这边怎么用

1. 安装（见 [docs/客户安装说明.md](docs/客户安装说明.md)）。
2. 对 AI 说：**帮我填上门前信息表**。
3. 大约 20 分钟，可以随时停，下次接着填。结果存在桌面的“上门前信息表”文件夹里。
4. 客户自己把 Excel、录屏和截图通过网盘或 U 盘发给 FDE。

## 表格包含什么

| 表 | 内容 |
| --- | --- |
| 填写说明 | 怎么填、文件怎么发、注意事项 |
| 1 基本情况 | 团队和对接人、业务量、期望、限制（15 个问题） |
| 2 需求与流程 | 最多 3 个需求；每个需求按步骤写谁做、用什么软件、多久、每天几次、最麻烦的地方 |
| 3 电脑与软件 | 电脑配置、相关软件及版本、网络和存储 |
| 4 文件清单 | 录屏、输入和产出样例、设置截图等，带提交状态 |
| 5 特殊设备 | 工控一体机、触摸屏、收银机、平板、PLC 等（没有可跳过）：型号、系统、厂商和保修、联网、接口、数据怎么进出、何时能停机；靠拍照和询问，不在设备上运行任何东西 |

## 安全和隐私原则

这些规则写在 [SKILL.md](fde-pre-visit-intake/SKILL.md) 里，AI 必须遵守：

- 不问、不记任何密码；建议账号和商品用代号。
- 查看电脑**只读**，每次查看前先说明要看什么，客户同意后才运行；不安装、不修改、不删除任何东西。
- 不读文档、聊天记录、浏览器数据；统计成片大小时只看文件大小，不看文件名。
- 不联网发送任何内容，结果只保存在本机，由客户自己决定发给谁。
- 只记录客户现在的做法，不替客户设计新流程。

## 目录结构

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

## 自己改表格

1. 修改 `tools/build_template.py` 里的问题、示例或列。
2. 运行 `python3 tools/build_template.py`，会重新生成 `template/intake_form.xlsx` 和 `template/fields.json`。
3. 如果改了字段名，同步更新 SKILL.md 里的问题表和 answers.json 格式说明。
4. 用示例检查填表脚本：

```bash
python3 fde-pre-visit-intake/scripts/fill_form.py --answers examples/answers.example.json --out /tmp/filled.xlsx
```

## 当前状态（v0.1）

- 已验证：Mac 上 `fill_form.py` 能把示例答案（含特殊字符、换行、特殊设备）填进正确的格子，并保留格式和下拉框；`build_template.py` 能复现模板和对照表。
- **未验证**：
  - `collect_env.ps1` 和 `fill_form.ps1` 从未在 Windows 真机上运行过；
  - Linux 和国产系统（统信 UOS、麒麟等）只写了查看命令，没有实测；
  - 整套对话流程还没有在真实的 AI 工具里、和真实客户走过一遍。
- 欢迎在 Issues 里反馈问题。
- 安装说明只写了 Claude Code 的目录。其他支持 skill 的 AI 工具，请按该工具的文档放置；不支持的，可以让 AI 直接读取 SKILL.md。

## 背景

这个 skill 来自一次真实的 FDE 上门实施复盘：现场时间被大量花在本可以提前获取的信息上。配套的方法论见作者的《FDE 实施方法论》。

## License

[MIT](LICENSE)
