---
name: project-report-template-word
description: 按用户的《项目报告模板》固定格式生成 Word (.docx) 项目报告/测试报告/技术报告。凡用户要求"写项目报告""生成 word/docx 报告""按模板格式出报告""写测试报告/技术报告/结题报告"，或要求报告格式为黑体标题、宋体+Times New Roman 正文、三线表、图表题注、图表目录、罗马/阿拉伯分节页码时，必须使用本技能——即使用户没有明说"模板"两个字。
---

# 项目报告模板 Word 生成

把报告内容填进用户的固定格式模板，产出 `.docx`。**格式来自模板底版，不要凭记忆或
其他技能的默认值排版**——官方 docx 技能的默认行距（1.3 倍）、字体方案与本模板冲突，
加载了它也不要采用它的 Formatting Standards。

## 工作流程

1. **收集内容**：向用户确认报告主题、章节结构、图表与参考文献。用户给了素材
   （md/txt/代码/数据）就直接提炼；没给就按常规报告结构起草后让用户确认。
2. **写驱动脚本**：参照下面的最小示例，新建一个 Python 脚本（放在输出目录），
   import 本技能的 `build_report`，把内容逐段填入。
3. **运行脚本**生成 docx。
4. **交付**：告知用户打开文档后目录/编号会自动刷新；若未刷新，全选按 F9
   （或右键"更新域"）。

## 环境准备

依赖 `python-docx`（首次使用先检查：`python -m pip show python-docx`，缺则
`python -m pip install python-docx`）。

## 最小可用示例

```python
import sys
sys.path.insert(0, r"D:\System\zhuomian\模板skill\.zcode\skills\project-report-template-word\scripts")
from build_report import Report, cn_num

r = Report()
r.set_cover(title="智能决策辅助系统测试报告", date="2026年 10月")  # 标题可含 \n 分两行

r.h1(f"{cn_num(1)}.测试概述")          # 一级标题编号用中文数字：一. 二. 三.
r.para("本报告对……进行全面测试。")      # 正文：自动首行缩进2字符、1.5倍行距、两端对齐
r.h2("1.1 测试目的")                   # 二级编号用"章号.序号"
r.h3("1.1.1 任务需求")
r.para_rich(["……满足任务需求", ("[1]", {"sup": True}), "。"])  # 上标引用

r.h1(f"{cn_num(2)}.测试内容")
r.table(caption="测试内容", headers=["序号", "项目", "内容"],
        rows=[["1", "功能测试", "……"]], widths_cm=[1.5, 3.5, 9.0])

r.h1(f"{cn_num(3)}.测试结果及分析")
r.figure(r"fig1.png", "各方案响应时间对比", width_cm=11.5)   # 图题注在下方
r.table(caption="构建框架", headers=["输入", "处理过程", "输出"],
        rows=[["想定条件", "t←0\nWhile … do\nend", "方案数据库"]],
        kind="threeline", cell_align="left")               # 三线式：框架/流程表

r.references(["张三. 某系统设计[J]. 系统工程与电子技术, 2024, 46(2): 512-520.",
              "GB/T 25000.51-2016, 系统与软件工程[S]. 北京: 中国标准出版社, 2016."])
r.save(r"D:\输出目录\报告.docx")
```

## 内容规则（模板约定，务必遵守）

- **标题编号写进标题文字**，模板不使用自动编号：
  - 一级：`一.` `二.` `三.`（中文数字 + 点，用 `cn_num(n)` 生成）
  - 二级：`1.1`、`3.2`（阿拉伯数字，章号 = 该章是第几个一级标题）
  - 三级：`1.1.1`
- **图表题注**自动带 SEQ 域编号，`caption` 参数只写标题文字，不要写"表 1"/"图 1"前缀。
  正文引用图表写"见表 1""见图 2"。
- **表格选型**：数据/结果表用默认 `kind="grid"`（全框线+浅灰表头）；框架、流程、
  伪代码类表用 `kind="threeline"`（上下粗线、无左右外框），此时单元格常配
  `cell_align="left"`。
- **参考文献**：GB/T 7714 顺序编码制。文末列表用 `references([...])`；正文引用点用
  `para_rich([("…",), ("[1]", {"sup": True})])` 上标标注，编号与文末一一对应。
- **无图或无表时**：在写正文前调用 `r.remove_list_of_figures()` /
  `r.remove_list_of_tables()` 删除对应的图表目录块，避免出现空标题。
- **章节较多正文较长时**不要每章手动分页——模板正文是连续排版的；只有当大图/大表
  被不合理切断时才用 `r.page_break()` 微调。

## API 速查

| 方法 | 用途 | 关键参数 |
|---|---|---|
| `set_cover(title, org, date)` | 封面标题/落款/日期 | title 可含 `\n` 分两行 |
| `h1/h2/h3(text)` | 三个级别标题 | 编号写在 text 里 |
| `para(text, indent, align)` | 正文段 | 默认缩进+两端对齐 |
| `para_rich(segments)` | 带格式正文 | 段=(文本, dict)；键：bold/italic/sup/sub/size/color/east/ascii |
| `figure(path, caption, width_cm)` | 插图+图题注 | 不给宽度则自动限宽 ≤14.5cm |
| `table(caption, headers, rows, kind, widths_cm, header_shading, header_bold, cell_align)` | 表格+表题注 | kind="grid"/"threeline" |
| `references(entries, heading)` | GB/T 7714 文末列表 | heading 默认"参考文献" |
| `page_break()` | 手动分页 | 慎用 |
| `remove_list_of_figures()` / `remove_list_of_tables()` | 删除图表目录块 | 无图/无表时 |
| `save(path)` | 保存 docx | |

多步骤长报告建议把章节内容写成数据（列表/字典）再循环调用，避免逐行硬编码。

## 格式底线（自动继承自底版，不要覆盖）

- 正文：宋体 + Times New Roman 小四（12pt）、1.5 倍行距、首行缩进 2 字符
- 标题1：黑体+Times New Roman 16pt 加粗居中；标题2：宋体 16pt 加粗；标题3：宋体 14pt 加粗
- 题注：宋体 10.5pt 居中，表题注在表上方、图题注在图下方
- 页面：A4，上下 2.54cm、左右 3.18cm；封面无页码 → 前置部分罗马页码（I、II）→ 正文阿拉伯页码从 1 起
- 目录与图表目录是 Word 域，条目和页码由 Word/WPS 打开时刷新生成

完整格式规范（含样式 ID、原始 XML 值、域代码）见
`references/format-spec.md`——只在需要排查格式异常或扩展新元素类型时阅读。

## 排查与维护

- 生成失败提示样式缺失 → `assets/template-base.docx` 损坏，重跑
  `scripts/make_template_base.py` 从 `assets/项目报告模板.docx` 重建底版。
- 需要预览效果时可参考：Word COM `doc.Fields.Update()` + `SaveAs2(FileFormat=17)`
  转 PDF（本机装有 Microsoft Word）；LibreOffice/WPS 亦可。
- 用户修改了格式要求（如换字体）→ 先改 `assets/项目报告模板.docx` 或底版样式，
  再同步更新 `references/format-spec.md`，不要在驱动脚本里逐段覆盖格式。
