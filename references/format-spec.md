# 《项目报告模板》完整格式规范

来源：对 `assets/项目报告模板.docx` 的 OOXML 逐项解析（styles.xml / document.xml /
theme1.xml / settings.xml）。数值保留原始单位：字号为半磅（half-points），
间距为缇（twip，1/20 磅；567 缇 = 1cm）。styleId 为 styles.xml 中的 `w:styleId`。

## 1. 页面与分节

A4 纵向（11906×16838 缇），单栏。全文 3 个分节，页边距完全相同：
上下 1440（2.54cm），左右 1800（3.18cm），页眉 851（1.5cm），页脚 992（1.75cm），
正文可用宽度 8306 缇 ≈ 14.65cm。`docGrid type="lines" linePitch="326"`（行网格）。

| 节 | 内容 | 页码 |
|---|---|---|
| 1 | 封面（无边框表格） | `titlePg` + 无 first 页脚 → 首页无页码 |
| 2 | 目录 + 图表目录 | `pgNumType fmt="upperRoman" start="1"` → I、II… |
| 3 | 正文（正文级 sectPr） | `pgNumType start="1"`（十进制）→ 从 1 重新开始 |

页脚：`footer1.xml` 居中 PAGE 域（9pt，样式 id `a5`），节 1 default 引用，节 2/3
继承；`footer2.xml` 为空页脚。无页眉。分节属性 XML 原样保留在底版中，不要改动。

settings.xml 关键项：`defaultTabStop 420`、`characterSpacingControl
compressPunctuation`（中文压缩标点）、`compatibilityMode 14`，
以及生成器注入的 `<w:updateFields w:val="true"/>`（打开时自动刷新域）。

## 2. 默认字体（docDefaults + 主题）

- docDefaults：西文/中文走主题 minorFont；lang en-US / zh-CN。无默认段落属性。
- 主题（WPS）：majorFont latin=Calibri Light，minorFont latin=Calibri；
  两者 ea 为空，`script Hans → 宋体`。因此 `majorEastAsia`/`minorEastAsia`
  实际解析为**宋体**。
- 中文黑体（标题）、宋体（正文）均通过 `w:rFonts w:eastAsia` 显式指定，
  西文一律 Times New Roman（继承自 Normal 样式）。

## 3. 核心样式表（styleId 对照）

### Normal（id `a0`，所有样式的基准）
- pPr：孤行控制关；`spacing line=360 lineRule=auto`（**1.5 倍行距**）；两端对齐
- rPr：`rFonts ascii/hAnsi="Times New Roman"`（中文=主题宋体）、kern=2、
  `sz/szCs=24`（**12pt 小四**）

### 标题族
| 样式 | styleId | 字体 | 字号 | 加粗 | 对齐 | 段距/行距 | 大纲级别 |
|---|---|---|---|---|---|---|---|
| heading 1 | `1` | 黑体 + TNR | 32（16pt） | 是 | 实际使用时直接居中 | before 340 / after 330，line 300（1.25 倍） | 0 |
| heading 2 | `2` | 宋体(主题major) + TNR | 32（16pt） | 是 | 继承两端 | before/after 260，line 416 auto | 1 |
| heading 3 | `3` | 宋体 + TNR | 28（14pt） | 是 | 继承两端 | before/after 260，line 416 auto | 2 |
| heading 4-9 | `4`–`9` | 宋体/Cambria | 24 | 部分 | 左 | 固定行距 440 等 | 3-8（模板正文未用） |

标题编号**不是自动编号**，是标题文字的一部分（一. / 1.1 / 1.1.1）。

### 其他
| 样式 | styleId | 要点 |
|---|---|---|
| caption 题注 | `af0` | 宋体，sz 21（**10.5pt**），居中；模板用法："表 "+`SEQ 表 \* ARABIC` 域+" 标题" |
| table of figures | `af2` | 图表目录条目；实际条目直接格式：右对齐点线制表位 `pos=8296`、`ind left=960 hanging=480` |
| toc 1/2/3 | `TOC1/2/3` | 条目分别缩进 0 / 420 / 840 缇；直接格式补点线制表位 8296 |
| TOC Heading | `TOC` | 基于 heading 1、取消加粗、颜色覆盖为 auto；实际"目录"标题直接格式：黑体、居中 |
| 正文 缩进2字符 | `21` | 封面间隔空段用：line 400 exact、firstLine 560 |
| footer 页脚 | `a5` | 9pt，居中 PAGE 域 |
| Table Grid | `aa` | 六边 `single sz=4`（0.5pt）全框线；基于 Normal Table（`a3`，cellMar 左右 108） |
| Hyperlink | `af` | 目录条目超链接字符样式，色 0026E5 |

## 4. 目录机制（Word 域，非静态文本）

- 主目录：`TOC \o "1-3" \h \z \u`，收集大纲级别 0-2 的段落（即标题 1-3）
- 图目录：`TOC \h \z \c "图"`；表目录：`TOC \h \z \c "表"`——收集含对应
  SEQ 标识符的题注段，Word 更新时自动建 `_Toc` 书签和 PAGEREF 超链接
- 条目格式由 TOC1/2/3、af2 样式 + 生成器写入的点线制表位呈现
- "目录"标题：黑体 16pt 居中；"图片"/"表格"标题：黑体 12pt 居中（pStyle af2）
- `settings.xml` 的 `updateFields=true` 使 Word/WPS 打开时自动刷新；
  题注 SEQ 域的缓存值由生成器写为真实序号，保证单次刷新即正确

## 5. 封面（底版保留标题封面）

无边框表格（tblW 8276、居中、fixed、cellMar 左右 0），4 列 grid：888/2193/3712/1483，
9 行，全部 vAlign=center，行高：1984 / 3402 / 1664 / 886 / 883×5。

- 行1（标题区）：跨 4 列，标题黑体 sz=52（**26pt**）居中，行距 line 240 auto
- 行3-7（签字栏）：模板标题封面留空；列2 标签 16pt 加粗右对齐、字间距 w:spacing=20，
  列3 为 14pt 下划线空格填写线
- 行8（落款）：跨 4 列，"北京航空航天大学"+"2026年 10月"，黑体 sz=32（16pt）居中，
  行距 line 400 exact

原始模板还有第 2 页签字封面（编制/审核/标准化/会签/批准），按用户要求生成的
报告**只保留标题封面**；原始两封面结构见 `assets/项目报告模板.docx`。
另：原模板签字封面后的间隔段会溢出一张空白页（模板自身瑕疵），底版已消除。

## 6. 表格规格（两种）

### grid（数据表，模板表 5-9）
- `tblStyle=aa`（Table Grid 0.5pt 全框线）、`jc=center`、tblW auto
- 行高 trHeight 433 atLeast；单元格 vAlign=center；内容居中
- 表头行：浅灰底 `w:shd fill="F3F3F3"`、加粗、居中；tblHeader + cantSplit

### threeline（框架/流程表，模板表 4）
- 无 tblStyle；tblBorders：top/bottom `single sz=12`（1.5pt 粗线）、
  insideH/insideV `single sz=4`（0.5pt 细线）、**无左右外框**
- 表头行单元格 tcBorders bottom `sz=12`（表头下粗线），自动加粗
- 列宽用 tblLayout fixed + gridCol/tcW；长文本单元格左对齐

## 7. 插图与引用

- 图片：inline 居中（`distT/B/L/R=0`），宽 ≤ 版心 14.5cm，题注在图下方
- 正文引用（GB/T 7714 顺序编码）：上标 `[1]`（Times New Roman superscript）
- 参考文献：heading 1 样式"参考文献"（黑体居中，无编号）+ 条目悬挂缩进
  （leftChars/hangingChars 200，left/hanging 480），条目 `[N] 作者. 题名[文献类型].
  出版地: 出版社, 年份.` 中英文均 TNR/宋体 12pt

## 8. 域清单（底版/生成器产出）

`TOC \o "1-3" \h \z \u`、`TOC \h \z \c "图"`、`TOC \h \z \c "表"`、
每题注一个 `SEQ 图|表 \* ARABIC`、页脚 `PAGE \* MERGEFORMAT`。
目录刷新后自动生成 `PAGEREF _Toc… \h` 与隐藏书签。
