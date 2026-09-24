# 教学设计 DOCX 样式规范

> 适用范围：`teaching-plan-generator` 技能产出的教学设计 DOCX（统一版 4 列流程表 / 分层版 5 列流程表）
> 基准样本：`计算机网络基础_计算机网络的产生与发展、定义与功能_教学设计.docx`（2026-09-24 定稿）
> 用途：后续课次直接按本规范生成或校对版式，保证同一门课的多份教案外观完全一致。

---

## 0. 量纲与字号对照

| 名称 | pt | 半磅值（OOXML `w:sz`） | python-docx |
|------|----|------------------------|-------------|
| 小二 | 18 | 36 | `Pt(18)` |
| 三号 | 16 | 32 | `Pt(16)` |
| 小四 | 12 | 24 | `Pt(12)` |
| 五号 | 10.5 | 21 | `Pt(10.5)` |
| 小五 | 9 | 18 | `Pt(9)` |

- 长度换算：`1 cm = 360000 EMU = 567 twips`，`1 pt = 12700 EMU`
- 正文可用宽度 = 21 − 3.0 − 3.0 = **15.0 cm**（所有表格列宽合计必须等于 15.0 cm）

---

## 1. 页面与节设置

| 项目 | 值 |
|------|-----|
| 纸张 | A4，21.0 × 29.7 cm |
| 页边距 | 上 2.54 / 下 2.54 / 左 3.00 / 右 3.00 cm |
| 页眉 | **无**（保持空白，`is_linked_to_previous = True`） |
| 页脚 | 居中，宋体 9pt，内容：`第 X 页  共 Y 页`（X/Y 之间为两个半角空格） |
| 页眉/页脚距边界 | 1.27 cm |
| 文档默认样式 Normal | 不设显式字号，靠各段落显式指定 |

页码实现（域代码）：

```python
# w:fldSimple instr=" PAGE " / " NUMPAGES "，run 上设 rFonts=宋体 + sz=18(9pt)
def add_page_field(paragraph):
    for instr in (' PAGE ', ' NUMPAGES '):
        fld = OxmlElement('w:fldSimple')
        fld.set(qn('w:instr'), instr)
        run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rf = OxmlElement('w:rFonts'); rf.set(qn('w:eastAsia'), '宋体'); rPr.append(rf)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '18'); rPr.append(sz)
        run.append(rPr)
        fld.append(run)
        paragraph._p.append(fld)
```

---

## 2. 段落样式

| 元素 | 样式名 | 字体（中/西） | 字号 | 加粗 | 对齐 | 段前 | 段后 | 行距 |
|------|--------|---------------|------|------|------|------|------|------|
| 文档大标题 | Normal（手改） | 黑体 | 18 pt（小二） | 是 | 居中 | 0 | 14 pt | 默认 |
| 章节标题 | **Heading 1** | 黑体 | 16 pt（三号） | 是 | 左 | 10 pt | 6 pt | 默认（单倍） |
| 正文段落 | Normal | 宋体 | 12 pt（小四） | 否 | 左 | 0 | 0 | **1.5 倍** |
| 三级粗体小标签 | Normal | 宋体 | 12 pt | **是** | 左 | 8 pt | 2 pt | 1.5 倍 |
| 板书树形图行 | Normal | 宋体 | 10.5 pt（五号） | 否 | 左 | 0 | 0 | 单倍（1.0） |

- **三级粗体小标签**指：`知识目标` / `技能目标` / `素养目标` / `主板书：` / `副板书：` 等分组小标题，独占一段，不带编号。
- 列表一律用纯文本前缀（`- ` 或 `1. `），**不使用** Word 自动编号 / 项目符号（避免编号错乱与缩进漂移）。

### 2.1 标题颜色（关键）

Heading 1 必须显式覆盖为纯黑，否则会继承 Word 内置主题蓝 `#365F91`：

```python
# 必须同时删除 themeColor / themeShade / themeTint，否则主题色仍然生效
c = style.element.rPr.find(qn('w:color'))
for attr in ('w:themeColor', 'w:themeShade', 'w:themeTint'):
    if c.get(qn(attr)) is not None:
        del c.attrib[qn(attr)]
c.set(qn('w:val'), '000000')
```

- 本规范未使用 Heading 2 / Heading 3；Word 默认的 Heading 2/3 仍为主题蓝 `#4F81BD`，若将来启用需一并改为黑色或另定色。

---

## 3. 表格通用规格

| 项目 | 值 |
|------|-----|
| 表格样式 | `Table Grid` |
| 自动调整 | 关闭（`table.autofit = False`） |
| 布局算法 | `w:tblLayout type="fixed"` |
| 列宽 | 逐列逐格显式设置，合计 15.0 cm |
| 表头跨页重复 | 表头行 `trPr/w:tblHeader`（无 `val` 属性） |
| 底纹 | `w:shd w:fill="..."`，插入位置须在 `w:noWrap / w:tcMar / w:textDirection / w:tcFitText / w:vAlign / w:hideMark` **之前** |
| 垂直居中 | `w:vAlign w:val="center"` |
| 表头行单元格段落 | 段前 2 pt / 段后 2 pt |
| 信息表单元格段落 | 段前 1 pt / 段后 1 pt / 行距 1.0 |

底纹与表头重复的 OOXML 顺序保障：

```python
tcPr.insert_element_before(shd, 'w:noWrap', 'w:tcMar', 'w:textDirection',
                           'w:tcFitText', 'w:vAlign', 'w:hideMark')
trPr.insert_element_before(tblHeader, 'w:cantSplit', 'w:trHeight')
```

### 3.1 表 0 —— 基础信息表（9 行 × 2 列）

| 项目 | 值 |
|------|-----|
| 列宽 | 4.0 / 11.0 cm |
| 标签列（第 0 列） | 黑体 **10.5 pt** 加粗、居中、底纹 `F2F2F2`、垂直居中 |
| 内容列（第 1 列） | 宋体 **10.5 pt**、左对齐、无底纹、垂直居中 |
| 单元格段落 | 段前 1 pt / 段后 1 pt / 行距 1.0 |
| 表头重复 | 无 |

字段顺序（**不带任何编号**）：

1. 课程名称　2. 课题名称　3. 授课班级　4. 授课人数　5. 学段
6. 教育类型　7. 使用教材　8. 教学课时　9. 授课类型

### 3.2 表 1 / 表 2 —— 教学流程表（统一版 17 行 × 4 列 / 分层版 5 列，每学时一张）

| 项目 | 值 |
|------|-----|
| 列宽 | 2.0 / 6.4 / 3.4 / 3.2 cm（时间环节 / 教师活动 / 学生活动 / 设计意图） |
| 表头行（第 0 行） | 黑体 **12 pt** 加粗、居中、底纹 `D9D9D9`、垂直居中、段前 2 / 段后 2、`tblHeader=True` |
| 分段带行 | **整行 4 列横向合并**；黑体 12 pt 加粗、**左对齐**、底纹 `EDEDED`、垂直居中、段前 2 / 段后 2 |
| 数据行 | 宋体 **12 pt**、左对齐、无底纹、**不设**垂直居中、段前 2 / 段后 2 |

分段带行分两级：

- **学时行**：`【第一学时 0—40分钟】1.1 计算机网络的产生与发展`
- **模块行**：`第一模块：破冰筑基（0—10分钟）—— 建立“四代演进”时间轴框架`

固定行结构（每学时：统一版 17 行；分层版模块三拆双实操时 18 行；序号从 0 计。以所加载模板为准，下表为统一版示例）：

| 行号 | 内容 | 类型 |
|------|------|------|
| 0 | 时间/环节 · 教师活动 · 学生活动 · 设计意图 | 表头 |
| 1 | 【第X学时 a—b分钟】小节标题 | 分段带（合并） |
| 2 | 第一模块：破冰筑基 | 分段带（合并） |
| 3–5 | 导入锚定 / 核心授课 / 小结 | 数据 |
| 6 | 第一次缓冲（a—b分钟）：…，不讲新知识 | 数据 |
| 7 | 第二模块：重难点攻坚 | 分段带（合并） |
| 8–10 | 衔接与价值锚定 / 重难点精讲 / 深度提问 | 数据 |
| 11 | 第二次缓冲（a—b分钟）：…，不讲新知识 | 数据 |
| 12 | 第三模块：落地实操 | 分段带（合并） |
| 13–15 | 知识串联 / 案例实操 / 要求梳理 | 数据 |
| 16 | 终局复盘（a—b分钟）：梳理核心＋布置课后任务＋留白答疑 | 数据 |

- 「时间/环节」列数据行文案格式：`0-2min 导入锚定`（数字与 `min` 之间无空格，半角连字符）。
- 缓冲行与终局复盘行不写短标签，直接写整段说明文字。
- **多学时连堂**：每个 40 分钟学时各出一张流程表。第二学时的表（表 2）在**表首行单元格的段落**上设 `page_break_before = True`，强制另起一页。
- 相邻两张 `w:tbl` 之间若无段落间隔，Word 会视觉上合并为一张表 —— 这也是用 `pageBreakBefore` 分页的原因之一。

### 3.3 表 3 —— 评价设计表（7 行 × 5 列）

| 项目 | 值 |
|------|-----|
| 列宽 | 2.2 / 3.4 / 5.4 / 2.2 / 1.8 cm（评价维度 / 评价指标 / 评价标准（分值）/ 评价主体 / 评价方式） |
| 表头行 | 黑体 **12 pt** 加粗、居中、底纹 `D9D9D9`、垂直居中、段前 2 / 段后 2、`tblHeader=True` |
| 数据行 | 宋体 **12 pt**、左对齐、无底纹、段前 2 / 段后 2 |
| 合计行（末行） | 第 0–1 列合并写「合计」，第 2 列写「100分」，第 3–4 列合并留空；黑体 12 pt 加粗、居中、底纹 `E7E7E7`、垂直居中 |

分值行建议保持 5 行（学习态度 / 学习行为 / 学习效果（过程）/ 学习效果（结果）× 2），合计 100 分。

---

## 4. 章节编号与顺序约定

基础信息表**不带编号**，其后 12 节用「中文数字＋顿号」编号，标题一律为 Heading 1：

| 序 | 标题 |
|----|------|
| — | （基础信息表） |
| 一 | 学情分析 |
| 二 | 教学目标 |
| 三 | 教学方法 |
| 四 | 学习方法 |
| 五 | 教学重点及解决措施 |
| 六 | 教学难点及突破措施 |
| 七 | 教学资源 |
| 八 | 教学思路 |
| 九 | 教学流程 |
| 十 | 评价设计 |
| 十一 | 板书设计 |
| 十二 | 教学反思 |

> 注：学情分析在前、教学目标在后（用户 2026-09-24 明确调整）。若后续课次需要回到「教学目标 → 学情分析」的旧序，全表编号需整体顺延调整。

---

## 5. 版式相关的写作约束

1. **板书树形图对齐**：行首 `├─` `└─` `│` 为全角字符，缩进必须用**全角空格 U+3000**（宋体下 1 个 `│` 占 1 个全角宽度，一级缩进 = 4 个 U+3000）；混用半角空格会错位。
2. **板书树形图字号**：10.5 pt、单倍行距，与正文小四区分，视觉上更像「板书」。
3. 三级小标签（知识目标等）与其下条目之间不空行，靠段前 8 pt 形成分组。
4. 「教学资源」用 `- ` 前缀分四类：设备 / 工具·仪器 / 软件·平台 / 资料·文档。
5. 「教学重点及解决措施」「教学难点及突破措施」每条用 `序号. 内容 —— 解决措施：…` 的「破折号＋标签」结构，破折号用 `——`（两个全角破折号）。
6. 「教学反思」正文为占位符 `（课后填写）`。
7. 80 分钟连堂的表述统一为「两个 40 分钟的三段式单元」，模块内部时间标尺仍为 10 + 3 + 10 + 3 + 10 + 4。

---

## 6. 已知陷阱与实现要点

**路径划分**：#1–#10 适用本技能标准路径（`generate_docx.py` 读 JSON 写 DOCX）；#11–#15 仅适用**交互式编辑已有 DOCX** 的工具路径（本技能生成流程不会调用 `doc_set_table_cells` / `save_file` / `close_file`，可跳过）。

| # | 适用路径 | 陷阱 | 处理 |
|---|---------|------|------|
| 1 | 脚本生成 | 字段标题继承 Heading 1 主题蓝 `#365F91` | 显式写 `w:color val="000000"`，并删除 `themeColor/themeShade/themeTint` |
| 2 | 脚本生成 | python-docx 默认表格列宽均分 → 长文本列严重折行 | `autofit=False` + `w:tblLayout=fixed` + 逐列逐格设宽 |
| 3 | 脚本生成 | 表头无底纹 | 加 `w:shd`，注意 `tcPr` 子元素顺序（见 §3） |
| 4 | 脚本生成 | 表头跨页不重复 | `trPr` 加 `w:tblHeader`，同样注意元素顺序 |
| 5 | 脚本生成 | 无页码 | 页脚加 `w:fldSimple`（`PAGE` / `NUMPAGES`） |
| 6 | 脚本生成 | 板书树形图错位 | 行首缩进半角空格全部替换为 U+3000 |
| 7 | 脚本生成 | **合并单元格会吞并空段落** | `cell.merge()` 会把被并入格各自的空段落并进结果格，产生 N−1 个多余 ¶；判据：`len({id(c._tc) for c in row.cells}) < len(row.cells)`；清理后必须复查 |
| 8 | 脚本生成 | 相邻两张表被 Word 视觉合并 | 用 `pageBreakBefore` 或插入空段落分隔 |
| 9 | 脚本生成 | 多学时连堂 | 每学时一张流程表，第二张首行 `page_break_before=True` |
| 10 | 脚本生成 | 新插入段落继承单倍行距 | 对正文段补 `line_spacing=1.5` |
| 11 | 交互编辑 | `doc_set_table_cells` 的 `text_format` 字段名是 **`font_size`**（不是 `font_size_pt`） | 键名写错会报 `[-16] … no effective operations` 或被静默忽略 |
| 12 | 交互编辑 | 只传 `text` 替换单元格会丢字号 | 需随后用 `text_format:{"font_size":…}` 补回 |
| 13 | 交互编辑 | `doc_delete_paragraph` 在表格单元格内不可靠 | 单元格内清段落改用 python-docx 脚本一次做完 |
| 14 | 交互编辑 | `save_file` 省略 `file_path` 不落盘 | 必须显式 `file_path=`，并用 mtime + 重新读取双重确认 |
| 15 | 交互编辑 | 文件已被编辑器打开时改盘 | 先 `close_file` force 关闭再写 |

---

## 7. 参数速查（可直接写入生成/校对脚本）

```python
from docx.shared import Pt, Cm

# 页面
PAGE = dict(w=Cm(21.0), h=Cm(29.7),
            top=Cm(2.54), bottom=Cm(2.54), left=Cm(3.0), right=Cm(3.0))

# 段落
TITLE    = dict(font='黑体', size=Pt(18),   bold=True, align='center', before=Pt(0),  after=Pt(14))
H1       = dict(font='黑体', size=Pt(16),   bold=True, align='left',   before=Pt(10), after=Pt(6))
BODY     = dict(font='宋体', size=Pt(12),   bold=False, line=1.5)
LABEL    = dict(font='宋体', size=Pt(12),   bold=True,  line=1.5, before=Pt(8), after=Pt(2))
BOARD    = dict(font='宋体', size=Pt(10.5), bold=False, line=1.0, before=Pt(0), after=Pt(0))

# 表格
T0 = dict(widths=[4.0, 11.0], label_font='黑体', label_size=Pt(10.5), label_bold=True,
          label_align='center', label_shd='F2F2F2', body_font='宋体', body_size=Pt(10.5),
          cell_before=Pt(1), cell_after=Pt(1), cell_line=1.0, header_repeat=False)

T12 = dict(widths=[2.0, 6.4, 3.4, 3.2], header_shd='D9D9D9', band_shd='EDEDED',
           header_font='黑体', header_size=Pt(12), header_bold=True, header_align='center',
           band_font='黑体', band_size=Pt(12), band_bold=True, band_align='left',
           body_font='宋体', body_size=Pt(12), cell_before=Pt(2), cell_after=Pt(2),
           header_repeat=True, second_sheet_page_break=True)

T3 = dict(widths=[2.2, 3.4, 5.4, 2.2, 1.8], header_shd='D9D9D9', total_shd='E7E7E7',
          header_font='黑体', header_size=Pt(12), header_bold=True, header_align='center',
          total_font='黑体', total_size=Pt(12), total_bold=True, total_align='center',
          body_font='宋体', body_size=Pt(12), cell_before=Pt(2), cell_after=Pt(2),
          header_repeat=True)

# 页脚
FOOTER = dict(font='宋体', size=Pt(9), align='center', text='第 {PAGE} 页  共 {NUMPAGES} 页')
```

---

## 8. 交付前检查清单

- [ ] 页边距 上2.54 / 下2.54 / 左3.0 / 右3.0 cm，页眉为空，页脚为「第 X 页  共 Y 页」
- [ ] 章节标题为 Heading 1 且颜色为纯黑（非主题蓝）
- [ ] 12 个章节编号连续、中文数字正确，基础信息表无编号
- [ ] 4 张表列宽合计均为 15.0 cm，且 `tblLayout=fixed`
- [ ] 表头行底纹 `D9D9D9` + 跨页重复；分段带行底纹 `EDEDED` 且整行合并；合计行底纹 `E7E7E7`
- [ ] 教学流程表 / 评价表单元格字号 = 小四（12 pt）；信息表 = 五号（10.5 pt）
- [ ] 无合并单元格残留的多余空段落（逐行核对 `len({id(c._tc) for c in row.cells})`）
- [ ] 板书树形图缩进为 U+3000，无半角空格
- [ ] 第二学时流程表首行设 `page_break_before`
- [ ] 保存后核对文件 mtime 与内容（确认已真正落盘）
