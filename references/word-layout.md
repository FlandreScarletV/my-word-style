# 样式与Word排版实现

## 用户样式映射

模板由用户原Word的样式与页面设置提取，正文内容和原图片已清空。样式定义不是根据截图估计的。

| 用途 | Word中指定的名称 | 内置模板的python-docx名称 | style ID |
|---|---|---|---|
| 正文、表题、图注、单元格、公式段落 | 正文1 | 正文1 | 11 |
| 一级标题 | 标题1（无空格，自定义） | 标题1 | 13 |
| 二级标题 | 标题 2 | Heading 2 | 2 |
| 三级标题 | 标题 3 | Heading 3 | 3 |
| 图片段落 | TOC 1 | toc 1 | TOC1 |

“图像小字标题”（ID a5）仍保留在模板中，但用户后续明确要求说明文字使用“正文1”，因此目前不用它作为默认图注。“Normal”或内置“Heading 1”（ID 1）都不是指定的正文/一级标题。

原“正文1”为Times New Roman西文、宋体中文、12pt，约1.25倍行距、两字符首行缩进；“标题1”为16pt、粗体且带段前分页。二/三级标题分别14/12pt，保留其设置。模板A4，边距上下2.54cm、左右3.175cm。以模板实际定义为准，不重复手写这些值。首个标题可局部取消段前分页，图注、表题、表格内段落和公式可局部清零首行缩进。

不要全局修改字体、行距或页面网格来修复局部表格。保留所有原样式；渲染环境没有宋体等字体时可能采用字体替代，不能据此把文档统一替换为其他字体。

## 可复用的辅助模块

`word_style.py`是可导入模块，也支持最小命令行：

```bash
python scripts/word_style.py new --output /path/to/report.docx
python scripts/word_style.py check /path/to/report.docx
```

`new`只创建空白文档；不添加标题、封面或示例内容。没有Word/PDF转换功能。依赖python-docx；使用当前环境已有可用解释器，不默认新建环境。

导入后可调用：

```python
import sys
from pathlib import Path
skill_root = Path('/actual/installed/path/my-word-style')
sys.path.insert(0, str(skill_root / 'scripts'))
from word_style import (
    new_document, add_body, add_heading, add_table, add_figure,
    add_equation, math_run, fraction, subscript, math_node,
    keep_table_together, check_document,
)

doc = new_document()
add_heading(doc, '1. 实验设置', level=1)
add_body(doc, '填写本次实际做法。')
add_table(doc, ['方案', '结果'], [['方案A', '填写实际值']], '方案对照。')
# 使用原生分数结构，不把LaTeX字符串当成已转换公式。
add_equation(doc, [
    math_run('Precision', upright=True), math_run('=', upright=True),
    fraction([math_run('TP', upright=True)], [math_run('TP+FP', upright=True)]),
])
# add_figure(doc, actual_image_path, '实际图片说明。')
doc.save('/path/to/report.docx')
check_document('/path/to/report.docx')
```

`new_document(template=...)`只用于从模板新建；编辑已有正文时用`Document(existing_path)`并复用`style_for`、`keep_table_together`等操作，不能调用新建函数清掉用户原文。

## 整表不跨页

辅助函数将两个表题段落和所有表格行连成keep-with-next链，最后一行解除，所有行设置cantSplit；取消单元格网格吸附，去掉表头重复设置。`keep_table_together`可以直接用于既有表格，默认识别其前方两段Table编号和说明。

如果整个块超过可用页高，任何keep设置都不能凭空让它放下。依次考虑：移到新页、调整列宽和换行、收紧单元格段前后距、调整可读字号；宽表可按任务需要单独横向页面。不要自动拆成“续表”或删数据。当前页面约束下仍无法完整容纳时，需要明确解决冲突，不能交付跨页表后声称已完成。

## 原生公式

`add_equation`接受OMML `<m:oMath>`或其子节点列表，添加右侧SEQ Equation字段和缓存编号。`math_run`、`fraction`、`subscript`只是构造原生数学结构的便捷方法，复杂公式按OMML结构补充；不支持把任意LaTeX字符串自动排成专业公式。

公式中的变量可遵循数学字体惯例，命名函数用直立体；图表“常规非斜体”的要求不应误套到数学变量。保存后应能在Word中直接选中并用公式工具编辑。

## 彩色图表

建议采用以下设置，而不是重新生成黑白图片：

```python
plt.rcParams.update({
    'font.family': 'Liberation Serif',  # 可用时优先Times New Roman
    'font.weight': 'normal',
    'font.style': 'normal',
    'axes.labelweight': 'normal',
    'axes.titleweight': 'normal',
})
# 横向分组条形：总高度1.25，两条中心±0.23，条高0.32，留下可见间隙。
y = np.arange(len(categories)) * 1.25
ax.barh(y - 0.23, values_a, height=0.32, color='#4C78A8', label='Raw scores')
ax.barh(y + 0.23, values_b, height=0.32, color='#F58518', label='Post-processed')
ax.set_axisbelow(True)
ax.grid(axis='x', color='#d1d1d1', linewidth=0.65)
```

图内英文标签，数值不加粗、不斜体，横轴灰线随实际刻度绘制；不是照抄上述尺寸的硬要求。图注“Fig.X.”放在Word段落，不烧录成图像上方大标题。

## 实际分页与不生成PDF的核对

已有Word/Writer会话时直接检查。需要自动化时可在临时配置目录启动本地Writer UNO会话，只读打开docx：

- `controller.getViewCursor()`配合`gotoRange`、`getPage()`，记录每个Table编号、说明、第一单元格起点及最后单元格终点的页号，必须一致。
- 用`getEmbeddedObjects()`检查导入后可编辑公式对象与其锚点。
- 需要截图时，使用`storeToURL`、`FilterName='writer_png_Export'`，FilterData包括PageNumber、PixelWidth、PixelHeight，直接输出临时PNG。结束后关闭文档，不通过Writer另存覆盖原Word。
- 不调用PDF导出，不产出HTML/ODT等第二份交付文档。没有可用排版引擎时只报告已做的结构检查，不把结构检查描述成实际页码验收。

脚本`check`检查样式、OMML计数、keep链、编号和页脚的结构，不承诺能测量实际页高；实际页号检查仍需Word/Writer。
