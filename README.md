## 📦 安装方法

1、将本仓库下载到 Codex 数据目录下的 `skills/my-word-style` 文件夹，默认位置是 `~/.codex/skills/my-word-style`。

2、在 Codex 中发送下面的使用指令。

## 💬 使用方法

> 使用 $my-word-style，把这些实验结果整理成周报。

> 使用 $my-word-style，更新这份 Word，保留原有样式。

> 使用 $my-word-style，合并这几组实验，加入方案选择过程、优缺点对照、指标图和表格。

## ⚡️ 排版规则

| 内容 | 样式与格式 |
| --- | --- |
| 正文 | 正文1 |
| 标题 | 标题1、标题 2、标题 3 |
| 图片 | TOC 1 居中，彩色图片，图内英文 Serif 字体 |
| 图注 | 图片下方左对齐，`Fig.1. 简略说明` |
| 表格 | 黑白三线表，与编号、说明完整放在同一页 |
| 表题 | 表格上方分两行写 `Table 1` 和简略说明 |
| 公式 | Word 原生可编辑公式，居左，编号 `(1)` 居右 |
| 页码 | 底部居中，首页不显示 |

## 🛠️ 模板与脚本

原 Word 样式保存在 [模板文件](assets/word-styles-template.docx) 中，可替换为自己的模板。完整规则见 [SKILL.md](SKILL.md)。

辅助脚本需要 `python-docx`，在仓库目录运行：

```bash
python scripts/word_style.py new --output report.docx
python scripts/word_style.py check report.docx
```
