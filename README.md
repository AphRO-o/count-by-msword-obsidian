# Count By MSWord

[English](#count-by-msword) | [中文说明](#中文说明)

Count words with the Microsoft Word desktop app and export Markdown notes to DOCX. Word statistics and DOCX export use the same conversion rules, so both work with the same document content.

**Language support:** This plugin currently only supports Chinese. Other languages have not been tested.

## Features

- Show the current note's Word count in the status bar, including unsaved editor changes.
- Hover over the count to see words, characters with and without spaces, paragraphs, lines, pages, and Asian-language characters.
- Refresh on click or when returning to a changed note. Editing marks the result as stale; it does not start Word after every keystroke.
- Export through the hover panel, a file's context menu, or the command palette.
- Configure fonts, font size, line spacing, indentation, full-width or half-width punctuation, and blank-line handling.
- Generate compressed DOCX files and keep previous statistics for unchanged notes.

## Requirements

- Windows 10 or Windows 11.
- Obsidian desktop 1.5.0 or later.
- A working installation of Microsoft Word desktop. Open Word once to complete any first-run setup or activation.

## Installation

1. Download `main.js`, `manifest.json`, and `styles.css` from the latest GitHub release.
2. Create a folder named `count-by-msword` inside your vault's `.obsidian/plugins/` folder.
3. Put all three files in `.obsidian/plugins/count-by-msword/`.
4. Restart Obsidian.
5. Open **Settings → Community plugins** and enable **Count By MSWord**.

## Usage

1. Open a Markdown note and wait for the status bar to show its Word count.
2. After editing, the label `待更新` means the displayed result needs refreshing. Click the status bar to refresh, or switch to another note and return with automatic updates enabled.
3. Hover over the status bar to view the full statistics and the last calculation time. Line and page counts follow the exported document's layout.
4. To export, click `点击导出 DOCX` in the hover panel, right-click a Markdown file and select `导出为 Word (.docx)`, or search for **Export current note to Word** in the command palette.
5. Enter a full Windows output path, such as `D:\Writing\My note.docx`. Confirm before replacing an existing file.

The command palette also provides **Refresh Word word count** and **Clear Word count cache**. Command names include both Chinese and English; the plugin interface is currently in Chinese.

## Export settings and content handling

Open **Settings → Count By MSWord** to change the output directory and formatting. Defaults are SimSun (宋体) for Chinese, Times New Roman for English, 12 pt text, 1.5 line spacing, two-character first-line indentation, Chinese full-width punctuation, and removal of blank separator lines between paragraphs.

- Common headings, emphasis, lists, blockquotes, and link labels are retained.
- Images contribute their alternative text. Embedded notes contribute their label or target name rather than expanding the other note.
- Tables are simplified to rows of cell text.
- Footnotes become body text at the end of the document and are included in the count.
- Code and formula source text are retained. Punctuation and blank lines inside code, formulas, and YAML are preserved.
- YAML frontmatter is excluded by default; enable its inclusion in settings if needed.

Changing content or formatting settings invalidates the corresponding statistics. Click the status bar to calculate again. Detailed conversion rules are available at the bottom of the settings page.

## Local processing and system permissions

The plugin processes notes locally and makes no network requests or uploads.

- **Filesystem access (`fs`)** creates and removes temporary DOCX files in the Windows temporary directory and saves exported DOCX files to the local path you choose. The original Markdown note is preserved.
- **Shell execution (`child_process`)** starts a hidden Windows PowerShell helper to automate Microsoft Word through COM. It opens the temporary DOCX read-only, obtains statistics, and closes the document and the Word instance it created.
- The helper has a configurable timeout and cancellation cleanup. Only Word processes identified as created by the plugin are terminated during cleanup.

These capabilities are used for local Word automation and document export. The plugin is desktop-only.

## Troubleshooting

If statistics fail, check that Word opens normally and has completed setup or activation. Enable detailed errors in the plugin settings to see the cause. Please report problems in this repository's Issues page.

## License

Licensed under the [MIT License](LICENSE). Copyright (c) 2026 AphRO.

## 中文说明

在 Obsidian 里写作，用 Microsoft Word 看字数，写完直接导出 Word 文档。

适合平时用 Obsidian 写稿、投稿时又要按 Word 字数核算的作者。统计由你电脑上的 Word 完成，文稿全程在本地处理。

### 能做什么

- **看 Word 字数**：右下角直接显示当前笔记的字数，包括编辑器里还没保存的修改。
- **看完整统计**：鼠标放到字数上，可以查看由 Microsoft 统计的字符数（计空格／不计空格）、段落数、行数、页数和亚洲语言字符数。
- **导出 Word 文档**：把当前笔记另存为 `.docx`，原笔记保留。
- **调整导出格式**：可以设置中英文字体、字号、行距、首行缩进、标点形式和是否保留空行。

统计和导出使用同一套内容转换规则。你统计的内容，就是插件准备导出的文档内容。行数和页数按导出后的文档排版计算。

### 使用前需要

- Windows 10 或 Windows 11。
- Obsidian 桌面版，版本 1.5.0 或以上。
- 电脑上已安装并能正常使用的 Microsoft Word 桌面版。

### 安装

1. 在你的 Obsidian 仓库中找到 `.obsidian/plugins/` 文件夹。
2. 新建 `count-by-msword` 文件夹，把 `main.js`、`manifest.json` 和 `styles.css` 放进去。
3. 重新启动 Obsidian。
4. 在「设置 → 第三方插件」中启用 **Count By MSWord**。

安装后的目录如下：

```text
你的仓库/
└── .obsidian/
    └── plugins/
        └── count-by-msword/
            ├── main.js
            ├── manifest.json
            └── styles.css
```

### 平时怎么用

打开一篇 Markdown 笔记，等右下角出现类似这样的结果：

```text
Word: 5,218 字
```

继续修改内容后，字数旁边会显示「待更新」。点击字数，就会重新统计。

默认情况下，切换到其他笔记再回来，也会检查是否需要更新。内容没变化时会直接使用上次结果；打字时只标记待更新，等你切换笔记或手动刷新时再调用 Word。

把鼠标放到字数上，可以查看详细统计、上次统计时间和当前笔记路径。

### 怎么导出

任选一种方式：

- 在统计悬浮窗里点击 **「点击导出 DOCX」**。
- 在笔记文件上右键，选择 **「导出为 Word (.docx)」**。
- 按 `Ctrl + P` 打开命令面板，搜索 **「导出当前笔记为 Word」**。

导出时输入完整的保存路径，例如 `D:\文稿\我的文章.docx`。默认位置是当前笔记所在的文件夹，也可以在插件设置里指定常用目录。目标文件已存在时，会先询问是否覆盖。

### 常用设置

在「设置 → Count By MSWord」里调整：

| 设置 | 默认值／作用 |
| --- | --- |
| 中文字体、英文字体 | 宋体、Times New Roman |
| 正文字号 | 12 pt |
| 行距 | 1.5 倍 |
| 首行缩进 | 2 个汉字 |
| 标点符号形式 | 默认中文全角，也可以选择英文半角；中文引号使用中文字体 |
| 保留空行 | 默认关闭：空行用于分段，导出时不额外插入空白段落；开启后保留 |
| 默认导出目录 | 留空时使用当前笔记所在文件夹 |
| 自动统计、返回笔记时更新 | 默认开启，可按自己的习惯关闭 |

修改字体、排版或内容相关设置后，旧统计会失效，点击状态栏即可重新统计。

### 笔记内容会怎么处理

支持常用的标题、粗体、斜体、删除线、列表、引用和链接。下面几类内容有明确的处理方式：

- **链接和嵌入**：保留显示文字；嵌入的其他笔记保留别名或名称，正文不展开。
- **图片和表格**：图片保留说明文字；表格按行保留单元格文字。
- **脚注**：放到文末作为正文，计入字数。
- **代码和公式**：保留代码文字、公式源码；其中的标点和空行保留原样。
- **笔记开头的属性信息**：也就是 YAML，默认排除；可以开启「包含 YAML」一起导出。

更完整的规则可以在插件设置页底部查看。

### 其他命令

命令面板里还可以找到：

- **刷新 Word 字数**：立即重新统计当前笔记。
- **清除 Word 字数缓存**：清掉已保存的统计结果。

如果出现「统计失败」，先确认 Word 能正常打开，并完成首次启动或授权设置。需要查看具体原因时，可以在插件设置中开启「显示详细错误信息」。请在Issue页发表你的问题。

### 许可证

本项目采用 [MIT 许可证](LICENSE)，版权署名为 `2026 AphRO`。可以使用、修改和分发本插件（包括商用），分发时请保留版权及许可声明。
