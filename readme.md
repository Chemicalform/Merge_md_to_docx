# Markdown 批量合并为 DOCX 工具

`merge_md_to_docx.py` 是一个自动化脚本，用于将目录树中的 Markdown 文件按文件夹批量合并为格式化的 Word 文档（`.docx`）。

---

## ✨ 功能特性

- **自动递归扫描** — 遍历根目录及所有子目录，找出每个包含 `.md` 文件的文件夹
- **按文件夹生成文档** — 每个文件夹独立输出一个 `.docx`，文件名与文件夹同名
- **章节化组织** — 每个 `.md` 文件作为一个章节，以文件名作为章节标题，章节间自动分页
- **Pandoc 格式转换** — 使用 pandoc 进行 Markdown → DOCX 转换，完整保留标题、列表、表格、代码块等格式
- **多线程并行处理** — 多个文件夹同时转换，大批量文件时显著提升速度
- **编码自适应** — UTF-8 优先，自动回退 GBK / GB2312 / Latin-1
- **排版美化** — 正文两端对齐，页边距可自定义

---

## 📦 安装依赖

### 1. Python 包

```bash
pip install python-docx
```

### 2. Pandoc

Pandoc 是**必须依赖**，请根据系统选择安装方式：

| 系统 | 安装命令 |
|------|----------|
| Windows | `winget install --id JohnMacFarlane.Pandoc` 或 [下载安装包](https://pandoc.org/installing.html) |
| macOS | `brew install pandoc` |
| Ubuntu/Debian | `sudo apt install pandoc` |

安装后运行 `pandoc --version` 验证。

---

## 🚀 使用方法

### 基本用法

将脚本放在目标目录下，直接运行：

```bash
python merge_md_to_docx.py
```

脚本会扫描所在目录及所有子目录，为每个含 `.md` 文件的文件夹生成一个同名 `.docx`。

### 命令行参数

```bash
python merge_md_to_docx.py [选项]
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `--dir` | 路径 | 脚本所在目录 | 指定要扫描的根目录 |
| `--margin` | 浮点数 | `1.27` | 页边距，单位 cm |
| `--encoding` | 字符串 | `utf-8` | 主编码（失败时自动回退 GBK 等） |
| `--workers` | 整数 | CPU 核心数 | 并行线程数 |

### 示例

```bash
# 处理指定目录
python merge_md_to_docx.py --dir "C:\Users\docs\notes"

# 自定义页边距为 2cm，使用 4 线程
python merge_md_to_docx.py --margin 2.0 --workers 4

# 单线程顺序处理
python merge_md_to_docx.py --workers 1
```

---

## 📂 输入输出示例

### 输入目录结构

```
root/
├── merge_md_to_docx.py
├── Google/
│   ├── gemini.md
│   ├── chrome.md
│   └── bard.md
├── OpenAI/
│   ├── gpt4.md
│   └── chatgpt.md
└── Misc/
    └── notes.md
```

### 运行输出

```
找到 3 个文件夹，使用 3 个线程并行处理…

[1/3] ✓ Google\Google.docx（3 个章节）
[2/3] ✓ OpenAI\OpenAI.docx（2 个章节）
[3/3] ✓ Misc\Misc.docx（1 个章节）

完成！共生成 3 个 DOCX 文件。
```

### 生成的文件

```
root/
├── Google/
│   ├── gemini.md
│   ├── chrome.md
│   ├── bard.md
│   └── Google.docx        ← 新生成
├── OpenAI/
│   ├── gpt4.md
│   ├── chatgpt.md
│   └── OpenAI.docx        ← 新生成
└── Misc/
    ├── notes.md
    └── Misc.docx           ← 新生成
```

---

## 📄 生成文档格式说明

每个 `.docx` 文件的结构如下：

```
┌──────────────────────────┐
│  章节标题 (文件名)          │  ← 一级标题，来自 .md 文件名（去掉扩展名）
│                            │
│  Markdown 正文内容          │  ← pandoc 转换，保留原始格式
│  - 标题层级 (h1-h6)        │
│  - 列表（有序/无序）         │
│  - 代码块（等宽字体）        │
│  - 表格                     │
│  - 粗体 / 斜体              │
│                            │
│  ════════ 分页 ════════     │  ← 自动分页
│                            │
│  下一章节标题                │
│  ...                       │
└──────────────────────────┘
```

- **排版**：正文两端对齐（Justify）
- **页边距**：默认 1.27cm（可通过 `--margin` 调整）
- **章节顺序**：按文件名字母序排列

---

## ⚙️ 工作原理

```
1. 扫描根目录，收集所有含 .md 文件的子目录
2. 检测 pandoc 是否可用（未找到则终止并提示安装）
3. 为每个文件夹创建一个线程任务：
   a. 收集文件夹内的 .md 文件（按文件名排序）
   b. 创建空白 Word 文档，设置页边距
   c. 遍历每个 .md 文件：
      - 添加章节标题（文件名）
      - 调用 pandoc 转换为临时 .docx
      - 将临时文档的段落复制到主文档（保留格式）
      - 如果 pandoc 转换失败，回退到内置简易解析器
      - 插入分页符
   d. 保存 .docx 到该文件夹下
4. 输出处理结果统计
```

---

## 🧩 内置简易解析器（Fallback）

当 pandoc 处理某个 `.md` 文件失败时，脚本会自动回退到内置的简易 Markdown 解析器，支持：

| 格式 | 支持情况 |
|------|----------|
| 标题 `# ~ ######` | ✅ 1-6 级 |
| 无序列表 `- * +` | ✅ |
| 有序列表 `1. 2.` | ✅ |
| 围栏代码块 ` ``` ` | ✅ Consolas 等宽字体 |
| 粗体 `**text**` | ✅ 去除标记 |
| 斜体 `*text*` | ✅ 去除标记 |
| 行内代码 `` `code` `` | ✅ 去除标记 |
| 水平线 `---` | ✅ 忽略 |

---

## ❓ 常见问题

**Q: 运行后提示「错误: 未找到 pandoc」**
A: 请先安装 pandoc，参见上方「安装依赖」部分。安装后确保 `pandoc` 命令在系统 PATH 中。

**Q: 生成的文档中中文显示异常**
A: 脚本默认使用 UTF-8 编码，如果源文件是 GBK 编码会自动回退。如仍有问题，请检查源 `.md` 文件的编码格式。

**Q: 如何只处理某个子目录？**
A: 使用 `--dir` 参数指定目录，例如：`python merge_md_to_docx.py --dir ./Google`

**Q: 文件名中有特殊字符怎么办？**
A: 脚本使用 `pathlib` 处理路径，支持大多数特殊字符。但建议避免文件名包含 `/` `\` `?` `*` 等系统保留字符。

**Q: 如何控制并行度？**
A: 使用 `--workers` 参数，例如 `--workers 4` 使用 4 个线程。设为 `--workers 1` 则顺序执行。
