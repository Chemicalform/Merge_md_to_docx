
merge_md_to_docx.py
 程序介绍
🎯 功能概述
这是一个 Markdown 批量合并为 Word 文档的工具。它会自动扫描指定目录（及所有子目录），对每个包含 
.md
 文件的文件夹，将其中的所有 Markdown 文件合并成一个 .docx 文档。

📂 工作流程
mermaid
graph LR
    A[扫描根目录] --> B[找出所有含 .md 的文件夹]
    B --> C[多线程并行处理每个文件夹]
    C --> D["每个 .md → 一个章节"]
    D --> E[合并输出为 文件夹名.docx]
例如，目录结构为：

root/
├── Google/
│   ├── gemini.md
│   ├── chrome.md
│   └── bard.md
├── OpenAI/
│   ├── gpt4.md
│   └── chatgpt.md
运行后生成：

Google/Google.docx（含 3 个章节，每章节对应一个 md 文件）
OpenAI/OpenAI.docx（含 2 个章节）
✨ 核心特性
特性	说明
Markdown 解析	支持 1-6 级标题、有序/无序列表、围栏代码块、粗体/斜体
章节分页	每个 
.md
 文件独占一章，章节间自动插入分页符
两端对齐	正文段落统一设为两端对齐（Justify）
pandoc 增强	检测到系统安装了 pandoc 时自动使用，保留更丰富的格式
多线程并行	使用 ThreadPoolExecutor 并行处理多个文件夹，加速大批量转换
编码自适应	UTF-8 优先，失败自动回退 GBK/GB2312/Latin-1
可定制参数	通过命令行指定目录、页边距、线程数等
🔧 命令行参数
bash
python merge_md_to_docx.py [选项]
参数	默认值	说明
--dir	脚本所在目录	指定要扫描的根目录
--margin	1.27	页边距，单位 cm
--encoding	utf-8	主编码（失败自动回退）
--workers	CPU 核心数	并行线程数
📦 依赖
必需：pip install python-docx
可选：安装 pandoc 可获得更完整的 Markdown 格式保留（表格、脚注等）
