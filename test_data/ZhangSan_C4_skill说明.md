# 文档转换器技能说明（Doc Converter）

## 解决什么问题

把 Word / Markdown 文档批量转换成规范化 Markdown，并自动提取标题层级，方便知识库导入。

## 使用场景

- 团队整理会议纪要、周报时批量转格式
- 把历史 docx 文档统一归档为 Markdown
- 为 AI 知识库准备可检索的文本语料

## 输入与输出

- **输入**：一个本地文件夹路径，内含若干 `.docx` 或 `.md` 文件
- **输出**：一个 `converted/` 目录，包含转换后的 `.md` 文件 + 一份 `转换清单.csv`

一句话：**输入一个文件夹路径，输出规范化 Markdown 文件与转换清单 CSV。**

## 安装与运行

1. 环境要求：Python 3.9+，无需额外第三方库（仅用标准库 + python-docx，可选）
2. 安装依赖：
   ```
   pip install python-docx
   ```
3. 运行：
   ```
   python doc_converter.py <输入文件夹> --out converted
   ```

## 依赖与兼容

- 依赖：`python-docx`（仅 docx 转换需要，纯 md 转换零依赖）
- 兼容：Windows / macOS / Linux 均可运行

## 边界情况

- 空文件夹：输出空清单并提示
- 非 docx/md 文件：跳过并在清单中标记
- 中文文件名：UTF-8 处理，GBK 回退
