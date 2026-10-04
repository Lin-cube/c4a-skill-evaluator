# 文档转换器教学说明

## 上手 5 分钟

1. 安装依赖：`pip install python-docx`（纯 md 场景可跳过）
2. 准备一个含 `.md` 文件的文件夹
3. 运行 `python doc_converter.py 你的文件夹`
4. 到 `converted/` 查看结果

## 常见坑

- **坑 1**：路径含空格要加引号：`python doc_converter.py "我的 文件夹"`
- **坑 2**：GBK 编码的旧文件可能乱码，脚本已做 UTF-8 优先、errors=ignore 兜底
- **坑 3**：只处理 `.md`，docx 需安装 python-docx

## 步骤拆解

1. 扫描输入目录
2. 逐文件规范化换行
3. 写入输出目录
4. 生成转换清单 CSV

## 注意事项

- 输出会覆盖同名文件，注意备份
- 转换清单用 utf-8-sig 编码，Excel 直接打开不乱码
