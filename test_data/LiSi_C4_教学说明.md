# 格式转换教学说明

## 上手

1. `pip install openpyxl`
2. `python format_convert.py 数据.csv`

## 步骤

- csv→xlsx：脚本读 csv 每行，写入 Excel 工作表
- xlsx→csv：读工作表，拼成逗号分隔文本

## 常见坑

- 没装 openpyxl 会报 ModuleNotFoundError，先 pip install
- csv 含逗号的内容会被切错，暂未处理引号转义

## 注意事项

- 需要 Python 3.8+
