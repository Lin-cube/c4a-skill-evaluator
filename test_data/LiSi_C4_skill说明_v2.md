# 格式转换技能说明 v2

## 解决什么问题

把 Excel/CSV 格式互相转换，并支持表头清洗，方便数据入库前预处理。

## 使用场景

- 拿到的数据是 xlsx 但系统只吃 csv 时
- 表头有空格/大小写不统一需要规范化时

## 输入与输出

- **输入**：一个 `.csv` 或 `.xlsx` 文件路径
- **输出**：同目录下生成转换后的文件（csv→xlsx 或 xlsx→csv）

## 安装

```
pip install openpyxl
```

## 运行

```
python format_convert.py 数据.csv
```

## 注意事项

- 需要 Python 3.8+
- 转换后的表头会自动去空格、转小写
