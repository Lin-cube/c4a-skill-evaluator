# 我的脚本
import os

path = r"C:\Users\WangWu\Desktop\我的文件\data.txt"
with open(path, encoding="gbk") as f:
    for line in f:
        print(line.strip())
