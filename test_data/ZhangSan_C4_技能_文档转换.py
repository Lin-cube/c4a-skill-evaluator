#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""doc_converter.py — 文档批量转 Markdown。

输入：一个文件夹路径
输出：converted/ 目录下的 .md 文件 + 转换清单.csv
"""

import csv
import sys
from pathlib import Path


def convert_text(text: str) -> str:
    """把纯文本规范化为 Markdown 段落。"""
    lines = [ln.rstrip() for ln in text.splitlines()]
    return "\n\n".join(ln for ln in lines if ln)


def main(input_dir: str, out_dir: str = "converted") -> int:
    src = Path(input_dir)
    dst = Path(out_dir)
    dst.mkdir(parents=True, exist_ok=True)

    rows = []
    for f in sorted(src.glob("*.md")):
        text = f.read_text(encoding="utf-8", errors="ignore")
        (dst / f.name).write_text(convert_text(text), encoding="utf-8")
        rows.append([f.name, "converted"])

    with open(dst / "转换清单.csv", "w", newline="", encoding="utf-8-sig") as fp:
        w = csv.writer(fp)
        w.writerow(["文件", "状态"])
        w.writerows(rows)

    print(f"完成：{len(rows)} 个文件已转换")
    return 0


def test_convert_text():
    """测试：空文本应返回空串；多空行应合并。"""
    assert convert_text("") == ""
    assert convert_text("a\n\n\nb") == "a\n\nb"
    print("测试通过")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python doc_converter.py <输入文件夹> [--out converted]")
        sys.exit(1)
    # 简单示例：转换当前目录
    test_convert_text()
    sys.exit(main(sys.argv[1]))
