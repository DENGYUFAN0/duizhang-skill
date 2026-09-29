#!/usr/bin/env python3
"""
dupcount.py —— 多口径查重计数

把同一批关键列，分别用几种不同的"归一化口径"（精确 / 去空白 / 忽略大小写 /
全半角归一 ……）计算重复行数，用来反推黑盒导入平台内部到底是按哪种口径判重、
按哪种策略（先到先得 / 同键全拒）处理重复数据。

用法示例：
    python dupcount.py data.xlsx --sheet Sheet1 --header-rows 4 --keys A,B,C,D,F --target 2013
    python dupcount.py data.csv --header-rows 1 --keys 1,2,3
"""

import argparse
import re
import sys
import unicodedata

import pandas as pd
from openpyxl.utils import column_index_from_string

sys.stdout.reconfigure(encoding="utf-8")

# 常见零宽字符：零宽空格、零宽连接符/非连接符、BOM、零宽单词连接符
ZERO_WIDTH_CHARS = ["​", "‌", "‍", "﻿", "⁠"]


def parse_col_spec(spec: str) -> int:
    """把一个列标识（字母如 A/AA，或数字序号如 1/27）转成 0-based 列索引。"""
    spec = spec.strip()
    if not spec:
        raise ValueError("列标识不能为空")
    if spec.isdigit():
        n = int(spec)
        if n < 1:
            raise ValueError(f"列序号必须从 1 开始，收到：{spec}")
        return n - 1
    return column_index_from_string(spec.upper()) - 1


def parse_col_list(spec: str):
    return [parse_col_spec(x) for x in spec.split(",") if x.strip()]


def read_table(path: str, sheet, header_rows: int) -> pd.DataFrame:
    """读取 xlsx 或 csv，跳过表头行，所有单元格按字符串读入，并丢弃全空行。"""
    lower = path.lower()
    if lower.endswith(".csv"):
        df = pd.read_csv(path, header=None, skiprows=header_rows, dtype=str)
    else:
        sheet_name = sheet if sheet is not None else 0
        df = pd.read_excel(
            path, sheet_name=sheet_name, header=None, skiprows=header_rows,
            dtype=str, engine="openpyxl",
        )
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def remove_zero_width(s: str) -> str:
    for ch in ZERO_WIDTH_CHARS:
        s = s.replace(ch, "")
    return s


def cipher_exact(v) -> str:
    if pd.isna(v):
        return ""
    return str(v)


def cipher_strip(v) -> str:
    return cipher_exact(v).strip()


def cipher_nospace(v) -> str:
    return re.sub(r"\s+", "", cipher_exact(v))


def cipher_lower(v) -> str:
    return cipher_exact(v).casefold()


def cipher_lower_nospace(v) -> str:
    return re.sub(r"\s+", "", cipher_exact(v).casefold())


def cipher_nfkc(v) -> str:
    s = unicodedata.normalize("NFKC", cipher_exact(v))
    s = remove_zero_width(s)
    s = re.sub(r"\s+", "", s)
    return s.casefold()


CIPHERS = [
    ("精确", cipher_exact),
    ("去首尾空白", cipher_strip),
    ("去所有空白", cipher_nospace),
    ("忽略大小写", cipher_lower),
    ("忽略大小写+去所有空白", cipher_lower_nospace),
    ("NFKC全半角归一+忽略大小写+去所有空白+去零宽字符", cipher_nfkc),
]


def build_composite(df: pd.DataFrame, key_idx, fn) -> pd.Series:
    """按给定口径函数把多个关键列拼成一个复合键（列向量化，避免逐行 apply）。"""
    cols = [df.iloc[:, i].map(fn) for i in key_idx]
    composite = cols[0]
    for c in cols[1:]:
        composite = composite + "\x1f" + c
    return composite


def main():
    parser = argparse.ArgumentParser(
        description="多口径查重计数：用不同归一化口径统计关键列组合的重复行数，反推黑盒平台的判重规则"
    )
    parser.add_argument("file", help="数据文件路径（.xlsx 或 .csv）")
    parser.add_argument("--sheet", default=None, help="工作表名（xlsx 用，不填则用第一个工作表）")
    parser.add_argument("--header-rows", type=int, default=0, help="数据前的表头行数")
    parser.add_argument("--keys", required=True, help="关键列，逗号分隔，可用列字母（A,B,C）或列序号（1,2,3）")
    parser.add_argument("--target", type=int, default=None, help="平台报的失败/多出条数，用来反推匹配的口径")
    args = parser.parse_args()

    try:
        key_idx = parse_col_list(args.keys)
    except ValueError as e:
        print(f"错误：--keys 解析失败：{e}", file=sys.stderr)
        sys.exit(1)

    if not key_idx:
        print("错误：--keys 不能为空", file=sys.stderr)
        sys.exit(1)

    df = read_table(args.file, args.sheet, args.header_rows)

    max_idx = max(key_idx)
    if max_idx >= df.shape[1]:
        print(f"错误：数据只有 {df.shape[1]} 列，但 --keys 指定的列超出范围", file=sys.stderr)
        sys.exit(1)

    print(f"文件：{args.file}")
    if args.sheet is not None:
        print(f"工作表：{args.sheet}")
    print(f"表头行数：{args.header_rows}")
    print(f"关键列（0-based 列索引）：{key_idx}")
    print(f"数据行数（跳过空行后）：{len(df)}")
    print()
    print("口径 | 先到先得(duplicated，保留首次出现) | 同键全拒(keep=False，全部judge失败)")

    results = []
    for name, fn in CIPHERS:
        composite = build_composite(df, key_idx, fn)
        dup_first = int(composite.duplicated(keep="first").sum())
        dup_all = int(composite.duplicated(keep=False).sum())
        results.append((name, dup_first, dup_all))
        print(f"[{name}] 先到先得={dup_first}  同键全拒={dup_all}")

    if args.target is not None:
        print()
        print(f"目标数（--target）：{args.target}")
        hits = []
        for name, dup_first, dup_all in results:
            if dup_first == args.target:
                hits.append((name, "先到先得（保留首次出现，其余判失败）"))
            if dup_all == args.target:
                hits.append((name, "同键全拒（涉及重复的所有行都判失败）"))
        if hits:
            for name, strategy in hits:
                print(f"命中：这说明平台很可能按【{name}】口径判重、【{strategy}】策略处理")
        else:
            print("没有口径恰好命中目标数，可能目标数由其他因素（跨列组合、其他判重列等）导致")


if __name__ == "__main__":
    main()
