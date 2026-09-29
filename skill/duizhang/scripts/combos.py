#!/usr/bin/env python3
"""
combos.py —— 关键列取值组合对照

把"导入成功的文件"和"导入失败的文件"在同一批查找列上的取值组合分别列出来，
再算出"只在这个文件出现、别的文件都没有"的组合差集，一眼看出两边写法上的差异。

用法示例：
    python combos.py good.xlsx::Sheet1::4 bad.xlsx::Sheet1::4 --cols A,B,C
    python combos.py a.csv b.csv --cols 1,2
"""

import argparse
import sys

import pandas as pd
from openpyxl.utils import column_index_from_string

sys.stdout.reconfigure(encoding="utf-8")


def parse_col_spec(spec: str) -> int:
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


def parse_file_spec(spec: str):
    """把 path 或 path::sheet 或 path::sheet::header_rows 拆开。"""
    parts = spec.split("::")
    path = parts[0]
    sheet = parts[1] if len(parts) > 1 and parts[1] != "" else None
    header_rows = int(parts[2]) if len(parts) > 2 and parts[2] != "" else 0
    return path, sheet, header_rows


def read_table(path: str, sheet, header_rows: int) -> pd.DataFrame:
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


def cell_str(v) -> str:
    return "" if pd.isna(v) else str(v)


def combo_counts(df: pd.DataFrame, col_idx):
    """返回 {组合元组: 出现行数} ，组合按原始值（不做任何归一化）取。"""
    sub = df.iloc[:, col_idx]
    counts = {}
    for row in sub.itertuples(index=False, name=None):
        key = tuple(cell_str(v) for v in row)
        counts[key] = counts.get(key, 0) + 1
    return counts


def main():
    parser = argparse.ArgumentParser(
        description="关键列取值组合对照：列出若干文件在同一批列上的取值组合与行数，并给出组合差集"
    )
    parser.add_argument(
        "files", nargs="+",
        help="若干个文件，形式为 path 或 path::sheet 或 path::sheet::header_rows（.xlsx 或 .csv）",
    )
    parser.add_argument("--cols", required=True, help="要对照的列，逗号分隔，可用列字母（A,B,C）或列序号（1,2,3）")
    args = parser.parse_args()

    try:
        col_idx = parse_col_list(args.cols)
    except ValueError as e:
        print(f"错误：--cols 解析失败：{e}", file=sys.stderr)
        sys.exit(1)

    specs = [parse_file_spec(s) for s in args.files]
    labels = [f"文件{i+1}" for i in range(len(specs))]

    all_counts = []
    for label, (path, sheet, header_rows) in zip(labels, specs):
        df = read_table(path, sheet, header_rows)
        max_idx = max(col_idx)
        if max_idx >= df.shape[1]:
            print(f"错误：{path} 只有 {df.shape[1]} 列，但 --cols 指定的列超出范围", file=sys.stderr)
            sys.exit(1)
        counts = combo_counts(df, col_idx)
        all_counts.append(counts)
        print(f"=== {label}：{path}（sheet={sheet}, header_rows={header_rows}） ===")
        print(f"数据行数：{len(df)}，唯一组合数：{len(counts)}")
        for combo in sorted(counts.keys()):
            print(f"  {combo} : {counts[combo]} 行")
        print()

    if len(specs) < 2:
        print("只有一个文件，无法计算差集。")
        return

    print("=== 组合差集（只在该文件出现，其他文件都没有） ===")
    any_diff = False
    for i, label in enumerate(labels):
        others = set()
        for j, counts in enumerate(all_counts):
            if j != i:
                others |= set(counts.keys())
        only_here = set(all_counts[i].keys()) - others
        if only_here:
            any_diff = True
            print(f"[只在 {label} 出现] 共 {len(only_here)} 个组合：")
            for combo in sorted(only_here):
                print(f"  {combo}")
    if not any_diff:
        print("没有发现只属于某一个文件的组合，所有文件的组合取值完全一致。")


if __name__ == "__main__":
    main()
