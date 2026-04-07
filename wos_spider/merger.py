"""CSV/XLSX 数据合并工具。

读取目录下所有数据文件，按指定列分组求和并排序。
修复原项目 bug：原 merge_xlsx.py 只合并了最后一个文件。
"""
import os
import glob
import pandas as pd


def merge_files(directory, group_col, value_col, output_path=None):
    """合并目录下所有 CSV/XLSX 文件。

    Args:
        directory: 包含数据文件的目录路径
        group_col: 分组列名（如 "country"）
        value_col: 求和列名（如 "count"）
        output_path: 可选，输出文件路径。不传则只返回 DataFrame

    Returns:
        pd.DataFrame: 合并后的数据，按 value_col 降序排列
    """
    csv_files = glob.glob(os.path.join(directory, "*.csv"))
    xlsx_files = glob.glob(os.path.join(directory, "*.xlsx"))
    all_files = csv_files + xlsx_files

    if not all_files:
        return pd.DataFrame(columns=[group_col, value_col])

    frames = []
    for f in all_files:
        try:
            if f.endswith(".csv"):
                df = pd.read_csv(f)
            else:
                df = pd.read_excel(f)
            frames.append(df)
        except Exception as e:
            print(f"Warning: 跳过文件 {f}: {e}")
            continue

    if not frames:
        return pd.DataFrame(columns=[group_col, value_col])

    merged = pd.concat(frames, ignore_index=True)

    # 统一列名：取第一列和第二列
    if group_col not in merged.columns or value_col not in merged.columns:
        cols = merged.columns.tolist()
        if len(cols) >= 2:
            merged = merged.rename(columns={cols[0]: group_col, cols[1]: value_col})

    result = merged.groupby(group_col)[value_col].sum().reset_index()
    result = result.sort_values(by=value_col, ascending=False).reset_index(drop=True)

    if output_path:
        if output_path.endswith(".xlsx"):
            result.to_excel(output_path, index=False)
        else:
            result.to_csv(output_path, index=False)

    return result
