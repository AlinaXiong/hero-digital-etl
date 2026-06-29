# -*- coding: utf-8 -*-
"""把单列里用分隔符(;；、，, 及空白)合并的原泛微项目编码平铺成一列并去重。

输入:一个只有一列"原泛微项目编码"的 Excel,部分单元格里塞了多个编码,
     用 ; ; 、 , 等分隔符连接。
输出:同目录下生成 *_平铺去重.xlsx,第一列每行一个编码,保持首次出现顺序去重。

用法:
    python etl/util/flatten_project_codes.py "C:\\path\\数据清洗涉及泛微项目编码_0629.xlsx"
    # 不传参数时默认处理下面 DEFAULT_INPUT
"""
import re
import sys
from pathlib import Path

import pandas as pd

DEFAULT_INPUT = Path(__file__).resolve().parents[2] / 'resources' / 'source' / 'other_cleaned_data' / '数据清洗涉及泛微项目编码_0629.xlsx'

# 分隔符:中英文分号、顿号、中英文逗号、斜杠、竖线及各类空白
SEP_RE = re.compile(r'[;；、，,/|\s]+')


def flatten(input_path: Path) -> Path:
    df = pd.read_excel(input_path, dtype=str)
    if df.shape[1] < 1:
        raise SystemExit('输入文件没有数据列')

    col = df.columns[0]
    codes = []
    for cell in df[col]:
        if cell is None or (isinstance(cell, float) and pd.isna(cell)):
            continue
        for part in SEP_RE.split(str(cell)):
            part = part.strip()
            if part:
                codes.append(part)

    # 保持首次出现顺序去重
    seen = set()
    unique = []
    for c in codes:
        if c not in seen:
            seen.add(c)
            unique.append(c)

    out_df = pd.DataFrame({col: unique})
    out_path = input_path.with_name(input_path.stem + '_平铺去重.xlsx')
    out_df.to_excel(out_path, index=False)

    print(f'原始单元格行数(去表头): {len(df)}')
    print(f'拆分后总编码数        : {len(codes)}')
    print(f'去重后编码数          : {len(unique)}')
    print(f'输出文件              : {out_path}')
    return out_path


if __name__ == '__main__':
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_INPUT
    if not src.exists():
        raise SystemExit(f'找不到文件: {src}')
    flatten(src)
