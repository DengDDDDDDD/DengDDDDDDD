#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd


CHANNEL_RE = re.compile(r"(?i)^(?:wifi\s*)?(2\.4g|5g|6g)[-_ ]*\s*(\d{1,3})\s*$")


def _norm_str(v: Any) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()


def _is_rx_token(v: Any) -> bool:
    s = _norm_str(v).upper()
    return s == "RX"


def _is_zero_token(v: Any) -> bool:
    s = _norm_str(v)
    return ("归零" in s) or (s.strip().lower() in {"zero", "zero point", "zeropoint"})


def _is_channel_token(v: Any) -> bool:
    s = _norm_str(v)
    if not s:
        return False
    return CHANNEL_RE.match(s) is not None


def _parse_channel(v: Any) -> Optional[str]:
    s = _norm_str(v)
    m = CHANNEL_RE.match(s)
    if not m:
        return None
    band = m.group(1).upper().replace("2.4G", "2.4G")
    ch = int(m.group(2))
    return f"{band}-{ch}"


def _coerce_number(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)):
        return float(v)
    s = _norm_str(v)
    if not s:
        return None
    s = s.replace("，", ",").replace("。", ".")
    s = re.sub(r"[^\d\.\-\+eE]", "", s)
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


def _best_column_by_name(columns: Sequence[Any], keywords: Sequence[str]) -> Optional[Any]:
    for c in columns:
        name = _norm_str(c).lower()
        for kw in keywords:
            if kw in name:
                return c
    return None


@dataclass(frozen=True)
class ExtractResult:
    source: str
    mapping: Dict[str, float]  # channel -> zero point
    notes: List[str]


def extract_rx_zero_points(path: Path, sheet: str = "RVR") -> ExtractResult:
    notes: List[str] = []

    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except ValueError as e:
        raise SystemExit(f"找不到 sheet `{sheet}`：{e}") from e

    # 1) Structured path: columns contain 信道/Channel + 归零点/Zero + RX marker
    channel_col = _best_column_by_name(df.columns, ["信道", "频道", "channel", "ch"])
    zero_col = _best_column_by_name(df.columns, ["归零", "zero"])
    rx_col = _best_column_by_name(df.columns, ["rx/tx", "rxtx", "方向", "dir", "收发", "mode", "txrx"])

    if channel_col is not None and zero_col is not None:
        dfx = df.copy()
        if rx_col is not None:
            mask = dfx[rx_col].map(_is_rx_token)
            dfx = dfx.loc[mask]
            notes.append(f"结构化解析：使用列 `{rx_col}` 过滤 RX 行")
        else:
            # If no explicit RX column, try find rows containing RX anywhere.
            mask_any = dfx.apply(lambda r: any(_is_rx_token(v) for v in r.values), axis=1)
            dfx = dfx.loc[mask_any]
            notes.append("结构化解析：未找到 RX 列，改为整行扫描 RX")

        mapping: Dict[str, float] = {}
        for _, row in dfx.iterrows():
            ch = _parse_channel(row.get(channel_col))
            zp = _coerce_number(row.get(zero_col))
            if ch is None or zp is None:
                continue
            mapping[ch] = zp

        if mapping:
            notes.append(f"结构化解析：从 `{channel_col}` + `{zero_col}` 提取到 {len(mapping)} 个信道归零点")
            return ExtractResult(source="structured", mapping=mapping, notes=notes)

        notes.append("结构化解析：未能提取到有效信道/归零点（可能列名匹配但内容不符合预期）")

    # 2) Heuristic path: scan raw grid for “归零点” row(s) within RX section(s)
    raw = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", header=None, dtype=object)
    raw = raw.replace({np.nan: None})

    # locate candidate header rows that contain many channel tokens
    header_rows: List[Tuple[int, List[Tuple[int, str]]]] = []
    for r in range(raw.shape[0]):
        chans: List[Tuple[int, str]] = []
        for c in range(raw.shape[1]):
            ch = _parse_channel(raw.iat[r, c])
            if ch is not None:
                chans.append((c, ch))
        if len(chans) >= 2:
            header_rows.append((r, chans))

    zero_cells: List[Tuple[int, int]] = []
    for r in range(raw.shape[0]):
        for c in range(raw.shape[1]):
            if _is_zero_token(raw.iat[r, c]):
                zero_cells.append((r, c))

    if not zero_cells:
        notes.append("启发式解析：未找到包含“归零”的单元格")
        return ExtractResult(source="none", mapping={}, notes=notes)

    mapping: Dict[str, float] = {}

    def nearest_header_row(zr: int) -> Optional[Tuple[int, List[Tuple[int, str]]]]:
        best: Optional[Tuple[int, List[Tuple[int, str]]]] = None
        best_dist = 10**9
        for hr, chans in header_rows:
            if hr < zr and (zr - hr) < best_dist:
                best = (hr, chans)
                best_dist = zr - hr
        return best

    for zr, zc in zero_cells:
        # ensure this "归零点" belongs to RX (either same row has RX token, or left neighbors / above label includes RX)
        row_vals = [raw.iat[zr, c] for c in range(raw.shape[1])]
        row_has_rx = any(_is_rx_token(v) for v in row_vals)
        left_has_rx = any(_is_rx_token(raw.iat[zr, c]) for c in range(max(0, zc - 3), zc + 1))
        above_has_rx = any(
            _is_rx_token(raw.iat[r, zc]) for r in range(max(0, zr - 5), zr + 1)
        )
        if not (row_has_rx or left_has_rx or above_has_rx):
            continue

        hdr = nearest_header_row(zr)
        if hdr is None:
            continue
        hr, chans = hdr

        # use the same columns as channel headers; take numeric values from zero row
        extracted = 0
        for col_idx, ch in chans:
            val = _coerce_number(raw.iat[zr, col_idx])
            if val is None:
                continue
            mapping[ch] = val
            extracted += 1

        if extracted:
            notes.append(f"启发式解析：在第 {zr+1} 行找到 RX 归零点，并用第 {hr+1} 行作为信道表头（提取 {extracted} 个）")

    if not mapping:
        notes.append("启发式解析：找到了“归零点”，但未能在 RX 区域提取到任何信道数值")
        return ExtractResult(source="heuristic-no-data", mapping={}, notes=notes)

    notes.append(f"启发式解析：总计提取到 {len(mapping)} 个信道归零点")
    return ExtractResult(source="heuristic", mapping=mapping, notes=notes)


def describe_diff(diffs: Sequence[float]) -> Dict[str, Any]:
    arr = np.array([d for d in diffs if d is not None], dtype=float)
    if arr.size == 0:
        return {"count": 0}
    return {
        "count": int(arr.size),
        "mean": float(arr.mean()),
        "std": float(arr.std(ddof=1)) if arr.size >= 2 else 0.0,
        "min": float(arr.min()),
        "p25": float(np.percentile(arr, 25)),
        "median": float(np.percentile(arr, 50)),
        "p75": float(np.percentile(arr, 75)),
        "max": float(arr.max()),
    }


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="比较两个 Excel 的 RVR(RX) 归零点差异（按相同信道，如 5G-48）"
    )
    ap.add_argument("excel_a", type=Path, help="第一个 Excel 文件路径（A表）")
    ap.add_argument("excel_b", type=Path, help="第二个 Excel 文件路径（B表）")
    ap.add_argument("--sheet", default="RVR", help="sheet 名称（默认：RVR）")
    ap.add_argument("--out-csv", default="rvr_zero_point_diff.csv", help="输出明细 CSV 文件名")
    ap.add_argument(
        "--out-json",
        default="rvr_zero_point_diff_summary.json",
        help="输出统计 JSON 文件名",
    )
    args = ap.parse_args(argv)

    a = extract_rx_zero_points(args.excel_a, sheet=args.sheet)
    b = extract_rx_zero_points(args.excel_b, sheet=args.sheet)

    chans = sorted(set(a.mapping.keys()) & set(b.mapping.keys()))
    rows: List[Dict[str, Any]] = []
    diffs: List[float] = []
    for ch in chans:
        za = a.mapping[ch]
        zb = b.mapping[ch]
        d = za - zb
        diffs.append(d)
        rows.append(
            {
                "channel": ch,
                "zero_point_a": za,
                "zero_point_b": zb,
                "diff_a_minus_b": d,
            }
        )

    df_out = pd.DataFrame(rows).sort_values(["channel"])
    df_out.to_csv(args.out_csv, index=False, encoding="utf-8-sig")

    summary = {
        "file_a": str(args.excel_a),
        "file_b": str(args.excel_b),
        "sheet": args.sheet,
        "extract_a": {"source": a.source, "notes": a.notes, "count": len(a.mapping)},
        "extract_b": {"source": b.source, "notes": b.notes, "count": len(b.mapping)},
        "intersection_channel_count": len(chans),
        "diff_stats_a_minus_b": describe_diff(diffs),
    }
    Path(args.out_json).write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"\n明细已输出：{args.out_csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

