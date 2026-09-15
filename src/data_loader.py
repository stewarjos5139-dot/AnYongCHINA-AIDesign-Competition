"""数据加载模块 —— 健壮读取 A / B 两系统 Excel 原始数据。

设计要点
--------
1. **无损读取**：名称列原样保留空格、全角括号、不可见字符，绝不在加载阶段做任何
   清洗，保证后续"原始字符串 vs 清洗后字符串"可对照。
2. **表头自动探测**：不假设表头一定在第 1 行，扫描前 N 行取命中期望字段最多的一行。
3. **软校验**：字段缺失只告警不中断；Sheet 缺失 / 文件缺失才抛错，并给出可选清单。
"""

from __future__ import annotations

import unicodedata
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

from . import config

# 常见"隐形字符"白名单：用于控制台诊断打印
INVISIBLE_CHARS: dict[str, str] = {
    " ": "NBSP 不换行空格",
    "　": "IDEOGRAPHIC SPACE 全角空格",
    "​": "ZERO WIDTH SPACE 零宽空格",
    "﻿": "BOM / ZERO WIDTH NO-BREAK SPACE",
    "\t": "TAB 制表符",
    "\n": "LF 换行",
    "\r": "CR 回车",
    " ": "LINE SEPARATOR",
}


class DataLoadError(RuntimeError):
    """数据文件缺失 / Sheet 缺失 / 结构不可用。"""


# --------------------------------------------------------------------------- #
#  底层工具
# --------------------------------------------------------------------------- #
def _norm_header(value: Any) -> str:
    """表头归一化：去掉首尾空白（含全角空格 / NBSP）。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value)
    return text.strip().strip("　 ﻿")


def list_sheets(path: Path | str) -> list[str]:
    """列出工作簿全部 Sheet 名（不依赖 openpyxl 全量加载）。"""
    path = Path(path)
    _ensure_file(path)
    with pd.ExcelFile(path, engine="openpyxl") as xls:
        return list(xls.sheet_names)


def _ensure_file(path: Path) -> None:
    if not path.exists():
        raise DataLoadError(f"文件不存在：{path}")
    if not path.is_file():
        raise DataLoadError(f"路径不是文件：{path}")


def _detect_header_row(
    path: Path, sheet: str, expected_cols: Iterable[str], max_scan: int = 10
) -> tuple[int, int]:
    """在前 max_scan 行中找出与期望字段命中数最多的行，返回 (行号, 命中数)。"""
    probe = pd.read_excel(
        path, sheet_name=sheet, header=None, nrows=max_scan, dtype=object
    )
    expected = set(expected_cols)
    best_row, best_hits = 0, -1
    for idx in range(len(probe)):
        cells = {
            _norm_header(v)
            for v in probe.iloc[idx].tolist()
            if not (v is None or (isinstance(v, float) and pd.isna(v)))
        }
        hits = len(expected & cells)
        if hits > best_hits:
            best_row, best_hits = idx, hits
    return best_row, max(best_hits, 0)


# --------------------------------------------------------------------------- #
#  主读取逻辑
# --------------------------------------------------------------------------- #
def load_table(
    path: Path | str,
    sheet: str,
    expected_cols: Iterable[str] = (),
    name_col: str | None = None,
    label: str = "表",
) -> pd.DataFrame:
    """读取单个 Sheet 为 DataFrame。

    参数
    ----
    path / sheet    : 文件与 Sheet 名
    expected_cols   : 期望字段（用于表头探测 + 缺失告警）
    name_col        : 核心对比列，强制转为 pandas ``string``  dtype 并原样保留空白
    label           : 报错信息中的表名标识
    """
    path = Path(path)
    _ensure_file(path)

    available = list_sheets(path)
    if sheet not in available:
        raise DataLoadError(
            f"[{label}] Sheet 不存在：{sheet!r}\n"
            f"    文件：{path}\n"
            f"    可选 Sheet：{available}"
        )

    header_row, hits = _detect_header_row(path, sheet, expected_cols)
    df = pd.read_excel(path, sheet_name=sheet, header=header_row, engine="openpyxl")
    df.columns = [_norm_header(c) for c in df.columns]

    # 丢掉表头探测产生的空列 / 空行
    df = df.loc[:, [c for c in df.columns if c != ""]]
    df = df.dropna(how="all").reset_index(drop=True)

    # 名称列：强制字符串 dtype，保留首尾空格与不可见字符
    if name_col:
        if name_col not in df.columns:
            raise DataLoadError(
                f"[{label}] 核心对比列缺失：{name_col!r}\n"
                f"    实际列名：{list(df.columns)}"
            )
        df[name_col] = df[name_col].astype("string")

    df.attrs.update(
        {
            "label": label,
            "path": str(path),
            "sheet": sheet,
            "header_row": header_row,
            "header_hits": hits,
            "missing_cols": [c for c in expected_cols if c not in df.columns],
            "extra_cols": [c for c in df.columns if c not in set(expected_cols)],
        }
    )
    return df


def load_a_system(path: Path | str | None = None, sheet: str | None = None) -> pd.DataFrame:
    """读取 A 系统（企业 ERP 客户交易明细）。"""
    return load_table(
        path or config.A_FILE,
        sheet or config.A_SHEET,
        expected_cols=config.EXPECTED_A_COLS,
        name_col=config.A_NAME_COL,
        label="A系统",
    )


def load_b_system(path: Path | str | None = None, sheet: str | None = None) -> pd.DataFrame:
    """读取 B 系统（银行流水明细）。"""
    return load_table(
        path or config.B_FILE,
        sheet or config.B_SHEET,
        expected_cols=config.EXPECTED_B_COLS,
        name_col=config.B_NAME_COL,
        label="B系统",
    )


def load_all() -> tuple[pd.DataFrame, pd.DataFrame]:
    """一键读取 A、B 两表，返回 ``(df_a, df_b)``。"""
    return load_a_system(), load_b_system()


# --------------------------------------------------------------------------- #
#  原始名称提取 / Unicode 诊断（供第二阶段清洗做对照）
# --------------------------------------------------------------------------- #
def raw_names(df: pd.DataFrame, name_col: str) -> list[str]:
    """按行序取出名称列的原始字符串（NA → 空串），不做任何 strip。"""
    return ["" if pd.isna(v) else str(v) for v in df[name_col].tolist()]


def unicode_report(text: str) -> dict[str, Any]:
    """给单条原始字符串做编码体检，供控制台对照打印。"""
    invisibles = [
        {"pos": i, "char": repr(ch), "meaning": INVISIBLE_CHARS[ch]}
        for i, ch in enumerate(text)
        if ch in INVISIBLE_CHARS
    ]
    return {
        "text": text,
        "repr": repr(text),
        "length": len(text),
        "codepoints": " ".join(f"U+{ord(ch):04X}" for ch in text),
        "has_leading_or_trailing_space": text != text.strip(),
        "invisibles": invisibles,
        "non_ascii": any(ord(ch) > 127 for ch in text),
        "categories": "".join(
            unicodedata.category(ch)[0] for ch in text
        ),  # L=字母 N=数字 P=标点 Z=分隔 S=符号
    }
