#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 docs/2026欧洲FB手工形态汇总.xlsx 的「汇总」表导出最新人工确认的类型库，
仅用于网页快速判断，不再连数据库重算。

输出：static/summary_types.json
"""

import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple


def _find_sheet_xml_by_name(z: zipfile.ZipFile, sheet_name: str) -> str:
    """从 workbook.xml 中按 sheet name 定位 worksheet xml 路径（如 xl/worksheets/sheet1.xml）。"""
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    wb_root = ET.parse(z.open("xl/workbook.xml")).getroot()
    rels_root = ET.parse(z.open("xl/_rels/workbook.xml.rels")).getroot()

    rid_to_target: Dict[str, str] = {}
    for rel in rels_root.findall("rel:Relationship", ns):
        rid = rel.get("Id")
        target = rel.get("Target") or ""
        if rid:
            rid_to_target[rid] = target

    sheets_el = wb_root.find("main:sheets", ns)
    if sheets_el is None:
        raise ValueError("workbook.xml missing sheets")

    for sh in sheets_el.findall("main:sheet", ns):
        name = sh.get("name") or ""
        if name != sheet_name:
            continue
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or ""
        target = rid_to_target.get(rid, "")
        if not target:
            raise ValueError(f"sheet '{sheet_name}' has no relationship target")
        # target like "worksheets/sheet1.xml"
        return "xl/" + target.lstrip("/")

    raise ValueError(f"sheet not found by name: {sheet_name}")


def _load_summary_sheet(path: str) -> List[Dict[str, Any]]:
    """解析 2026欧洲FB手工形态汇总.xlsx 的「汇总」表，返回规则行。"""
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    with zipfile.ZipFile(path, "r") as z:
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # 读取 sharedStrings
        with z.open("xl/sharedStrings.xml") as f:
            ss_root = ET.parse(f).getroot()
        strings: List[str] = []
        for si in ss_root.findall(".//main:si", ns):
            texts = si.findall(".//main:t", ns)
            strings.append("".join(t.text or "" for t in texts) if texts else "")

        sheet_xml = _find_sheet_xml_by_name(z, "汇总")
        with z.open(sheet_xml) as f:
            sheet_root = ET.parse(f).getroot()

    rows_raw: Dict[int, Dict[str, str]] = defaultdict(dict)

    for row in sheet_root.findall(".//main:sheetData/main:row", ns):
        r_idx_attr = row.get("r")
        try:
            r_idx = int(r_idx_attr)
        except (TypeError, ValueError):
            continue
        for c in row.findall("main:c", ns):
            ref = c.get("r")  # 如 "A2"
            if not ref:
                continue
            v = c.find("main:v", ns)
            t = c.get("t")
            if v is not None and v.text is not None:
                if t == "s":
                    idx = int(v.text)
                    val = strings[idx] if 0 <= idx < len(strings) else ""
                else:
                    val = v.text
            else:
                val = ""
            m = re.match(r"([A-Z]+)\d+", ref)
            if not m:
                continue
            col = m.group(1)
            rows_raw[r_idx][col] = (val or "").strip()

    def is_group_header(s: str) -> bool:
        s = (s or "").strip()
        # 例如 "0/0", "0/0.25", "0.25/0", "0.5/0.25"
        return bool(re.match(r"^\d+(?:\.\d+)?/\d+(?:\.\d+)?$", s))

    rules: List[Dict[str, Any]] = []
    current_group = ""
    last_side = ""

    for r_idx in sorted(rows_raw.keys()):
        row = rows_raw[r_idx]
        a_val = row.get("A", "")
        if is_group_header(a_val):
            current_group = a_val
            last_side = ""
            continue

        if a_val in ("主", "客"):
            last_side = a_val
        elif not a_val and any(row.get(col, "") for col in ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K")) and last_side:
            a_val = last_side
            row["A"] = a_val
        else:
            continue
        if not current_group:
            continue

        # 忽略全空行（条件 + 预测 + 统计 + 标注提示）
        if not any(row.get(col, "") for col in ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P")):
            continue

        cols = {}
        for col in ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"):
            cols[col] = row.get(col, "").strip()

        rules.append(
            {
                "row_index": r_idx,
                "group": current_group,
                "cols": cols,
            }
        )

    return rules


_HEADER_NAME_MAP = {
    "B": "马会",
    "C": "水差",
    "D": "澳平",
    "E": "马主",
    "F": "马平",
    "G": "主差",
    "H": "平差",
    "I": "客差",
    "J": "澳平客差",
}

# 规则表列 -> 用户输入 A-R 列
_COL_TO_INPUT_COL = {
    "B": "G",  # 汇总「马会」 -> 数据 G 列（上水）
    "C": "I",  # 汇总「水差」      -> 数据 I 列（水差）
    "D": "L",  # 汇总「澳平」      -> 数据 L 列（澳）
    "E": "O",  # 汇总「马主」      -> 数据 O 列
    "F": "P",  # 汇总「马平」      -> 数据 P 列（马）
    "G": "U",  # 汇总「主差」      -> 数据 U 列（主差）
    "H": "V",  # 汇总「平差」      -> 数据 V 列（平差）
    "I": "W",  # 汇总「客差」      -> 数据 W 列（客差）
    "J": "X",  # 汇总「澳平客差」  -> 数据 X 列（澳平客差）
}


def _normalize_cond_text(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    s = (
        s.replace("（", "(")
        .replace("）", ")")
        .replace("＜", "<")
        .replace("＞", ">")
        .replace("≦", "<=")
        .replace("≧", ">=")
        .replace("≤", "<=")
        .replace("≥", ">=")
        .replace(" ", "")
    )
    # 用 "且" 作为 AND 分隔符
    s = s.replace("且", "&&")
    return s


def _parse_atom(atom: str) -> Dict[str, Any]:
    """解析一个原子条件字符串为 {op, value}。失败则抛异常。"""
    atom = atom.strip()
    if not atom:
        raise ValueError("empty atom")
    op = None
    val_str = None

    if atom.startswith(">="):
        op = ">="
        val_str = atom[2:]
    elif atom.startswith("<="):
        op = "<="
        val_str = atom[2:]
    elif atom.startswith(">"):
        op = ">"
        val_str = atom[1:]
    elif atom.startswith("<"):
        op = "<"
        val_str = atom[1:]
    else:
        # 纯数字，视为等号
        op = "="
        val_str = atom
    value = float(val_str)
    return {"op": op, "value": value}


def _parse_conditions_for_col(text: str) -> List[Dict[str, Any]]:
    """将单元格中的条件文本解析为原子条件列表。

    支持两种形式：
    1）用 "且" 连接的多个不等式：<3.36且>3.19
    2）用 "~" 表示区间：0.01~0.03  等价于 >=0.01 且 <=0.03
    """
    s = _normalize_cond_text(text)
    if not s:
        return []
    conds: List[Dict[str, Any]] = []

    # 先处理纯区间形式：形如 "a~b"（不含 且）
    if "&&" not in s and "~" in s:
        parts = s.split("~")
        if len(parts) == 2 and parts[0] and parts[1]:
            try:
                v1 = float(parts[0])
                v2 = float(parts[1])
                lo, hi = (v1, v2) if v1 <= v2 else (v2, v1)
                conds.append({"op": ">=", "value": lo})
                conds.append({"op": "<=", "value": hi})
                return conds
            except Exception:
                # 回退到普通解析
                pass

    # 普通形式：用 且/&& 连接的多个原子条件
    parts = s.split("&&")
    for p in parts:
        if not p:
            continue
        try:
            conds.append(_parse_atom(p))
        except Exception:
            # 解析失败就忽略该原子条件
            continue

    # 修正：当一个条件是 "=" 且与 ">" 或 "<" 配合时，说明原文缺少运算符
    # 例如 "0.04且＞-0.04" 应为 "<0.04且>-0.04"
    if len(conds) >= 2:
        eq_indices = [i for i, c in enumerate(conds) if c["op"] == "="]
        other_ops = [c["op"] for i, c in enumerate(conds) if c["op"] != "="]
        for idx in eq_indices:
            eq_val = conds[idx]["value"]
            has_gt = any(c["op"] in (">", ">=") and c["value"] < eq_val for c in conds if c is not conds[idx])
            has_lt = any(c["op"] in ("<", "<=") and c["value"] > eq_val for c in conds if c is not conds[idx])
            if has_gt:
                conds[idx]["op"] = "<"
            elif has_lt:
                conds[idx]["op"] = ">"

    return conds


def _build_feature_text(cols: Dict[str, str]) -> str:
    parts: List[str] = []
    for col in ("B", "C", "D", "E", "F", "G", "H", "I", "J"):
        v = cols.get(col, "")
        if not v:
            continue
        name = _HEADER_NAME_MAP.get(col, col)
        parts.append(f"{name}:{v}")
    feat = "，".join(parts)
    pred = cols.get("K", "")
    if pred:
        if feat:
            return f"{feat}；预测:{pred}"
        return f"预测:{pred}"
    return feat


def _load_base_table(path: str) -> List[Dict[str, Optional[float]]]:
    """Load base table rows with numeric columns and result (AB)."""
    with zipfile.ZipFile(path, "r") as z:
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with z.open("xl/sharedStrings.xml") as f:
            ss_root = ET.parse(f).getroot()
        strings: List[str] = []
        for si in ss_root.findall(".//main:si", ns):
            texts = si.findall(".//main:t", ns)
            strings.append("".join(t.text or "" for t in texts) if texts else "")

        sheet_xml = _find_sheet_xml_by_name(z, "基本表")
        with z.open(sheet_xml) as f:
            sheet_root = ET.parse(f).getroot()

    rows_raw: Dict[int, Dict[str, str]] = defaultdict(dict)
    for row in sheet_root.findall(".//main:sheetData/main:row", ns):
        r_idx_attr = row.get("r")
        try:
            r_idx = int(r_idx_attr)
        except (TypeError, ValueError):
            continue
        for c in row.findall("main:c", ns):
            ref = c.get("r")
            if not ref:
                continue
            v = c.find("main:v", ns)
            t = c.get("t")
            if v is not None and v.text is not None:
                if t == "s":
                    idx = int(v.text)
                    val = strings[idx] if 0 <= idx < len(strings) else ""
                else:
                    val = v.text
            else:
                val = ""
            m = re.match(r"([A-Z]+)\d+", ref)
            if m:
                rows_raw[r_idx][m.group(1)] = (val or "").strip()

    # 需要的列: B(主/客), D(澳门), F(马会) 用于 morph 匹配
    # E(上水澳), G(上水马), I(水差), K(澳主), L(澳平), M(澳客),
    # O(马主), P(马平), Q(马客), U(主差), V(平差), W(客差), X(澳平客差) 用于条件+范围
    # AB(亚果) 用于统计上/走/下
    data_rows: List[Dict[str, Any]] = []
    num_cols = ["E", "G", "I", "K", "L", "M", "O", "P", "Q", "U", "V", "W", "X", "Y"]
    for r_idx in sorted(rows_raw.keys()):
        if r_idx <= 1:
            continue  # 跳过表头
        row = rows_raw[r_idx]
        b_val = row.get("B", "").strip()
        if b_val not in ("主", "客"):
            continue
        d_raw = row.get("D", "").strip()
        f_raw = row.get("F", "").strip()
        ab_val = row.get("AB", "").strip()
        if not ab_val:
            continue  # 没有结果的行跳过

        try:
            d_val = str(round(float(d_raw), 2)).rstrip("0").rstrip(".")
            f_val = str(round(float(f_raw), 2)).rstrip("0").rstrip(".")
        except (ValueError, TypeError):
            d_val = d_raw
            f_val = f_raw

        parsed: Dict[str, Any] = {
            "side": b_val,
            "D": d_val,
            "F": f_val,
            "result": ab_val,  # 上/走/下
        }
        for col in num_cols:
            raw = row.get(col, "").strip()
            if raw:
                try:
                    parsed[col] = round(float(raw), 2)
                except ValueError:
                    parsed[col] = None
            else:
                parsed[col] = None
        data_rows.append(parsed)

    return data_rows


def _compute_column_ranges(
    base_rows: List[Dict[str, Any]],
) -> Dict[str, Dict[str, Dict[str, float]]]:
    """按 morph 分组统计每个数值列的 min/max。

    返回: { "主|0.25|0.25": { "G": {"min":..,"max":..}, "I": {...}, ... }, ... }
    """
    num_cols = ["G", "I", "L", "O", "P", "U", "V", "W", "X"]
    buckets: Dict[str, Dict[str, List[float]]] = defaultdict(lambda: defaultdict(list))

    for row in base_rows:
        key = f"{row['side']}|{row['D']}|{row['F']}"
        for col in num_cols:
            v = row.get(col)
            if v is not None:
                buckets[key][col].append(v)

    result: Dict[str, Dict[str, Dict[str, float]]] = {}
    for key, cols_data in buckets.items():
        result[key] = {}
        for col, vals in cols_data.items():
            if vals:
                result[key][col] = {
                    "min": round(min(vals), 6),
                    "max": round(max(vals), 6),
                }
    return result


_RANGE_COLS = ["E", "G", "I", "K", "L", "M", "O", "P", "Q", "U", "V", "W", "X"]

_RANGE_COL_NAMES = {
    "E": "澳上水", "G": "马上水", "I": "水差",
    "K": "澳主", "L": "澳平", "M": "澳客",
    "O": "马主", "P": "马平", "Q": "马客",
    "U": "主差", "V": "平差", "W": "客差", "X": "澳平客差",
}


def _compute_ai_stats(
    types: List[Dict[str, Any]], base_rows: List[Dict[str, Any]]
) -> None:
    """对每个规则，用基本表数据统计 AI 上/走/下，并计算匹配行各列的 min/max 范围。"""

    # 按 morph 分桶加速
    morph_buckets: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in base_rows:
        key = f"{row['side']}|{row['D']}|{row['F']}"
        morph_buckets[key].append(row)

    EPS = 1e-9
    for t in types:
        morph = t["morph"]
        key = f"{morph[0]}|{morph[1]}|{morph[2]}"
        candidates = morph_buckets.get(key, [])

        shang = zou = xia = 0
        conds = t.get("conditions", {})
        matched_rows: List[Dict[str, Any]] = []

        for row in candidates:
            ok = True
            for col_name, cond_list in conds.items():
                v = row.get(col_name)
                if v is None:
                    ok = False
                    break
                for c in cond_list:
                    op = c["op"]
                    val = c["value"]
                    if op == ">=" and not (v >= val - EPS):
                        ok = False; break
                    elif op == "<=" and not (v <= val + EPS):
                        ok = False; break
                    elif op == ">" and not (v > val):
                        ok = False; break
                    elif op == "<" and not (v < val):
                        ok = False; break
                    elif op == "=" and not (abs(v - val) <= EPS):
                        ok = False; break
                if not ok:
                    break

            if ok:
                matched_rows.append(row)
                r = row["result"]
                if r == "上":
                    shang += 1
                elif r == "走":
                    zou += 1
                elif r == "下":
                    xia += 1

        t["ai_stats"] = {"shang": shang, "zou": zou, "xia": xia}

        # 计算匹配行各列的 min/max 范围（全部匹配行）
        cond_ranges: Dict[str, Dict[str, float]] = {}
        for col in _RANGE_COLS:
            vals = [r[col] for r in matched_rows if r.get(col) is not None]
            if vals:
                cond_ranges[col] = {
                    "min": round(min(vals), 2),
                    "max": round(max(vals), 2),
                    "name": _RANGE_COL_NAMES.get(col, col),
                }
        t["condition_ranges"] = cond_ranges

        # 只用结果一致性最高的行计算范围
        result_counts = {"上": 0, "下": 0, "走": 0}
        for r in matched_rows:
            res = r.get("result", "")
            if res in result_counts:
                result_counts[res] += 1
        dominant_result = max(result_counts, key=result_counts.get)
        pred_rows = [r for r in matched_rows if r["result"] == dominant_result] if result_counts[dominant_result] > 0 else matched_rows
        pred_ranges: Dict[str, Dict[str, float]] = {}
        for col in _RANGE_COLS:
            vals = [r[col] for r in pred_rows if r.get(col) is not None]
            if vals:
                pred_ranges[col] = {
                    "min": round(min(vals), 2),
                    "max": round(max(vals), 2),
                    "name": _RANGE_COL_NAMES.get(col, col),
                }
        t["pred_ranges"] = pred_ranges


def build_summary_types(
    xlsx_path: str = "docs/202606欧洲FB.xlsx",
    base_xlsx_path: str = "docs/202608欧洲FB基础数据库.xlsx",
) -> Dict[str, Any]:
    """构建基于"汇总"表的手工类型库（只使用人工统计 + AI统计）。"""
    raw_rules = _load_summary_sheet(xlsx_path)
    types: List[Dict[str, Any]] = []

    for idx, r in enumerate(raw_rules, start=1):
        cols = r["cols"]
        group_str = r["group"]  # 例如 "0/0", "0/0.25"
        side = cols.get("A", "")
        if side not in ("主", "客"):
            continue

        try:
            d_group, f_group = [x.strip() for x in group_str.split("/", 1)]
        except ValueError:
            continue

        morph = [side, d_group, f_group]

        # 解析条件
        conds: Dict[str, List[Dict[str, Any]]] = {}
        for col_letter, input_col in _COL_TO_INPUT_COL.items():
            text = cols.get(col_letter, "")
            if not text:
                continue
            parsed = _parse_conditions_for_col(text)
            if parsed:
                conds[input_col] = parsed

        # 人工统计的上/走/下
        def _to_int(x: str) -> int:
            x = (x or "").strip().lstrip(",")
            if not x:
                return 0
            try:
                return int(float(x))
            except Exception:
                return 0

        shang = _to_int(cols.get("L", ""))
        zou = _to_int(cols.get("M", ""))
        xia = _to_int(cols.get("N", ""))

        t = {
            "id": idx,
            "row_index": r["row_index"],
            "group": group_str,
            "morph": morph,
            "conditions": conds,
            "prediction": cols.get("K", ""),
            "feature_text": _build_feature_text(cols),
            "stats": {
                "shang": shang,
                "zou": zou,
                "xia": xia,
            },
            "mark": cols.get("O", ""),
            "tip": cols.get("P", ""),
        }
        types.append(t)

    # AI统计：用基本表数据验证每条规则
    print("正在从基本表加载数据进行AI统计...")
    base_rows = _load_base_table(base_xlsx_path)
    print(f"  基本表数据行: {len(base_rows)}")
    _compute_ai_stats(types, base_rows)

    # 统计差异
    diff_count = 0
    for t in types:
        s = t["stats"]
        ai = t["ai_stats"]
        if s["shang"] != ai["shang"] or s["zou"] != ai["zou"] or s["xia"] != ai["xia"]:
            diff_count += 1
    print(f"  人工与AI统计不一致: {diff_count}/{len(types)} 条")

    # 按 morph 分组统计各列 min/max
    col_ranges = _compute_column_ranges(base_rows)
    print(f"  列范围统计分组数: {len(col_ranges)}")

    return {
        "meta": {
            "source_file": base_xlsx_path,
            "sheet": "汇总",
            "total_types": len(types),
            "base_table_rows": len(base_rows),
        },
        "types": types,
        "column_ranges": col_ranges,
    }


def _export_scatter_data(base_rows: List[Dict[str, Any]], out_path: str) -> None:
    """导出散点图数据：按 morph 分组，包含所有行。

    格式紧凑：每组为 { cols: [...], rows: [[v1,v2,...,result], ...] }
    """
    SCATTER_COLS = ["E", "G", "L", "P", "U", "X", "Y"]
    EXTRA_COLS = ["I", "V", "W"]  # 集中度统计用的额外列
    buckets: Dict[str, List[List[Any]]] = defaultdict(list)

    for row in base_rows:

        key = f"{row['side']}|{row['D']}|{row['F']}"
        vals: List[Any] = []
        for col in SCATTER_COLS:
            v = row.get(col)
            vals.append(v if v is not None else None)
        for col in EXTRA_COLS:
            v = row.get(col)
            vals.append(v if v is not None else None)
        vals.append(row.get("result", ""))
        buckets[key].append(vals)

    scatter = {
        "cols": SCATTER_COLS,
        "extra_cols": EXTRA_COLS,
        "groups": buckets,
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(scatter, f, ensure_ascii=False)
    total = sum(len(v) for v in buckets.values())
    print(f"已导出散点图数据到: {out_path}")
    print(f"  分组数: {len(buckets)}，总行数: {total}")


def main() -> None:
    data = build_summary_types()
    os.makedirs("static", exist_ok=True)
    out_path = os.path.join("static", "summary_types_v2.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"已导出汇总类型库到: {out_path}")
    print(f"  类型总数: {data['meta']['total_types']}")

    # 导出散点图数据
    base_rows = _load_base_table(data["meta"]["source_file"])
    _export_scatter_data(base_rows, os.path.join("static", "scatter_data.json"))


if __name__ == "__main__":
    main()

