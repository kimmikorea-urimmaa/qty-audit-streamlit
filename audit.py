#!/usr/bin/env python3
"""조경 시설물 수량산출서 자동검토 도구 (최종 통합버전)"""

from __future__ import annotations

import argparse
import ast
import csv
import math
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class ErrorRecord:
    row: int
    cell: str
    check_type: str
    reason: str
    severity: str
    related_formula: str = ""
    actual_value: Optional[float] = None
    expected_value: Optional[float] = None
    difference: Optional[float] = None
    tol: Optional[float] = None
    rule_name: str = ""


# =============================
# 기본 유틸
# =============================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="시설물 수량산출서 자동검토")
    parser.add_argument("xlsx", help="입력 XLSX 파일 경로")
    parser.add_argument("--rules", default="rules.yml", help="룰 YAML 파일 경로")
    parser.add_argument("--outdir", default="output", help="결과 출력 폴더")
    return parser.parse_args()


def load_rules(path: str) -> Dict[str, Any]:
    try:
        import yaml
    except Exception as exc:
        raise RuntimeError("pyyaml 미설치: pip install pyyaml 필요") from exc

    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data or {}


def normalize_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def as_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return None
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def has_cell_reference(expr: str) -> bool:
    return bool(re.search(r"\$?[A-Za-z]{1,3}\$?\d+", expr))


# =============================
# 수식 계산
# =============================

def safe_eval_numeric(expr: str) -> Optional[float]:
    expr = expr.strip()
    if expr.startswith("="):
        expr = expr[1:].strip()

    allowed_nodes = (
        ast.Expression,
        ast.BinOp,
        ast.UnaryOp,
        ast.Constant,
        ast.Add,
        ast.Sub,
        ast.Mult,
        ast.Div,
        ast.Pow,
        ast.USub,
        ast.UAdd,
        ast.Mod,
        ast.FloorDiv,
    )

    try:
        tree = ast.parse(expr, mode="eval")
        for node in ast.walk(tree):
            if not isinstance(node, allowed_nodes):
                return None
        result = eval(compile(tree, "<expr>", "eval"), {"__builtins__": {}}, {})
        return as_float(result)
    except Exception:
        return None


# =============================
# ROUND / 허용오차
# =============================

def parse_round_digits(formula: str) -> Optional[int]:
    if not formula:
        return None
    m = re.search(r"ROUND\s*\(.*?,\s*(-?\d+)\s*\)", formula, flags=re.IGNORECASE)
    if not m:
        return None
    return int(m.group(1))


def get_round_digits(e_formula: str, default_digits: int) -> int:
    n = parse_round_digits(e_formula)
    return n if n is not None else default_digits


def tol_from_round_digits(round_digits: int) -> float:
    return 0.6 * (10 ** (-round_digits))


# =============================
# 분류 로직
# =============================

def classify_row_type(work: str, spec: str, unit: str, bigo: str, rules: Dict[str, Any]) -> str:
    text = f"{work} {spec} {unit} {bigo}".lower()

    material_keys = [str(x).lower() for x in rules.get("material_keywords_any", [])]
    install_keys = [str(x).lower() for x in rules.get("installation_keywords_any", [])]

    if any(k in text for k in material_keys):
        return "material"
    if any(k in text for k in install_keys):
        return "installation"
    return "unknown"


# =============================
# 리포트
# =============================

def build_reports(errors: List[ErrorRecord], outdir: str) -> None:
    from openpyxl import Workbook

    os.makedirs(outdir, exist_ok=True)

    columns = [
        "row", "cell", "check_type", "reason", "severity",
        "rule_name", "related_formula",
        "actual_value", "expected_value",
        "difference", "tol"
    ]

    csv_path = os.path.join(outdir, "report.csv")
    xlsx_path = os.path.join(outdir, "report.xlsx")

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        for e in errors:
            writer.writerow([
                e.row, e.cell, e.check_type, e.reason, e.severity,
                e.rule_name, e.related_formula,
                e.actual_value, e.expected_value,
                e.difference, e.tol
            ])

    wb = Workbook()
    ws = wb.active
    ws.title = "Errors"
    ws.append(columns)
    for e in errors:
        ws.append([
            e.row, e.cell, e.check_type, e.reason, e.severity,
            e.rule_name, e.related_formula,
            e.actual_value, e.expected_value,
            e.difference, e.tol
        ])
    wb.save(xlsx_path)


# =============================
# 메인 로직
# =============================

def main():
    args = parse_args()
    rules = load_rules(args.rules)

    from openpyxl import load_workbook
    wb_formula = load_workbook(args.xlsx, data_only=False)
    wb_value = load_workbook(args.xlsx, data_only=True)

    sheet = wb_formula.sheetnames[0]
    ws_formula = wb_formula[sheet]
    ws_value = wb_value[sheet]

    default_round_digits = int(rules.get("round_default_digits", 3))
    errors: List[ErrorRecord] = []

    for r in range(2, ws_formula.max_row + 1):

        work = normalize_text(ws_formula.cell(r, 2).value)
        spec = normalize_text(ws_formula.cell(r, 3).value)
        d_text = normalize_text(ws_formula.cell(r, 4).value)
        e_formula = normalize_text(ws_formula.cell(r, 5).value)
        e_value = as_float(ws_value.cell(r, 5).value)
        unit = normalize_text(ws_formula.cell(r, 6).value)
        bigo = normalize_text(ws_formula.cell(r, 7).value)

        if not any([work, spec, d_text, e_formula, unit, bigo]):
            continue

        round_digits = get_round_digits(e_formula, default_round_digits)
        tol = tol_from_round_digits(round_digits)
        row_type = classify_row_type(work, spec, unit, bigo, rules)

        # === 산출근거 계산 비교 ===
        if d_text and not has_cell_reference(d_text):
            d_val = safe_eval_numeric(d_text)
            if d_val is not None and e_value is not None:
                expected = round(d_val, round_digits)
                diff = abs(expected - e_value)
                if diff > tol:
                    errors.append(
                        ErrorRecord(
                            r, f"D{r}/E{r}",
                            "calc_text_check",
                            f"ROUND {round_digits}자리 기준 불일치",
                            "HIGH",
                            f"D:{d_text} | E:{e_formula}",
                            e_value, expected, diff, tol
                        )
                    )

        # === 정책 검사 ===
        percent_detected = "%" in bigo

        if row_type == "installation" and percent_detected:
            errors.append(
                ErrorRecord(
                    r, f"E{r}",
                    "allowance_policy_check",
                    "설치품에 할증이 적용됨",
                    "HIGH",
                    f"D:{d_text} | E:{e_formula}"
                )
            )

        if row_type == "material" and not percent_detected:
            errors.append(
                ErrorRecord(
                    r, f"E{r}",
                    "allowance_policy_check",
                    "재료에 할증이 누락됨",
                    "MEDIUM",
                    f"D:{d_text} | E:{e_formula}"
                )
            )

    build_reports(errors, args.outdir)
    print("[OK] 완료")


if __name__ == "__main__":
    main()
