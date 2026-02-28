#!/usr/bin/env python3
"""조경 시설물 수량산출서 자동검토 도구 - 실무 안정 최종본"""

from __future__ import annotations

import argparse
import ast
import csv
import math
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple


# ==============================
# 데이터 구조
# ==============================

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


# ==============================
# 기본 유틸
# ==============================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="시설물 수량산출서 자동검토")
    p.add_argument("xlsx", help="입력 XLSX 파일 경로")
    p.add_argument("--rules", default="rules.yml", help="룰 YAML 파일 경로")
    p.add_argument("--outdir", default="output", help="결과 출력 폴더")
    return p.parse_args()


def load_rules(path: str) -> Dict[str, Any]:
    try:
        import yaml
    except Exception:
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def normalize_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def as_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return None
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None


def has_cell_reference(expr: str) -> bool:
    return bool(re.search(r"\$?[A-Za-z]{1,3}\$?\d+", expr))


# ==============================
# ROUND / 허용오차
# ==============================

def parse_round_digits(formula: str) -> Optional[int]:
    if not formula:
        return None
    m = re.search(r"ROUND\s*\(.*?,\s*(-?\d+)\s*\)", formula, flags=re.IGNORECASE)
    return int(m.group(1)) if m else None


def get_round_digits(e_formula: str, default_digits: int = 3) -> int:
    n = parse_round_digits(e_formula)
    return n if n is not None else default_digits


def tol_from_round_digits(round_digits: int) -> float:
    """
    실무 기준:
    - 최소 허용오차 0.01
    """
    if round_digits <= 0:
        base_tol = 1.0
    else:
        base_tol = 2.0 * (10 ** (-round_digits))
    return max(base_tol, 0.01)


# ==============================
# 수식 계산
# ==============================

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
    )

    try:
        tree = ast.parse(expr, mode="eval")
        for node in ast.walk(tree):
            if not isinstance(node, allowed_nodes):
                return None
        return float(eval(compile(tree, "<expr>", "eval"), {"__builtins__": {}}, {}))
    except Exception:
        return None


def d_has_multiplier(d_text: str, mult: float) -> bool:
    """D에 이미 *1.04 같은 계수 포함 여부 검사"""
    if not d_text:
        return False
    nums = re.findall(r"\*\s*([0-9]+(?:\.[0-9]+)?)", d_text.replace(",", ""))
    for n in nums:
        try:
            if abs(float(n) - mult) < 1e-9:
                return True
        except Exception:
            pass
    return False


# ==============================
# 행 분류
# ==============================

def classify_row_type(work: str, spec: str, unit: str, bigo: str, rules: Dict[str, Any]) -> str:
    text = f"{work} {spec} {unit} {bigo}".lower()
    material_keys = [str(x).lower() for x in (rules.get("material_keywords_any") or [])]
    install_keys = [str(x).lower() for x in (rules.get("installation_keywords_any") or [])]

    if material_keys and any(k in text for k in material_keys):
        return "material"
    if install_keys and any(k in text for k in install_keys):
        return "installation"
    return "unknown"


# ==============================
# 메인
# ==============================

def main():
    args = parse_args()
    rules = load_rules(args.rules)

    from openpyxl import load_workbook

    wb_formula = load_workbook(args.xlsx, data_only=False)
    wb_value = load_workbook(args.xlsx, data_only=True)

    ws_formula = wb_formula.active
    ws_value = wb_value.active

    percent_regex = re.compile(r"(\d+(\.\d+)?)%")
    allowance_map = rules.get("allowance_multiplier_map", {})

    errors: List[ErrorRecord] = []

    for r in range(2, ws_formula.max_row + 1):
        work = normalize_text(ws_formula[f"B{r}"].value)
        spec = normalize_text(ws_formula[f"C{r}"].value)
        d_text = normalize_text(ws_formula[f"D{r}"].value)
        e_formula = normalize_text(ws_formula[f"E{r}"].value)
        e_value = as_float(ws_value[f"E{r}"].value)
        unit = normalize_text(ws_formula[f"F{r}"].value)
        bigo = normalize_text(ws_formula[f"G{r}"].value)

        if not any([work, spec, d_text, e_formula, unit, bigo]):
            continue

        round_digits = get_round_digits(e_formula)
        tol = tol_from_round_digits(round_digits)

        # -------------------------
        # calc_text_check
        # -------------------------
        if d_text and not has_cell_reference(d_text):
            d_numeric = safe_eval_numeric(d_text)
            if d_numeric is not None and e_value is not None:
                expected = round(d_numeric, round_digits)
                diff = abs(expected - e_value)
                if diff > tol:
                    errors.append(ErrorRecord(
                        row=r,
                        cell=f"D{r}/E{r}",
                        check_type="calc_text_check",
                        reason="D 계산값과 E 수량 불일치",
                        severity="HIGH",
                        related_formula=f"D:{d_text} | E:{e_formula}",
                        actual_value=e_value,
                        expected_value=expected,
                        difference=diff,
                        tol=tol,
                    ))

        # -------------------------
        # 할증 (% 있을 때만)
        # -------------------------
        m = percent_regex.search(bigo)
        if not m:
            continue

        percent_text = f"{m.group(1)}%"
        if percent_text not in allowance_map:
            continue

        multiplier = float(allowance_map[percent_text])
        row_type = classify_row_type(work, spec, unit, bigo, rules)

        if row_type == "installation":
            errors.append(ErrorRecord(
                row=r,
                cell=f"E{r}",
                check_type="allowance_policy_check",
                reason="설치품에 할증% 명시됨",
                severity="HIGH",
            ))
            continue

        if row_type == "material":
            d_numeric = safe_eval_numeric(d_text)
            if d_numeric is None or e_value is None:
                continue

            if d_has_multiplier(d_text, multiplier):
                expected = round(d_numeric, round_digits)
            else:
                expected = round(d_numeric * multiplier, round_digits)

            diff = abs(expected - e_value)
            if diff > tol:
                errors.append(ErrorRecord(
                    row=r,
                    cell=f"E{r}",
                    check_type="allowance_check",
                    reason="비고 할증 적용값과 E 수량 불일치",
                    severity="MEDIUM",
                    related_formula=f"D:{d_text} | E:{e_formula}",
                    actual_value=e_value,
                    expected_value=expected,
                    difference=diff,
                    tol=tol,
                ))

    print(f"[완료] 오류 건수: {len(errors)}")


if __name__ == "__main__":
    main()
