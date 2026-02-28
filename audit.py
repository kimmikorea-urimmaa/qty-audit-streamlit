#!/usr/bin/env python3
"""조경 시설물 수량산출서 자동검토 도구 - 실무 안정 최종본

핵심
- tol 최소 0.01
- 비고(BIGO)에 % 있을 때만 할증 검토
- 설치품 공종명(고정 리스트)에 해당하면 할증 검증(allowance_check) 제외
- D에 이미 *1.04 같은 계수가 있으면 이중할증 방지
- report.csv / report.xlsx 생성

검토 항목
1) calc_text_check (항상):
   - D(산출근거) 텍스트 수식을 계산하여 E(수량) 값과 비교
   - E가 ROUND(…,n)이면 n 사용, 없으면 기본 n(기본 3)
   - tol(허용오차) = max(ROUND기반, 0.01)

2) allowance_check (비고% 있을 때만):
   - 설치품이면 스킵(검증 제외)
   - 설치품이 아니면 D와 E가 비고%에 맞는지 검증
   - D에 이미 계수가 있으면 이중할증 방지
"""

from __future__ import annotations

import argparse
import ast
import csv
import math
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple


# ------------------------------
# 설치품 공종명(사용자 제공)
# ------------------------------
INSTALLATION_WORK_NAMES = [
    "혼합골재포설및다짐",
    "레미콘타설",
    "통석놓기",
    "철근가공조립",
    "잡철물제작및설치",
    "목재가공 및 설치",
    "목재가공및설치",
    "플랜터 설치",
    "플랜터설치",
    "우레탄도장",
    "석재판석붙임",
    "친환경스테인도장",
    "데크깔기",
]


def is_installation_item(work: str) -> bool:
    """공종명 기준 설치품 여부(띄어쓰기/특수문자 차이를 조금 흡수)."""
    w = (work or "").strip()
    if not w:
        return False

    # 비교용 정규화: 공백 제거
    w_norm = re.sub(r"\s+", "", w)
    for key in INSTALLATION_WORK_NAMES:
        key_norm = re.sub(r"\s+", "", key.strip())
        if key_norm and key_norm in w_norm:
            return True
    return False


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


# ------------------------------
# ROUND / 허용오차
# ------------------------------
def parse_round_digits(formula: str) -> Optional[int]:
    if not formula:
        return None
    m = re.search(r"ROUND\s*\(.*?,\s*(-?\d+)\s*\)", formula, flags=re.IGNORECASE)
    return int(m.group(1)) if m else None


def get_round_digits(e_formula: str, default_digits: int = 3) -> int:
    n = parse_round_digits(e_formula)
    return n if n is not None else default_digits


def tol_from_round_digits(round_digits: int) -> float:
    # 최소 0.01 허용
    if round_digits <= 0:
        base_tol = 1.0
    else:
        base_tol = 2.0 * (10 ** (-round_digits))
    return max(base_tol, 0.01)


# ------------------------------
# 수식 계산
# ------------------------------
def safe_eval_numeric(expr: str) -> Optional[float]:
    """숫자/연산자/괄호만 허용해 계산(셀참조/문자 포함 시 None)."""
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
        return float(eval(compile(tree, "<expr>", "eval"), {"__builtins__": {}}, {}))
    except Exception:
        return None


def d_has_multiplier(d_text: str, mult: float) -> bool:
    """D 산출근거에 '* 1.04' 같이 계수가 직접 포함되어 있는지 감지(이중 할증 방지)."""
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


# ------------------------------
# 리포트 생성
# ------------------------------
def build_reports(errors: List[ErrorRecord], outdir: str) -> Tuple[str, str]:
    from openpyxl import Workbook

    os.makedirs(outdir, exist_ok=True)
    csv_path = os.path.join(outdir, "report.csv")
    xlsx_path = os.path.join(outdir, "report.xlsx")

    columns = [
        "row",
        "cell",
        "check_type",
        "reason",
        "severity",
        "rule_name",
        "related_formula",
        "actual_value",
        "expected_value",
        "difference",
        "tol",
    ]

    # CSV
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(columns)
        for e in errors:
            w.writerow([
                e.row, e.cell, e.check_type, e.reason, e.severity, e.rule_name,
                e.related_formula, e.actual_value, e.expected_value, e.difference, e.tol
            ])

    # XLSX
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.append(["check_type", "severity", "count"])

    summary: Dict[Tuple[str, str], int] = {}
    for e in errors:
        summary[(e.check_type, e.severity)] = summary.get((e.check_type, e.severity), 0) + 1
    for (ct, sev), cnt in sorted(summary.items(), key=lambda x: (x[0][0], x[0][1])):
        ws_sum.append([ct, sev, cnt])

    ws_err = wb.create_sheet("Errors")
    ws_err.append(columns)
    for e in errors:
        ws_err.append([
            e.row, e.cell, e.check_type, e.reason, e.severity, e.rule_name,
            e.related_formula, e.actual_value, e.expected_value, e.difference, e.tol
        ])

    wb.save(xlsx_path)
    return csv_path, xlsx_path


# ------------------------------
# 메인
# ------------------------------
def main() -> None:
    args = parse_args()
    rules = load_rules(args.rules)

    from openpyxl import load_workbook

    wb_formula = load_workbook(args.xlsx, data_only=False)
    wb_value = load_workbook(args.xlsx, data_only=True)

    ws_formula = wb_formula.active
    ws_value = wb_value.active

    percent_regex = re.compile(rules.get("allowance_percent_extract_regex", r"(\d+(\.\d+)?)%"))
    allowance_map = rules.get("allowance_multiplier_map", {})
    default_round_digits = int(rules.get("round_default_digits", 3))

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

        round_digits = get_round_digits(e_formula, default_digits=default_round_digits)
        tol = tol_from_round_digits(round_digits)

        # -------------------------
        # 1) calc_text_check (항상)
        # -------------------------
        d_numeric: Optional[float] = None
        if d_text and not has_cell_reference(d_text):
            d_numeric = safe_eval_numeric(d_text)

        if d_numeric is not None and e_value is not None:
            expected = round(d_numeric, round_digits)
            diff = abs(expected - e_value)
            if diff > tol:
                errors.append(ErrorRecord(
                    row=r, cell=f"D{r}/E{r}",
                    check_type="calc_text_check",
                    reason=f"D 계산값과 E 수량 불일치(ROUND {round_digits})",
                    severity="HIGH",
                    related_formula=f"D:{d_text} | E:{e_formula} | BIGO:{bigo}",
                    actual_value=e_value, expected_value=expected,
                    difference=diff, tol=tol,
                    rule_name=f"ROUND({round_digits})",
                ))

        # -------------------------
        # 2) allowance_check (비고% 있을 때만)
        #    ✅ 설치품이면 할증 검증 제외
        # -------------------------
        m = percent_regex.search(bigo)
        if not m:
            continue  # 할증 검토만 스킵

        # 설치품이면 할증검증 제외(요청사항)
        if is_installation_item(work):
            continue

        percent_text = f"{m.group(1)}%"
        if percent_text not in allowance_map:
            errors.append(ErrorRecord(
                row=r, cell=f"G{r}",
                check_type="allowance_check",
                reason=f"비고에 '{percent_text}'가 있으나 allowance_multiplier_map에 정의되지 않음",
                severity="MEDIUM",
                related_formula=f"BIGO:{bigo}",
                rule_name="allowance_multiplier_map missing",
            ))
            continue

        multiplier = float(allowance_map[percent_text])

        # D 계산이 가능해야 검증 가능(셀참조 있으면 skip)
        if d_numeric is None or e_value is None:
            continue

        # D에 이미 *1.04 포함이면 이중할증 금지: expected=D 자체
        if d_has_multiplier(d_text, multiplier):
            expected_allow = round(d_numeric, round_digits)
            rule2 = f"비고 {percent_text} | D already has multiplier"
        else:
            expected_allow = round(d_numeric * multiplier, round_digits)
            rule2 = f"비고 {percent_text} | D * multiplier"

        diff2 = abs(expected_allow - e_value)
        if diff2 > tol:
            errors.append(ErrorRecord(
                row=r, cell=f"E{r}",
                check_type="allowance_check",
                reason=f"비고 할증 적용값과 E 수량 불일치(ROUND {round_digits})",
                severity="MEDIUM",
                related_formula=f"D:{d_text} | E:{e_formula} | BIGO:{bigo}",
                actual_value=e_value, expected_value=expected_allow,
                difference=diff2, tol=tol,
                rule_name=rule2,
            ))

    csv_path, xlsx_path = build_reports(errors, args.outdir)
    print(f"[완료] 오류 건수: {len(errors)}")
    print(f"[OK] report saved: {csv_path} / {xlsx_path}")


if __name__ == "__main__":
    main()
