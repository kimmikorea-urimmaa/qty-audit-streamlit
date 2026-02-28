#!/usr/bin/env python3
"""조경 시설물 수량산출서 자동검토 도구 - 실무 안정 (할증 2건 현상 원인추적/복구)

핵심
- tol 최소 0.01
- 비고(BIGO)에 % 있을 때만 할증 검토
- 설치품 공종명 리스트는 할증 검증(allowance_check) 제외
- D에 이미 *1.04 같은 계수가 있으면 이중할증 방지
- report.csv / report.xlsx 생성
- ✅ 왜 할증 오류가 안 나오는지 '스킵 사유 카운터'를 로그로 출력
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
    w = (work or "").strip()
    if not w:
        return False
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
    except Exception as exc:
        raise RuntimeError("pyyaml 미설치: `pip install pyyaml` 필요") from exc
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


# --------------------------
# 시트/헤더 선택 강화
# --------------------------
HEADER_KEYWORDS = ["공종", "규격", "산출근거", "수량", "단위", "비고"]


def score_sheet(ws) -> int:
    """시트 내 상단에서 헤더 키워드가 얼마나 보이는지로 점수화."""
    max_r = min(ws.max_row, 40)
    max_c = min(ws.max_column, 60)
    score = 0
    for r in range(1, max_r + 1):
        row_text = " ".join(normalize_text(ws.cell(r, c).value) for c in range(1, max_c + 1))
        for k in HEADER_KEYWORDS:
            if k in row_text:
                score += 2
        # '산출근거/수량'은 특히 중요
        if "산출근거" in row_text:
            score += 5
        if "수량" in row_text:
            score += 5
    return score


def choose_sheet_name(wb) -> str:
    # 1) 점수 가장 높은 시트 선택
    best = None
    best_score = -1
    for name in wb.sheetnames:
        ws = wb[name]
        s = score_sheet(ws)
        if s > best_score:
            best_score = s
            best = name
    return best or wb.sheetnames[0]


def detect_columns(ws) -> Tuple[int, Dict[str, int]]:
    """
    헤더에서 열 자동 탐지.
    ✅ hit 기준을 4로 올려서(공종/규격/산출근거/수량 등) 엉뚱한 행을 헤더로 잡는 오탐을 줄임.
    """
    columns = {"work": 2, "spec": 3, "basis": 4, "qty": 5, "unit": 6, "bigo": 7}
    header_row = 1

    for r in range(1, min(ws.max_row, 40) + 1):
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 60) + 1)]
        hit = 0
        for idx, text in enumerate(row_values, start=1):
            low = text.lower()
            if "공종" in text:
                columns["work"] = idx
                hit += 1
            if "규격" in text:
                columns["spec"] = idx
                hit += 1
            if "산출근거" in text:
                columns["basis"] = idx
                hit += 1
            if "수량" in text:
                columns["qty"] = idx
                hit += 1
            if "단위" in text:
                columns["unit"] = idx
                hit += 1
            if "비고" in text or "remark" in low:
                columns["bigo"] = idx
                hit += 1
        if hit >= 4:
            header_row = r
            break

    return header_row, columns


# --------------------------
# ROUND / tol
# --------------------------
def parse_round_digits(formula: str) -> Optional[int]:
    if not formula:
        return None
    m = re.search(r"ROUND\s*\(.*?,\s*(-?\d+)\s*\)", formula, flags=re.IGNORECASE)
    return int(m.group(1)) if m else None


def get_round_digits(e_formula: str, default_digits: int = 3) -> int:
    n = parse_round_digits(e_formula)
    return n if n is not None else default_digits


def tol_from_round_digits(round_digits: int) -> float:
    if round_digits <= 0:
        base_tol = 1.0
    else:
        base_tol = 2.0 * (10 ** (-round_digits))
    return max(base_tol, 0.01)


# --------------------------
# 수식 계산
# --------------------------
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
        return float(eval(compile(tree, "<expr>", "eval"), {"__builtins__": {}}, {}))
    except Exception:
        return None


def d_has_multiplier(d_text: str, mult: float) -> bool:
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


# --------------------------
# report
# --------------------------
def build_reports(errors: List[ErrorRecord], outdir: str) -> Tuple[str, str]:
    from openpyxl import Workbook

    os.makedirs(outdir, exist_ok=True)
    csv_path = os.path.join(outdir, "report.csv")
    xlsx_path = os.path.join(outdir, "report.xlsx")

    columns = [
        "row", "cell", "check_type", "reason", "severity", "rule_name",
        "related_formula", "actual_value", "expected_value", "difference", "tol",
    ]

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(columns)
        for e in errors:
            w.writerow([
                e.row, e.cell, e.check_type, e.reason, e.severity, e.rule_name,
                e.related_formula, e.actual_value, e.expected_value, e.difference, e.tol
            ])

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


def main() -> None:
    args = parse_args()
    rules = load_rules(args.rules)

    from openpyxl import load_workbook

    wb_formula = load_workbook(args.xlsx, data_only=False)
    wb_value = load_workbook(args.xlsx, data_only=True)

    sheet_name = choose_sheet_name(wb_formula)
    ws_formula = wb_formula[sheet_name]
    ws_value = wb_value[sheet_name]

    header_row, cols = detect_columns(ws_formula)

    percent_regex = re.compile(rules.get("allowance_percent_extract_regex", r"(\d+(\.\d+)?)%"))
    allowance_map = rules.get("allowance_multiplier_map", {})
    default_round_digits = int(rules.get("round_default_digits", 3))

    errors: List[ErrorRecord] = []

    # ✅ 원인 추적 카운터
    rows_total = 0
    rows_with_percent = 0
    skipped_installation = 0
    skipped_no_e_value = 0
    skipped_no_d_numeric = 0
    skipped_no_map = 0

    for r in range(header_row + 1, ws_formula.max_row + 1):
        work = normalize_text(ws_formula.cell(r, cols["work"]).value)
        spec = normalize_text(ws_formula.cell(r, cols["spec"]).value)
        d_text = normalize_text(ws_formula.cell(r, cols["basis"]).value)
        e_formula = normalize_text(ws_formula.cell(r, cols["qty"]).value)
        e_value = as_float(ws_value.cell(r, cols["qty"]).value)
        unit = normalize_text(ws_formula.cell(r, cols["unit"]).value)
        bigo = normalize_text(ws_formula.cell(r, cols["bigo"]).value)

        if not any([work, spec, d_text, e_formula, unit, bigo]):
            continue

        rows_total += 1

        round_digits = get_round_digits(e_formula, default_digits=default_round_digits)
        tol = tol_from_round_digits(round_digits)

        # calc_text_check
        d_numeric: Optional[float] = None
        if d_text and not has_cell_reference(d_text):
            d_numeric = safe_eval_numeric(d_text)

        if d_numeric is not None and e_value is not None:
            expected = round(d_numeric, round_digits)
            diff = abs(expected - e_value)
            if diff > tol:
                errors.append(ErrorRecord(
                    row=r, cell=f"{ws_formula.cell(r, cols['basis']).coordinate}/{ws_formula.cell(r, cols['qty']).coordinate}",
                    check_type="calc_text_check",
                    reason=f"D 계산값과 E 수량 불일치(ROUND {round_digits})",
                    severity="HIGH",
                    related_formula=f"D:{d_text} | E:{e_formula} | BIGO:{bigo}",
                    actual_value=e_value, expected_value=expected,
                    difference=diff, tol=tol,
                    rule_name=f"ROUND({round_digits})",
                ))

        # allowance: 비고% 있을 때만
        m = percent_regex.search(bigo)
        if not m:
            continue

        rows_with_percent += 1

        # 설치품은 할증 검증 제외
        if is_installation_item(work):
            skipped_installation += 1
            continue

        percent_text = f"{m.group(1)}%"
        if percent_text not in allowance_map:
            skipped_no_map += 1
            errors.append(ErrorRecord(
                row=r, cell=f"{ws_formula.cell(r, cols['bigo']).coordinate}",
                check_type="allowance_check",
                reason=f"비고 '{percent_text}'가 allowance_multiplier_map에 없음",
                severity="MEDIUM",
                related_formula=f"BIGO:{bigo}",
                rule_name="allowance_multiplier_map missing",
            ))
            continue

        if e_value is None:
            skipped_no_e_value += 1
            continue

        if d_numeric is None:
            skipped_no_d_numeric += 1
            continue

        multiplier = float(allowance_map[percent_text])

        if d_has_multiplier(d_text, multiplier):
            expected_allow = round(d_numeric, round_digits)
            rule2 = f"비고 {percent_text} | D already has multiplier"
        else:
            expected_allow = round(d_numeric * multiplier, round_digits)
            rule2 = f"비고 {percent_text} | D * multiplier"

        diff2 = abs(expected_allow - e_value)
        if diff2 > tol:
            errors.append(ErrorRecord(
                row=r, cell=f"{ws_formula.cell(r, cols['qty']).coordinate}",
                check_type="allowance_check",
                reason=f"비고 할증 적용값과 E 수량 불일치(ROUND {round_digits})",
                severity="MEDIUM",
                related_formula=f"D:{d_text} | E:{e_formula} | BIGO:{bigo}",
                actual_value=e_value, expected_value=expected_allow,
                difference=diff2, tol=tol,
                rule_name=rule2,
            ))

    csv_path, xlsx_path = build_reports(errors, args.outdir)

    print(f"[OK] sheet='{sheet_name}', header_row={header_row}, rows_checked={rows_total}")
    print(f"[OK] with_percent={rows_with_percent}, skipped_installation={skipped_installation}, skipped_no_map={skipped_no_map}")
    print(f"[OK] skipped_no_e_value={skipped_no_e_value}, skipped_no_d_numeric={skipped_no_d_numeric}")
    print(f"[완료] 오류 건수: {len(errors)}")
    print(f"[OK] report saved: {csv_path} / {xlsx_path}")


if __name__ == "__main__":
    main()
