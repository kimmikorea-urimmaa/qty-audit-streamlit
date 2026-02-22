#!/usr/bin/env python3
"""조경 시설물 수량산출서 자동검토 도구."""

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
    rule_name: str = ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="시설물 수량산출서 자동검토")
    parser.add_argument("xlsx", help="입력 XLSX 파일 경로")
    parser.add_argument("--rules", default="rules.yml", help="룰 YAML 파일 경로")
    parser.add_argument("--outdir", default="output", help="결과 출력 폴더")
    return parser.parse_args()


def load_rules(path: str) -> Dict[str, Any]:
    # 외부 의존성 import는 런타임으로 지연하여 --help 동작 안정성 확보
    try:
        import yaml
    except Exception as exc:
        raise RuntimeError("pyyaml 미설치: `pip install pyyaml` 필요") from exc

    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data or {}


def normalize_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def choose_sheet_name(wb) -> str:
    if "시설물산출" in wb.sheetnames:
        return "시설물산출"

    scored: List[Tuple[int, str]] = []
    for name in wb.sheetnames:
        score = 0
        if "시설물" in name:
            score += 2
        if "산출" in name:
            score += 2
        if "수량" in name:
            score += 1
        if score > 0:
            scored.append((score, name))

    if scored:
        scored.sort(reverse=True)
        return scored[0][1]

    return wb.sheetnames[0]


def detect_columns(ws) -> Tuple[int, Dict[str, int]]:
    """헤더에서 열 자동 탐지, 실패 시 기본 B~G 사용."""
    columns = {
        "work": 2,
        "spec": 3,
        "basis": 4,
        "qty": 5,
        "unit": 6,
        "bigo": 7,
    }
    header_row = 1

    for r in range(1, min(ws.max_row, 20) + 1):
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 40) + 1)]
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
        if hit >= 2:
            header_row = r
            break

    return header_row, columns


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


def safe_eval_numeric(expr: str) -> Optional[float]:
    """숫자/연산자/괄호만 허용해 계산."""
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


def parse_round_digits(formula: str) -> Optional[int]:
    if not formula:
        return None
    m = re.search(r"ROUND\s*\(.*?,\s*(-?\d+)\s*\)", formula, flags=re.IGNORECASE)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def rule_matches(rule: Dict[str, Any], combined_text: str, row_bigo_text: str) -> bool:
    keywords = rule.get("keywords_any", [])
    require = rule.get("require_in_row_or_bigo_any", [])
    case_insensitive = bool(rule.get("case_insensitive", True))

    src1 = combined_text.lower() if case_insensitive else combined_text
    src2 = row_bigo_text.lower() if case_insensitive else row_bigo_text
    keys = [k.lower() for k in keywords] if case_insensitive else keywords
    reqs = [k.lower() for k in require] if case_insensitive else require

    if keys and not any(k in src1 for k in keys):
        return False
    if reqs and not any(k in src2 for k in reqs):
        return False
    return True


def pick_single_allowance_rule(
    standard_rules: List[Dict[str, Any]],
    allowance_map: Dict[str, Any],
    percent_match_text: Optional[str],
    combined_text: str,
    row_bigo_text: str,
) -> Optional[Dict[str, Any]]:
    """한 행에 다중 규칙이 매칭되어도 priority 최대 1개만 선택."""
    candidates: List[Dict[str, Any]] = []

    if percent_match_text and percent_match_text in allowance_map:
        # 비고에 명시된 %는 사용자 직접지정으로 간주, 높은 우선순위 부여
        candidates.append(
            {
                "name": f"비고 퍼센트({percent_match_text})",
                "priority": 1000,
                "multiplier": float(allowance_map[percent_match_text]),
            }
        )

    for rule in standard_rules:
        if rule_matches(rule, combined_text, row_bigo_text):
            candidates.append(rule)

    if not candidates:
        return None

    candidates.sort(key=lambda x: int(x.get("priority", 0)), reverse=True)
    return candidates[0]


def unit_weight_check(work: str, spec: str, unit: str) -> List[Tuple[str, str, str]]:
    issues: List[Tuple[str, str, str]] = []
    w = f"{work} {spec}".lower()
    u = unit.strip().lower()

    if "아연도각관" in f"{work} {spec}":
        if u == "kg":
            issues.append(("HIGH", "아연도각관은 m 단가 처리 가능 품목인데 kg 단위로 입력됨", "unit_weight:아연도각관"))
        return issues

    if "st pl" in w or "sts pl" in w:
        if not re.search(r"\bT\s*\d+(\.\d+)?\b", spec, flags=re.IGNORECASE):
            issues.append(("MEDIUM", "PL 품목인데 규격에 두께(T값) 정보가 없음", "unit_weight:plate-thickness"))

    if "angle" in w and u in {"m", "m2", "㎡"}:
        issues.append(("LOW", "angle 품목은 39.65 kg/m2 기준 검토 대상", "unit_weight:angle-39.65"))

    if "이형철근" in w:
        if not re.search(r"\bD\s*\d+\b", spec, flags=re.IGNORECASE):
            issues.append(("MEDIUM", "이형철근 품목인데 규격에 D값이 없음", "unit_weight:rebar-diameter"))

    return issues


def build_reports(errors: List[ErrorRecord], outdir: str) -> None:
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
    ]

    # CSV
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        for e in errors:
            writer.writerow([
                e.row,
                e.cell,
                e.check_type,
                e.reason,
                e.severity,
                e.rule_name,
                e.related_formula,
                e.actual_value,
                e.expected_value,
                e.difference,
            ])

    # Summary 집계
    summary: Dict[Tuple[str, str], int] = {}
    for e in errors:
        key = (e.check_type, e.severity)
        summary[key] = summary.get(key, 0) + 1

    # XLSX
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.append(["check_type", "severity", "count"])
    for (check_type, severity), cnt in sorted(summary.items(), key=lambda x: (x[0][0], x[0][1])):
        ws_summary.append([check_type, severity, cnt])

    ws_errors = wb.create_sheet("Errors")
    ws_errors.append(columns)
    for e in errors:
        ws_errors.append([
            e.row,
            e.cell,
            e.check_type,
            e.reason,
            e.severity,
            e.rule_name,
            e.related_formula,
            e.actual_value,
            e.expected_value,
            e.difference,
        ])

    wb.save(xlsx_path)


def main() -> None:
    args = parse_args()

    try:
        rules = load_rules(args.rules)
    except Exception as e:
        raise SystemExit(f"[ERROR] rules 로드 실패: {e}")

    try:
        from openpyxl import load_workbook

        wb_formula = load_workbook(args.xlsx, data_only=False)
        wb_value = load_workbook(args.xlsx, data_only=True)
    except Exception as e:
        raise SystemExit(f"[ERROR] xlsx 로드 실패: {e}")

    sheet_name = choose_sheet_name(wb_formula)
    ws_formula = wb_formula[sheet_name]
    ws_value = wb_value[sheet_name]

    header_row, cols = detect_columns(ws_formula)

    percent_regex = re.compile(rules.get("allowance_percent_extract_regex", r"(\d+(\.\d+)?)%"))
    allowance_map = rules.get("allowance_multiplier_map", {})
    standard_rules = rules.get("allowance_standard_rules", [])
    threshold = float(rules.get("calc_text_check", {}).get("error_threshold_abs", 0.009))

    errors: List[ErrorRecord] = []

    for r in range(header_row + 1, ws_formula.max_row + 1):
        work = normalize_text(ws_formula.cell(r, cols["work"]).value)
        spec = normalize_text(ws_formula.cell(r, cols["spec"]).value)
        d_formula_or_text = normalize_text(ws_formula.cell(r, cols["basis"]).value)
        e_formula = normalize_text(ws_formula.cell(r, cols["qty"]).value)
        e_value = as_float(ws_value.cell(r, cols["qty"]).value)
        unit = normalize_text(ws_formula.cell(r, cols["unit"]).value)
        bigo = normalize_text(ws_formula.cell(r, cols["bigo"]).value)

        if not any([work, spec, d_formula_or_text, e_formula, unit, bigo]):
            continue

        combined = f"{work} {spec} {bigo}"
        row_bigo = f"{work} {bigo}"

        # 1) calc_text_check
        d_numeric: Optional[float] = None
        if d_formula_or_text and not has_cell_reference(d_formula_or_text):
            d_numeric = safe_eval_numeric(d_formula_or_text)
            round_digits = parse_round_digits(e_formula)
            if d_numeric is not None and e_value is not None and round_digits is not None:
                expected = round(d_numeric, round_digits)
                diff = abs(expected - e_value)
                if diff >= threshold:
                    errors.append(
                        ErrorRecord(
                            row=r,
                            cell=f"D{r}/E{r}",
                            check_type="calc_text_check",
                            reason="D 산출근거 계산값(ROUND 자리수 반영)과 E 수량 불일치",
                            severity="HIGH",
                            related_formula=f"D:{d_formula_or_text} | E:{e_formula}",
                            actual_value=e_value,
                            expected_value=expected,
                            difference=diff,
                            rule_name="ROUND 비교",
                        )
                    )

        # 2~3) allowance: 퍼센트/표준 룰 중 priority 최대 1개만 적용
        percent_text: Optional[str] = None
        m = percent_regex.search(bigo)
        if m:
            percent_text = f"{m.group(1)}%"

        selected_rule = pick_single_allowance_rule(
            standard_rules=standard_rules,
            allowance_map=allowance_map,
            percent_match_text=percent_text,
            combined_text=combined,
            row_bigo_text=row_bigo,
        )

        if selected_rule and d_numeric is not None and e_value is not None:
            applied_multiplier = float(selected_rule.get("multiplier", 1.0))
            expected = d_numeric * applied_multiplier
            diff = abs(expected - e_value)
            if diff >= threshold:
                errors.append(
                    ErrorRecord(
                        row=r,
                        cell=f"E{r}",
                        check_type="allowance_check",
                        reason="선택된 할증 규칙 적용값과 E 수량이 다름",
                        severity="MEDIUM",
                        related_formula=f"D:{d_formula_or_text} | E:{e_formula}",
                        actual_value=e_value,
                        expected_value=expected,
                        difference=diff,
                        rule_name=str(selected_rule.get("name", "")),
                    )
                )

        # 4) unit_weight
        for sev, reason, rule_name in unit_weight_check(work, spec, unit):
            errors.append(
                ErrorRecord(
                    row=r,
                    cell=f"F{r}",
                    check_type="unit_weight",
                    reason=reason,
                    severity=sev,
                    rule_name=rule_name,
                )
            )

    try:
        build_reports(errors, args.outdir)
    except Exception as e:
        raise SystemExit(f"[ERROR] 리포트 저장 실패: {e}")

    print(f"[OK] sheet='{sheet_name}', header_row={header_row}, rows_checked={ws_formula.max_row - header_row}")
    print(f"[OK] reports saved to: {os.path.join(args.outdir, 'report.csv')} / report.xlsx")


if __name__ == "__main__":
    main()
