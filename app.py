from __future__ import annotations

import re
import sys
import subprocess
import tempfile
from pathlib import Path

import streamlit as st
import pandas as pd


APP_TITLE = "qty-audit"
DEFAULT_RULES_FILE = "rules.yml"


def find_repo_root() -> Path:
    return Path(__file__).resolve().parent


def run_audit(repo_root: Path, xlsx_path: Path, rules_path: Path, outdir: Path):
    audit_py = repo_root / "audit.py"

    cmd = [
        sys.executable,
        str(audit_py),
        str(xlsx_path),
        "--rules",
        str(rules_path),
        "--outdir",
        str(outdir),
    ]

    proc = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=str(repo_root),
    )

    logs = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    return proc.returncode, logs


def _cell_to_sortkey(cell: str):
    """
    'E145', 'D12/E12', 'F450' 같은 cell 값을 정렬 가능한 키로 변환
    - 여러 셀 표기(D12/E12)는 앞의 셀(D12)을 기준으로 정렬
    """
    if not isinstance(cell, str) or not cell.strip():
        return (9999, 999999)

    token = cell.split("/")[0].strip().upper()  # 'D12/E12' -> 'D12'
    m = re.match(r"([A-Z]+)(\d+)$", token)
    if not m:
        return (9999, 999999)

    col_letters, row_num = m.group(1), int(m.group(2))

    # 엑셀 컬럼 문자 -> 숫자(A=1, Z=26, AA=27...)
    col_num = 0
    for ch in col_letters:
        col_num = col_num * 26 + (ord(ch) - ord("A") + 1)

    return (col_num, row_num)


def sort_and_group_errors(df: pd.DataFrame) -> pd.DataFrame:
    """
    정렬 규칙:
    1) severity: HIGH -> MEDIUM -> LOW
    2) cell 순서: (컬럼, 행)
    3) row
    """
    sev_order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}

    df = df.copy()

    # column normalize
    if "severity" in df.columns:
        df["severity"] = df["severity"].astype(str).str.strip().str.upper()
    else:
        df["severity"] = ""

    if "cell" in df.columns:
        df["cell"] = df["cell"].astype(str).str.strip()
    else:
        df["cell"] = ""

    if "row" not in df.columns:
        df["row"] = -1

    df["_sev_rank"] = df["severity"].map(sev_order).fillna(99).astype(int)

    cell_keys = df["cell"].map(_cell_to_sortkey)
    df["_cell_col"] = cell_keys.map(lambda x: x[0])
    df["_cell_row"] = cell_keys.map(lambda x: x[1])

    df = df.sort_values(
        by=["_sev_rank", "_cell_col", "_cell_row", "row"],
        ascending=[True, True, True, True],
        kind="mergesort",
    )

    return df.drop(columns=["_sev_rank", "_cell_col", "_cell_row"], errors="ignore")


def show_grouped_errors(df: pd.DataFrame) -> None:
    """HIGH/MEDIUM/LOW 묶어서 표로 출력."""
    df_sorted = sort_and_group_errors(df)

    st.write(f"총 오류 건수: **{len(df_sorted)}건**")

    # 그룹별 표시(접기)
    for sev in ["HIGH", "MEDIUM", "LOW"]:
        g = df_sorted[df_sorted["severity"] == sev]
        if len(g) == 0:
            continue

        with st.expander(f"{sev} ({len(g)}건)", expanded=(sev == "HIGH")):
            st.dataframe(g, use_container_width=True, height=450)


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🌿", layout="wide")

    st.title("🌿 qty-audit")
    st.caption("조경 시설물 수량산출서 자동 검토 시스템")

    repo_root = find_repo_root()
    rules_path = repo_root / DEFAULT_RULES_FILE

    uploaded = st.file_uploader("📂 XLSX 파일 업로드", type=["xlsx"])
    run_clicked = st.button("🔍 검토 실행", type="primary", disabled=(uploaded is None))

    # (선택) 디버그 토글: “2개만 보임” 같은 상황 진단용
    debug = st.toggle("디버그 정보 표시", value=False)

    if run_clicked:
        if uploaded is None:
            st.warning("파일을 먼저 업로드하세요.")
            return

        with tempfile.TemporaryDirectory(prefix="qty_audit_") as tmpdir:
            tmpdir_path = Path(tmpdir)
            input_dir = tmpdir_path / "input"
            output_dir = tmpdir_path / "output"

            input_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)

            xlsx_path = input_dir / uploaded.name
            xlsx_path.write_bytes(uploaded.getvalue())

            with st.spinner("검토 중입니다..."):
                code, logs = run_audit(
                    repo_root=repo_root,
                    xlsx_path=xlsx_path,
                    rules_path=rules_path,
                    outdir=output_dir,
                )

            if code != 0:
                st.error("❌ 검토 중 오류 발생")
                with st.expander("📜 실행 로그"):
                    st.code(logs)
                return

            st.success("✅ 검토 완료")

            csv_path = output_dir / "report.csv"
            xlsx_report_path = output_dir / "report.xlsx"

            if debug:
                st.info(f"repo_root={repo_root}")
                st.info(f"rules_path={rules_path} (exists={rules_path.exists()})")
                st.info(f"csv_path={csv_path} (exists={csv_path.exists()})")
                st.info(f"xlsx_report_path={xlsx_report_path} (exists={xlsx_report_path.exists()})")
                if csv_path.exists():
                    st.info(f"report.csv size={csv_path.stat().st_size} bytes")
                if xlsx_report_path.exists():
                    st.info(f"report.xlsx size={xlsx_report_path.stat().st_size} bytes")
                with st.expander("📜 실행 로그(성공 케이스)"):
                    st.code(logs)

            # ===============================
            # 결과 표 화면 표시 (정렬/그룹)
            # ===============================
            if csv_path.exists():
                st.subheader("📋 검토 결과 (중요도별)")

                # utf-8-sig로 저장하므로 여기서도 동일하게 읽기(환경에 따라 깨짐 방지)
                df = pd.read_csv(csv_path, encoding="utf-8-sig")

                if debug:
                    st.info(f"df rows={len(df)} / columns={list(df.columns)}")

                if len(df) == 0:
                    st.warning("오류가 0건입니다. (report.csv는 생성되었으나 내용이 비어있음)")
                else:
                    show_grouped_errors(df)
            else:
                st.warning("report.csv가 생성되지 않았습니다. (audit.py가 report.csv를 저장하는지 확인 필요)")

            # ===============================
            # 다운로드 버튼
            # ===============================
            st.divider()
            st.subheader("⬇️ 결과 다운로드")

            col1, col2 = st.columns(2)

            with col1:
                if csv_path.exists():
                    st.download_button(
                        "⬇️ report.csv 다운로드",
                        data=csv_path.read_bytes(),
                        file_name="report.csv",
                        mime="text/csv",
                    )
                else:
                    st.caption("report.csv 없음")

            with col2:
                if xlsx_report_path.exists():
                    st.download_button(
                        "⬇️ report.xlsx 다운로드",
                        data=xlsx_report_path.read_bytes(),
                        file_name="report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.caption("report.xlsx 없음")


if __name__ == "__main__":
    main()
