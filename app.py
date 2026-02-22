from __future__ import annotations

import os
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


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🌿", layout="wide")

    st.title("🌿 qty-audit")
    st.caption("조경 시설물 수량산출서 자동 검토 시스템")

    repo_root = find_repo_root()
    rules_path = repo_root / DEFAULT_RULES_FILE

    uploaded = st.file_uploader("📂 XLSX 파일 업로드", type=["xlsx"])

    run_clicked = st.button("🔍 검토 실행", type="primary", disabled=(uploaded is None))

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

            # ===============================
            # 🔥 결과 전체 표 화면 표시
            # ===============================
            if csv_path.exists():
                st.subheader("📋 검토 결과 전체 목록")

                df = pd.read_csv(csv_path)

                st.dataframe(
                    df,
                    use_container_width=True,
                    height=500
                )

                st.write(f"총 오류 건수: {len(df)}건")

            else:
                st.warning("report.csv가 생성되지 않았습니다.")

            # ===============================
            # 다운로드 버튼
            # ===============================
            col1, col2 = st.columns(2)

            with col1:
                if csv_path.exists():
                    st.download_button(
                        "⬇️ report.csv 다운로드",
                        data=csv_path.read_bytes(),
                        file_name="report.csv",
                        mime="text/csv",
                    )

            with col2:
                if xlsx_report_path.exists():
                    st.download_button(
                        "⬇️ report.xlsx 다운로드",
                        data=xlsx_report_path.read_bytes(),
                        file_name="report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


if __name__ == "__main__":
    main()
