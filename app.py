# app.py
# Streamlit UI for qty-audit (XLSX 업로드 → audit.py 실행 → report.csv/xlsx 다운로드)
#
# 핵심 포인트
# - audit.py는 "같은 파이썬 환경"에서 실행되어야 합니다 → sys.executable 사용
# - PyYAML(yaml) 설치/임포트 상태를 화면에서 즉시 진단
# - 업로드 파일은 서버에 영구 저장하지 않고, 실행 후 임시폴더에서 결과만 제공

from __future__ import annotations

import os
import sys
import shutil
import subprocess
import tempfile
from pathlib import Path

import streamlit as st

APP_TITLE = "qty-audit"
DEFAULT_RULES_FILE = "rules.yml"  # repo 루트에 rules.yml이 있다고 가정


def find_repo_root() -> Path:
    # Streamlit Cloud 기준: 현재 작업 디렉토리가 repo 루트인 경우가 많음
    # 안전하게 app.py 위치 기준으로도 확인
    here = Path(__file__).resolve().parent
    # app.py가 루트에 있다면 here가 루트
    return here


def run_audit(
    repo_root: Path,
    xlsx_path: Path,
    rules_path: Path,
    outdir: Path,
) -> tuple[int, str]:
    """
    audit.py를 sys.executable로 실행하여
    Streamlit Cloud의 '다른 python' 문제를 피합니다.
    """
    audit_py = repo_root / "audit.py"
    if not audit_py.exists():
        raise FileNotFoundError(f"audit.py를 찾을 수 없습니다: {audit_py}")

    cmd = [
        sys.executable,  # ★중요: 현재 앱이 돌고 있는 python으로 실행
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
        cwd=str(repo_root),  # audit.py가 rules.yml을 상대경로로 찾는 경우 대비
    )

    combined = ""
    if proc.stdout:
        combined += proc.stdout
    if proc.stderr:
        combined += ("\n" if combined else "") + proc.stderr

    return proc.returncode, combined


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🌿", layout="centered")

    st.title("🌿 qty-audit")
    st.caption("조경 수량산출 XLSX 파일을 업로드하면 자동 검토 후 결과(report.csv, report.xlsx)를 제공합니다.")

    repo_root = find_repo_root()
    rules_path = repo_root / DEFAULT_RULES_FILE

    # ---- 환경 진단: PyYAML ----
    st.subheader("🧪 환경 진단 (PyYAML)")
    try:
        import yaml  # noqa

        st.success(f"PyYAML 설치됨: yaml 버전 = {getattr(yaml, '__version__', 'unknown')}")
    except Exception as e:
        st.error(f"PyYAML(yaml) import 실패: {e}")
        st.info("Streamlit Cloud에서는 requirements.txt에 `pyyaml`(또는 `PyYAML`)이 포함되어 있어야 합니다.")

    # ---- 파일/설정 ----
    st.subheader("📄 검토할 XLSX 파일 업로드")
    uploaded = st.file_uploader("XLSX 파일을 선택하세요", type=["xlsx"])

    with st.expander("⚙️ 설정(기본값 권장)", expanded=False):
        st.write("rules.yml 경로와 출력 폴더 이름을 조정할 수 있습니다.")
        rules_input = st.text_input("rules.yml 경로", value=str(rules_path))
        outdir_name = st.text_input("출력 폴더 이름", value="output")

    # rules 경로 확정
    rules_path = Path(rules_input).expanduser()
    if not rules_path.is_absolute():
        # 상대경로면 repo_root 기준으로 해석
        rules_path = (repo_root / rules_path).resolve()

    if not rules_path.exists():
        st.warning(f"rules.yml이 보이지 않습니다: {rules_path}\n\nrepo에 rules.yml이 있는지 확인하세요.")

    st.divider()

    # ---- 실행 ----
    run_clicked = st.button("🔍 검토 실행", type="primary", disabled=(uploaded is None))

    if run_clicked:
        if uploaded is None:
            st.warning("먼저 XLSX 파일을 업로드하세요.")
            return

        if not rules_path.exists():
            st.error(f"rules.yml이 없어 실행할 수 없습니다: {rules_path}")
            return

        # 임시 작업 폴더 (실행 후 자동 삭제)
        with tempfile.TemporaryDirectory(prefix="qty_audit_") as tmpdir:
            tmpdir_path = Path(tmpdir)
            input_dir = tmpdir_path / "input"
            output_dir = tmpdir_path / outdir_name
            input_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)

            # 업로드 파일 저장
            xlsx_path = input_dir / uploaded.name
            xlsx_path.write_bytes(uploaded.getvalue())

            st.info(f"업로드 완료: {uploaded.name}")
            st.write(f"- 입력 파일: `{xlsx_path}`")
            st.write(f"- rules: `{rules_path}`")
            st.write(f"- 출력 폴더: `{output_dir}`")

            with st.spinner("검토 중..."):
                try:
                    code, logs = run_audit(
                        repo_root=repo_root,
                        xlsx_path=xlsx_path,
                        rules_path=rules_path,
                        outdir=output_dir,
                    )
                except Exception as e:
                    st.error(f"실행 중 예외 발생: {e}")
                    st.stop()

            # 로그 출력
            with st.expander("📜 실행 로그(디버깅)", expanded=(code != 0)):
                st.code(logs or "(로그 없음)", language="text")

            if code != 0:
                st.error("검토 중 오류가 발생했습니다. 위 로그에서 원인을 확인하세요.")
                st.stop()

            # 결과 파일 찾기
            csv_path = output_dir / "report.csv"
            xlsx_report_path = output_dir / "report.xlsx"

            st.success("검토 완료!")

            # 다운로드 버튼
            cols = st.columns(2)
            with cols[0]:
                if csv_path.exists():
                    st.download_button(
                        label="⬇️ report.csv 다운로드",
                        data=csv_path.read_bytes(),
                        file_name="report.csv",
                        mime="text/csv",
                    )
                else:
                    st.warning("report.csv가 생성되지 않았습니다.")

            with cols[1]:
                if xlsx_report_path.exists():
                    st.download_button(
                        label="⬇️ report.xlsx 다운로드",
                        data=xlsx_report_path.read_bytes(),
                        file_name="report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.warning("report.xlsx가 생성되지 않았습니다.")

            st.caption("※ 업로드 파일은 서버에 영구 저장되지 않으며, 실행이 끝나면 임시 폴더가 삭제됩니다.")


if __name__ == "__main__":
    main()
