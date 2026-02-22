import os
import subprocess
import tempfile
from datetime import datetime

import streamlit as st

st.set_page_config(page_title="qty-audit", layout="wide")

st.title("🌿 qty-audit")
st.write("조경 수량산출 XLSX 파일을 업로드하면 자동 검토 후 결과를 제공합니다.")

# -------------------------------
# ✅ PyYAML 설치/임포트 진단 블록
# -------------------------------
st.subheader("🧪 환경 진단 (PyYAML)")
try:
    import yaml  # PyYAML이 제공하는 모듈명은 yaml 입니다.
    st.success(f"PyYAML 설치됨: yaml 버전 = {getattr(yaml, '__version__', 'unknown')}")
except Exception as e:
    st.error(f"PyYAML(yaml) import 실패: {e}")

st.divider()

# -------------------------------
# 파일 업로드
# -------------------------------
uploaded_file = st.file_uploader("📂 검토할 XLSX 파일 업로드", type=["xlsx"])

if uploaded_file:
    st.success(f"파일 업로드 완료: {uploaded_file.name}")

    if st.button("🔍 검토 실행", type="primary"):
        with st.spinner("검토 중입니다..."):
            with tempfile.TemporaryDirectory() as tmpdir:

                # 업로드 파일 저장
                input_path = os.path.join(tmpdir, uploaded_file.name)
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # output 폴더 생성
                output_dir = os.path.join(tmpdir, "output")
                os.makedirs(output_dir, exist_ok=True)

                # audit.py 실행
                cmd = [
                    "python",
                    "audit.py",
                    input_path,
                    "--rules",
                    "rules.yml",
                    "--outdir",
                    output_dir,
                ]

                result = subprocess.run(cmd, capture_output=True, text=True)

                if result.returncode != 0:
                    st.error("❌ 검토 중 오류 발생")
                    st.code((result.stdout or "") + "\n" + (result.stderr or ""))
                else:
                    st.success("✅ 검토 완료")

                    report_xlsx = os.path.join(output_dir, "report.xlsx")
                    report_csv = os.path.join(output_dir, "report.csv")

                    st.subheader("📥 결과 다운로드")

                    if os.path.exists(report_xlsx):
                        with open(report_xlsx, "rb") as f:
                            st.download_button(
                                label="report.xlsx 다운로드",
                                data=f,
                                file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )
                    else:
                        st.warning("report.xlsx가 생성되지 않았습니다.")

                    if os.path.exists(report_csv):
                        with open(report_csv, "rb") as f:
                            st.download_button(
                                label="report.csv 다운로드",
                                data=f,
                                file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                mime="text/csv",
                            )
                    else:
                        st.warning("report.csv가 생성되지 않았습니다.")
else:
    st.info("먼저 XLSX 파일을 업로드하세요.")

st.caption("※ 업로드 파일은 서버에 저장되지 않으며 실행 후 자동 삭제됩니다.")
