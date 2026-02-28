qty-audit

조경 시설물 수량산출서(XLSX)를 자동 검토하는 Streamlit 기반 검증 도구입니다.

산출근거(D열) 계산값과 수량(E열)을 비교하고, 비고란의 할증(%) 적용 적정성을 자동 점검합니다.

⸻

주요 기능

1️⃣ 산출근거 계산 검증 (calc_text_check)
	•	D열(산출근거)이 숫자 계산식일 경우 직접 계산
	•	E열 수량과 비교
	•	ROUND 자릿수 기준 적용
	•	최소 허용오차(tol) = 0.01

검증 예시:

D: 11.15 * 0.45 * 18.05
E: =ROUND(11.15 * 0.45 * 18.05, 3)

→ 계산값과 수량이 허용오차를 초과하면 HIGH 오류 발생

⸻

2️⃣ 할증 검증 (allowance_check)
	•	비고란(BIGO)에 %가 있을 때만 검증 수행
	•	rules.yml에 정의된 multiplier 기준 사용
	•	설치품 공종은 할증 검증 대상에서 제외

정책:
	•	설치품 → 할증 검증 제외
	•	재료 항목 → D × multiplier 후 E와 비교
	•	D에 이미 *1.04 등 계수 포함 시 이중할증 방지

⸻

3️⃣ 설치품 공종 처리 정책

다음 공종은 할증 검증 제외 대상:
	•	혼합골재포설및다짐
	•	레미콘타설
	•	통석놓기
	•	철근가공조립
	•	잡철물제작및설치
	•	목재가공 및 설치
	•	플랜터 설치
	•	우레탄도장
	•	석재판석붙임
	•	친환경스테인도장
	•	데크깔기

⸻

4️⃣ 허용오차 기준

ROUND 자릿수 기반으로 허용오차 계산:

tol = max(2 × 10^(-round_digits), 0.01)

예:

ROUND 자릿수	계산 허용오차
3자리	0.01 이상
2자리	0.01 이상
0자리	1.0 이상


⸻

5️⃣ 오류 등급

Severity	설명
HIGH	계산 불일치 (산출근거 vs 수량)
MEDIUM	할증 적용 불일치
LOW	데이터 부족 또는 계산 불가

Streamlit 화면에서는 중요도(HIGH → MEDIUM → LOW) 순으로 그룹 표시됩니다.

⸻

📂 프로젝트 구조

qty-audit-streamlit/
├── app.py            # Streamlit UI
├── audit.py          # 검증 엔진
├── rules.yml         # 할증 및 정책 설정
├── requirements.txt
└── README.md


⸻

⚙️ 실행 방법

1️⃣ 로컬 실행

pip install -r requirements.txt
streamlit run app.py

브라우저에서 XLSX 파일 업로드 후 “검토 실행” 클릭

⸻

2️⃣ CLI 직접 실행

python audit.py sample.xlsx --rules rules.yml --outdir output

출력:

output/report.csv
output/report.xlsx


⸻

📝 rules.yml 예시

round_default_digits: 3

allowance_percent_extract_regex: "(\\d+(\\.\\d+)?)%"

allowance_multiplier_map:
  "4%": 1.04
  "5%": 1.05
  "10%": 1.10


⸻

기술적 특징
	•	openpyxl 기반 XLSX 처리
	•	data_only=True 사용
	•	셀참조 포함 산출근거는 계산 제한
	•	일부 ROUND 패턴 직접 평가 지원
	•	자동 시트 탐지 및 헤더 자동 탐지

⸻

한계 사항
	•	openpyxl은 엑셀 수식을 완전 계산하지 않음
	•	셀참조가 복잡한 경우 일부 검증 제한
	•	엑셀 파일은 저장 시 계산값 포함 상태여야 정확한 검증 가능
