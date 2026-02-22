# qty-audit

조경 시설물 수량산출서(XLSX)를 자동 점검하는 도구입니다.

## 기능

- 할증률 적용 여부 확인
- 할증 기준 적합성 검토
- 산출근거 계산값과 수량 결과값 비교
- 단위중량 적용 검토

---

## 파일 구성

- `rules.yml` : 할증/계산/단위중량 검토 룰
- `audit.py` : 감사 실행 스크립트

---

## 실행 환경

이 저장소는 **Google Colab + Google Drive 환경에서 실행**을 전제로 합니다.

- 입력 파일 : Drive의 `qty-audit/input`
- 출력 파일 : Drive의 `qty-audit/output`

---

## Colab 실행 절차

1. Drive의 `qty-audit/input` 폴더에 검토할 `.xlsx` 파일 업로드
2. Colab 노트북 실행
3. 최신 파일 자동 선택 후 감사 실행
4. 결과는 `output/report.xlsx` 생성

---

## 출력 결과

- `report.xlsx`
  - `Summary` : 오류 건수 집계
  - `Errors` : 오류 상세 (행 / 열 / 원인 / severity / 관련수식 / 차이)

- `report.csv` : 오류 상세 목록

---

## 버전 관리

- 코드 및 룰은 GitHub에서 관리
- 입력 및 결과 파일은 Google Drive에서 관리
