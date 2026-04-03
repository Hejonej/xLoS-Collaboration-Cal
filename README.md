# xLoS / 동일LoS 분석 도구

Raw 엑셀 파일(전기/당기 xLoS, 동일LoS)을 업로드하면 중복 제거, 승인 기준 계산, LoS별 집계까지 자동 처리하고 완성된 엑셀 파일을 다운로드할 수 있는 사내 웹 앱입니다.

## 주요 기능

- 전기/당기 Raw 데이터 병합 및 UniqueKey 기반 중복 제거
- 승인 기준 Y/N, 관리반영 기준 Y/N 자동 계산 (FY 기간 기반)
- LoS별 건수/금액/Refer EM 집계
- 처리 결과 자동 검증 (건수 보존, 승인 기준 정합성, 음수값 감지 등)
- iOS Liquid Glass 스타일 UI

## 빠른 실행

### 방법 1: run.bat 더블클릭 (추천)
`run.bat`을 더블클릭하고 모드 선택:
- `1` → 내 PC에서만 사용 (localhost)
- `2` → 사내 네트워크 공유 (동료 접속 가능)

### 방법 2: 수동 실행
```powershell
pip install -r requirements.txt
python app.py
```
브라우저에서 http://localhost:5000 접속

### 사내 네트워크 배포
```powershell
pip install -r requirements.txt
python -c "from waitress import serve; from app import app; serve(app, host='0.0.0.0', port=5000, threads=4)"
```
동료는 `http://<이 PC의 IP>:5000`으로 접속 (설치 없이 브라우저만 있으면 OK)

## 프로젝트 구조

```
├── app.py              # Flask 웹 서버
├── processor.py        # 엑셀 처리 + 검증 엔진
├── templates/
│   └── index.html      # Liquid Glass UI
├── static/
│   ├── default_template.xlsx  # 내장 템플릿
│   └── pwc_logo.png
├── requirements.txt
└── run.bat             # 원클릭 실행 스크립트
```
