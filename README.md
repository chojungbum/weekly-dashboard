# 주간 실적 대시보드 — 자동화 가이드

## 📁 폴더 구조

```
weekly-dashboard/
├── data/
│   └── 실적데이터.xlsx        ← 매주 이 파일만 업데이트!
├── .github/
│   └── workflows/
│       └── update-dashboard.yml  ← 자동 변환 워크플로우
├── convert.py                 ← 엑셀 → JSON 변환 스크립트
├── dashboard_data.json        ← 자동 생성 (직접 수정 불필요)
└── index.html                 ← 대시보드 웹페이지
```

---

## 🚀 매주 업데이트 방법 (3줄!)

```bash
# 1. 엑셀 파일을 data/ 폴더에 덮어쓰기 후

git add data/실적데이터.xlsx
git commit -m "4월 2주차 실적 업데이트"
git push origin main
```

**→ GitHub이 자동으로:**
1. Python 실행해서 엑셀 → JSON 변환
2. `dashboard_data.json` 업데이트
3. 웹사이트 자동 반영 (1~2분 후)

---

## ⚙️ 최초 설정 (딱 한 번만)

### 1. 엑셀 파일 규칙
- 파일명: 반드시 `실적데이터.xlsx`
- 위치: `data/` 폴더 안
- 필수 시트: `매출 DB_2026`, `매출 DB_2025`

### 2. 매출 DB_2026 필수 컬럼
| 컬럼명 | 설명 |
|--------|------|
| 마감월 | 숫자 (1, 2, 3, 4...) |
| 원화판매금액 | 매출 금액 (원) |
| 이익액 | GP 금액 (원) |
| 국내/해외 | '국내' 또는 '해외' |
| 채널그룹 | CIS권-B2B, 올리브영 등 |
| 라인 | 펩타이드 9, 레드 락토 콜라겐 등 |
| 유통구조 | 다이소, 올리브영 등 |
| 실적일자 | 날짜 |
| 베이스품명 | SKU 기준 품명 |

### 3. GitHub Actions 권한 설정
Settings → Actions → General → Workflow permissions
→ **Read and write permissions** 선택 → Save

---

## 🌐 웹사이트 주소

```
https://chojungbum.github.io/weekly-dashboard
```

---

## 🔧 로컬 테스트 방법 (선택사항)

Python이 있는 경우:
```bash
pip install pandas openpyxl
python convert.py
```
이후 `index.html`을 브라우저로 열면 확인 가능.
