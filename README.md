# 🔁 Repetitive_Export_Task_Automation

반복적인 Excel 데이터 가공 및 내보내기 작업을 자동화하는 Python 스크립트 모음

---

## 📋 프로젝트 개요
이 프로젝트는 재고 및 주문 관련 데이터를 매일 자동으로 정리, 필터링, 병합하여 최종 Excel 파일로 저장하는 작업을 자동화합니다.  
Python과 Pandas, OpenPyXL 등 데이터 처리 도구를 활용하며, 반복되는 수작업을 효율적으로 대체합니다.

---

## 🔑 주요 스크립트 설명

### 📄 K_sort_by_brand.py
- `2024 Lagerbestand.xlsx` 재고 목록과 당일 생성된 `*_sort_by_excel.xlsx` 파일을 병합
- 'Artikel'과 'Nummer' 기준으로 데이터를 결합하고 새로운 열(New_Column) 추가
- 시트별로 데이터를 저장하고 최종 Excel 파일(`*_sort_by_excel_ko.xlsx`)로 저장

### 📄 Sort_by_brand.py
- 원본 Excel 파일(`{날짜}.xlsx`)을 불러와 `Lieferanten.Name`(공급자명) 기준으로 정렬
- 공급자별로 데이터를 분리하여 각각의 시트에 저장 (`{날짜}_sort_by_excel.xlsx`)

### 📄 debug.py
- CSV 파일(`{날짜}.csv`)의 구분자와 일부 내용을 확인하는 스크립트
- 데이터 파일의 포맷 및 구분자를 자동으로 감지

### 📄 export_csv.py
- `{날짜}.csv` 파일을 읽어 'Artikel' 컬럼에서 'V'로 시작하는 데이터를 제거
- 필터링된 데이터를 `{날짜}.xlsx` 및 복사 경로에 저장

### 📄 final.py
- CSV 파일 필터링 → Excel 정리 → 공급자별 시트 분리 → 재고 파일 병합까지 자동 처리
- 전체 작업을 한 번에 수행하며, 최종 데이터 파일(`{날짜}_sort_by_excel_ko.xlsx`) 생성

---

## 🛠️ 사용 기술
- **Python**: Pandas, OpenPyXL, CSV, datetime
- **데이터 처리**: Excel/CSV 파일 읽기 및 쓰기, 필터링, 데이터 병합 및 정렬, 시트별 저장
- **자동화**: 일별 파일 이름 생성 및 고유 파일명 처리, 반복작업 자동화 로직

---

## 🗂️ 데이터 파일 경로 예시
- 📂 `C:/Users/verwa/Desktop/Privat/export_inventory/`
- 📂 `J:/Lager/EXPORT Local/`

(환경에 맞게 경로 수정 필요)

---

## 🚀 향후 발전 방향
- 재고 목록 및 파일 자동 탐색 및 처리 기능 추가
- 로그 기록 및 예외 처리 강화
- GUI 또는 웹 기반 인터페이스로 비전문가도 쉽게 사용 가능하도록 개선

---

🔗 [포트폴리오 메인으로 돌아가기](https://github.com/jisuseo/Portfolio)
