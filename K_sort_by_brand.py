import pandas as pd
from datetime import datetime

# 경로 설정
folder_path = "C:/Users/verwa/Desktop/Privat/export_inventory/"
lagerbestand_file = f"{folder_path}2024 Lagerbestand.xlsx"
current_date = datetime.now().strftime("%Y.%m.%d")
input_file = f"{folder_path}{current_date}_sort_by_excel.xlsx"
output_file_path = f"{folder_path}{current_date}_sort_by_excel_ko.xlsx"

# 1. 2024 Lagerbestand 파일 읽기
try:
    lager_df = pd.read_excel(lagerbestand_file, engine="openpyxl")
    print("Lagerbestand 파일의 열 이름:", lager_df.columns.tolist())

    # 공백 제거 및 열 이름 정리
    lager_df.columns = lager_df.columns.str.strip()

    # 확인한 열 이름으로 접근
    if "K_Name" not in lager_df.columns:
        raise KeyError("'K_Name' 열이 존재하지 않습니다.")
except Exception as e:
    print(f"파일 읽기 오류: {e}")
    exit()

# 2. 이전 단계에서 생성된 Excel 파일 읽기
try:
    sort_by_excel_df = pd.read_excel(input_file, sheet_name=None, engine="openpyxl")
    print(f"이전 생성 파일 {input_file}을 성공적으로 읽었습니다.")
except FileNotFoundError:
    print(f"파일 {input_file}을 찾을 수 없습니다.")
    exit()

# 3. 각 Sheet 처리 및 결과 생성
result_data = {}

for sheet_name, df in sort_by_excel_df.items():
    # Artikel과 Nummer 데이터 타입 통일 (문자열 변환)
    df["Artikel"] = df["Artikel"].astype(str)
    lager_df["K_Name"] = lager_df["K_Name"].astype(str)

    # Artikel과 K_Name 병합
    merged_df = df.merge(
        lager_df[["Nummer", "K_Name"]].rename(columns={"K_Name": "New_Column"}),  # 필요한 열 추가
        how="left",
        left_on="Artikel",
        right_on="Nummer"
    )

    # A열과 B열 사이에 New_Column 삽입
    col_order = ["Artikel", "New_Column"] + [col for col in merged_df.columns if col not in ["Artikel", "New_Column"]]
    merged_df = merged_df[col_order]

    # 결과 저장
    result_data[sheet_name] = merged_df

# 4. 새로운 Excel 파일로 저장
try:
    with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
        for sheet_name, data in result_data.items():
            data.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"새로운 파일이 {output_file_path}에 저장되었습니다.")
except Exception as e:
    print(f"Excel 파일 저장 중 오류 발생: {e}")
