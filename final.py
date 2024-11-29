import pandas as pd
from datetime import datetime
import os

# 현재 날짜 가져오기
current_date = datetime.now().strftime("%Y.%m.%d")

# 경로 설정
csv_folder_path = "J:/Lager/EXPORT Local/"
csv_file_path = f"{csv_folder_path}{current_date}.csv"
excel_file_base = f"{csv_folder_path}{current_date}.xlsx"
excel_file_base_copy = f"C:/Users/verwa/Desktop/Privat/export_inventory/{current_date}.xlsx"

# Lagerbestand 관련 경로 설정
lagerbestand_folder_path = "C:/Users/verwa/Desktop/Privat/export_inventory/"
lagerbestand_file = f"{lagerbestand_folder_path}2024 Lagerbestand.xlsx"
output_file_path_sort = f"{lagerbestand_folder_path}{current_date}_sort_by_excel.xlsx"
output_file_path_final = f"{lagerbestand_folder_path}{current_date}_sort_by_excel_ko.xlsx"

# 고유 파일 이름 생성 함수
def generate_unique_filename(base_path):
    if not os.path.exists(base_path):
        return base_path
    counter = 2
    while True:
        new_path = base_path.replace(".xlsx", f" ({counter}).xlsx")
        if not os.path.exists(new_path):
            return new_path
        counter += 1

# 1. CSV 파일 읽기 및 필터링
try:
    df = pd.read_csv(
        csv_file_path,
        encoding="Windows-1252",
        sep=";",
        quotechar='"',
        low_memory=False
    )
    print(f"CSV 데이터를 성공적으로 읽어왔습니다. 열 이름: {df.columns.tolist()}")
except Exception as e:
    print(f"CSV 파일 읽기 오류: {e}")
    exit()

# 'Artikel' 열에서 'V'로 시작하는 데이터 제외
if 'Artikel' in df.columns:
    filtered_df = df[~df['Artikel'].str.startswith('V', na=False)]
    print("필터링 성공! 결과 데이터프레임 준비 완료.")
else:
    print("'Artikel' 열이 존재하지 않습니다. CSV 파일의 열 이름을 확인하세요.")
    print(f"현재 열 이름: {df.columns.tolist()}")
    exit()

# 필터링된 데이터를 Excel로 저장
try:
    excel_file_path = generate_unique_filename(excel_file_base)
    excel_file_path_copy = generate_unique_filename(excel_file_base_copy)
    filtered_df.to_excel(excel_file_path, index=False, engine="openpyxl")
    filtered_df.to_excel(excel_file_path_copy, index=False, engine="openpyxl")
    print(f"필터링된 데이터가 {excel_file_path}와 {excel_file_path_copy}에 저장되었습니다.")
except Exception as e:
    print(f"Excel 저장 중 오류 발생: {e}")
    exit()

# 2. Excel 데이터 정리
try:
    df = pd.read_excel(excel_file_path_copy, engine="openpyxl")
    print(f"Excel 파일 {excel_file_path_copy}을 성공적으로 읽었습니다.")
except FileNotFoundError:
    print(f"파일 {excel_file_path_copy}을 찾을 수 없습니다.")
    exit()

# KolliBestand가 null인 행 삭제
df = df.dropna(subset=["KolliBestand"])

# 필요 없는 열 삭제
columns_to_remove = ["NettoVerfügbar", "NettoBestand", "NettoBestellt", "NettoEingeliefert", "NettoReserviert"]
df = df.drop(columns=columns_to_remove, errors="ignore")

# Lieferanten.Name 열 기준으로 데이터 정렬
df = df.sort_values(by=["Lieferanten.Name"])

# Lieferanten.Name에 따라 Sheet를 생성
try:
    output_file_path_sort = generate_unique_filename(output_file_path_sort)
    with pd.ExcelWriter(output_file_path_sort, engine="openpyxl") as writer:
        unique_lieferanten = df["Lieferanten.Name"].dropna().unique()
        for lieferant in unique_lieferanten:
            sheet_data = df[df["Lieferanten.Name"] == lieferant]
            sheet_name = str(lieferant)[:30]
            sheet_data.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"정리된 데이터가 {output_file_path_sort}에 저장되었습니다.")
except Exception as e:
    print(f"Excel 정리 중 오류 발생: {e}")
    exit()

# 3. Lagerbestand와 병합하여 최종 파일 생성
try:
    lager_df = pd.read_excel(lagerbestand_file, engine="openpyxl")
    print("Lagerbestand 파일의 열 이름:", lager_df.columns.tolist())
    lager_df.columns = lager_df.columns.str.strip()
    if "K_Name" not in lager_df.columns:
        raise KeyError("'K_Name' 열이 존재하지 않습니다.")
except Exception as e:
    print(f"파일 읽기 오류: {e}")
    exit()

try:
    sort_by_excel_df = pd.read_excel(output_file_path_sort, sheet_name=None, engine="openpyxl")
    print(f"정리된 Excel 파일 {output_file_path_sort}을 성공적으로 읽었습니다.")
except FileNotFoundError:
    print(f"파일 {output_file_path_sort}을 찾을 수 없습니다.")
    exit()

# 각 Sheet 처리 및 병합
result_data = {}
for sheet_name, df in sort_by_excel_df.items():
    df["Artikel"] = df["Artikel"].astype(str)
    lager_df["K_Name"] = lager_df["K_Name"].astype(str)

    # Artikel과 K_Name 병합
    merged_df = df.merge(
        lager_df[["Nummer", "K_Name"]].rename(columns={"K_Name": "New_Column"}),
        how="left",
        left_on="Artikel",
        right_on="Nummer"
    )

    # A열과 B열 사이에 New_Column 삽입
    col_order = ["Artikel", "New_Column"] + [col for col in merged_df.columns if col not in ["Artikel", "New_Column"]]
    merged_df = merged_df[col_order]

    result_data[sheet_name] = merged_df

# 최종 Excel 파일로 저장
try:
    output_file_path_final = generate_unique_filename(output_file_path_final)
    with pd.ExcelWriter(output_file_path_final, engine="openpyxl") as writer:
        for sheet_name, data in result_data.items():
            data.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"최종 데이터가 {output_file_path_final}에 저장되었습니다.")
except Exception as e:
    print(f"최종 Excel 저장 중 오류 발생: {e}")
