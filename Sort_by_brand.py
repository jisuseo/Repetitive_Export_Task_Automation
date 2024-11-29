import pandas as pd
from datetime import datetime
import os


# 현재 날짜 가져오기 (yyyy.mm.dd 형식)
current_date = datetime.now().strftime("%Y.%m.%d")

# 기존 Excel 파일 경로 설정
folder_path = "C:/Users/verwa/Desktop/Privat/export_inventory/"
current_date = datetime.now().strftime("%Y.%m.%d")
input_file_path = f"{folder_path}{current_date}.xlsx"

# 새로 생성될 Excel 파일 이름 설정
output_file_path = f"{folder_path}{current_date}_sort_by_excel.xlsx"

# Excel 파일 읽기
try:
    df = pd.read_excel(input_file_path, engine="openpyxl")
    print(f"Excel 파일 {input_file_path}을 성공적으로 읽었습니다.")
except FileNotFoundError:
    print(f"파일 {input_file_path}을 찾을 수 없습니다.")
    exit()

# 1. KolliBestand가 null인 행 삭제
df = df.dropna(subset=["KolliBestand"])

# 2. 필요 없는 열 삭제
columns_to_remove = ["NettoVerfügbar", "NettoBestand", "NettoBestellt", "NettoEingeliefert", "NettoReserviert"]
df = df.drop(columns=columns_to_remove, errors="ignore")

# 3. Lieferanten.Name 열을 기준으로 데이터 정렬
df = df.sort_values(by=["Lieferanten.Name"])

# 4. Lieferanten.Name 값에 따라 Sheet를 생성하여 새로운 Excel 파일로 저장
try:
    with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
        unique_lieferanten = df["Lieferanten.Name"].dropna().unique()
        for lieferant in unique_lieferanten:
            sheet_data = df[df["Lieferanten.Name"] == lieferant]
            sheet_name = str(lieferant)[:30]  # Excel에서 Sheet 이름은 31자 제한
            sheet_data.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"새로운 Excel 파일이 {output_file_path}에 저장되었습니다.")
except Exception as e:
    print(f"Excel 파일 생성 중 오류 발생: {e}")
