import pandas as pd
from datetime import datetime
import os

# 현재 날짜 가져오기 (yyyy.mm.dd 형식)
current_date = datetime.now().strftime("%Y.%m.%d")

# 경로 및 파일 이름 설정
folder_path = "J:/Lager/EXPORT Local/"
csv_file_path = f"{folder_path}{current_date}.csv"
excel_file_base = f"{folder_path}{current_date}.xlsx"

# 다른 폴더에도 복사

folder_path_copy = "C:/Users/verwa/Desktop/Privat/export_inventory/"
excel_file_base_copy = f"{folder_path_copy}{current_date}.xlsx"

# 고유 파일 이름 생성
def generate_unique_filename(base_path):
    """중복되지 않는 파일 이름 생성"""
    if not os.path.exists(base_path):
        return base_path
    counter = 2
    while True:
        new_path = base_path.replace(".xlsx", f" ({counter}).xlsx")
        if not os.path.exists(new_path):
            return new_path
        counter += 1

# 고유한 Excel 파일 경로 생성
excel_file_path = generate_unique_filename(excel_file_base)
excel_file_path_copy = generate_unique_filename(excel_file_base_copy)

# CSV 파일 읽기
try:
    df = pd.read_csv(
        csv_file_path,
        encoding="Windows-1252",  # 파일 인코딩
        sep=";",  # 구분자
        quotechar='"',  # 문자열 감싸는 문자
        low_memory=False,  # 대용량 데이터 처리
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
    filtered_df.to_excel(excel_file_path, index=False, engine="openpyxl")
    
    filtered_df.to_excel(excel_file_path_copy, index=False, engine="openpyxl")
    print(f"필터링된 데이터가 {excel_file_path}와 {excel_file_path_copy}에 저장되었습니다.")

except Exception as e:
    print(f"Excel 저장 중 오류 발생: {e}")
