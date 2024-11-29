# -*- coding: utf-8 -*-

import csv

csv_file_path = "J:/Lager/EXPORT Local/2024.11.29.csv"

# 파일 읽기 및 구분자 감지
with open(csv_file_path, 'r', encoding="Windows-1252") as file:
    sample = file.read(1024)  # 파일의 일부만 읽음
    sniffer = csv.Sniffer()
    detected_delimiter = sniffer.sniff(sample).delimiter
    print(f"감지된 구분자: '{detected_delimiter}'")

# 파일 내용 확인
with open(csv_file_path, 'r', encoding="Windows-1252") as file:
    print("파일 내용 (첫 5줄):")
    for i, line in enumerate(file):
        print(line.strip())  # 줄바꿈 문자 제거 후 출력
        if i >= 4:  # 처음 5줄만 출력
            break
