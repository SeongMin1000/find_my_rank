import openpyxl
import pandas as pd

file_path = 'sys.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook['Sheet1']

data_list = []

# 이름, 중간, 기말, 가중치 추출
for row in range(16, 59):
    name=sheet[f'E{row}'].value
    mid=sheet[f'F{row}'].value
    end=sheet[f'G{row}'].value
    val=(mid*30+end*35)*0.01
    
    data_dict = {
        'name': name,
        'mid': mid,
        'end': end,
        'val': val
    }
    data_list.append(data_dict)

# 가중치 기준 내림차순 정렬
sorted_data = sorted(data_list, key=lambda x: x['val'], reverse=True) 
# 석차 기록
rank = 1
for data in sorted_data:
    data['rank']= rank 
    rank+=1

name= input("이름 : ")
# 이름 검색
found = False
for data in sorted_data:
    if data['name'] == name:
        found = True
        print(f"이름: {data['name']}")
        print(f"중간고사 점수: {data['mid']}")
        print(f"기말고사 점수: {data['end']}")
        print(f"가중치 점수: {data['val']:.2f}")
        print(f"석차: {data['rank']}")
        break

if not found:
    print("입력한 이름을 찾을 수 없습니다.")
