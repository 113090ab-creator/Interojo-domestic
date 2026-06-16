import pandas as pd
import numpy as np

excel_file = '생산요청등록(국내)_5월.xlsx'
df = pd.read_excel(excel_file, sheet_name='Sheet1')

print('=== 상태별 현황 ===')
status_counts = df['상태'].value_counts()
print(status_counts)
print()

print('=== 상태별 물량(수량-PACK) ===')
status_qty = df.groupby('상태')['수량(PACK)'].sum()
print(status_qty)
print()

print('=== 구분별 현황 ===')
category_counts = df['구분'].value_counts()
print(category_counts)
print()

print('=== 납기일자 범위 ===')
print(f'최소: {df["납기일자"].min()}')
print(f'최대: {df["납기일자"].max()}')
print()

print('=== 전체 통계 ===')
print(f'총 요청건수: {len(df)}')
print(f'총 요청물량(PACK): {df["수량(PACK)"].sum():,.0f}')
print(f'총 요청물량(PCS): {df["수량(PCS)"].sum():,.0f}')
print()

print('=== 컬럼 전체 확인 ===')
print(df.columns.tolist())
