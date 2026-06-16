import pandas as pd
import numpy as np

print("="*60)
print("1. 생산요청등록(국내)_5월.xlsx 분석")
print("="*60)

excel_file = '생산요청등록(국내)_5월.xlsx'
df_request = pd.read_excel(excel_file, sheet_name='Sheet1')

print(f'\n총 요청건수: {len(df_request):,}')
print(f'총 요청물량(PACK): {df_request["수량(PACK)"].sum():,.0f}')
print(f'\n상태 현황:')
print(df_request['상태'].value_counts())

print(f'\n납기일자 범위: {df_request["납기일자"].min()} ~ {df_request["납기일자"].max()}')

print(f'\n구분: {df_request["구분"].unique()}')

print(f'\n주요 컬럼:')
print(df_request.columns.tolist())

print("\n"+"="*60)
print("2. 포장설비투입현황.xlsx 분석")
print("="*60)

packing_file = '포장설비투입현황.xlsx'
xls_packing = pd.ExcelFile(packing_file)
print(f'\nSheet names: {xls_packing.sheet_names}')

for sheet_name in xls_packing.sheet_names:
    df_packing = pd.read_excel(packing_file, sheet_name=sheet_name)
    print(f'\n[{sheet_name}] 데이터 수: {len(df_packing):,}, 컬럼 수: {len(df_packing.columns)}')
    print(f'컬럼: {list(df_packing.columns)}')
    print(f'\n샘플 데이터:')
    print(df_packing.head(3))
    
    if '수량' in df_packing.columns or '투입' in ' '.join(df_packing.columns):
        qty_cols = [col for col in df_packing.columns if '수량' in col or '투입' in col or 'QTY' in col]
        if qty_cols:
            print(f'\n물량 관련 컬럼 통계:')
            for col in qty_cols:
                print(f'{col}: {df_packing[col].sum():,.0f}')

print("\n"+"="*60)
print("3. 데이터 통합 분석")
print("="*60)

# 날짜 기준 분석
print(f'\n생산요청 납기일: {df_request["납기일자"].min().date()} ~ {df_request["납기일자"].max().date()}')

# 포장 데이터도 날짜가 있는지 확인
for sheet_name in xls_packing.sheet_names:
    df_packing = pd.read_excel(packing_file, sheet_name=sheet_name)
    date_cols = [col for col in df_packing.columns if '날짜' in col or '일' in col or '일자' in col]
    if date_cols:
        print(f'\n[{sheet_name}] 날짜 컬럼: {date_cols}')
