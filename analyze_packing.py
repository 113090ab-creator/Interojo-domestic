import pandas as pd
import numpy as np
from datetime import datetime, timedelta

print("="*70)
print("포장실적 상세 분석")
print("="*70)

df_pack = pd.read_excel('포장설비투입현황.xlsx', sheet_name='Sheet1')

print(f'\n[ 기본 정보 ]')
print(f'총 포장실적 건수: {len(df_pack):,}')
print(f'데이터 기간: {df_pack["일마감"].min()} ~ {df_pack["일마감"].max()}')

print(f'\n[ 물량 현황 ]')
print(f'생산수량 총합: {df_pack["생산수량"].sum():,.0f}')
print(f'팩수량 총합: {df_pack["팩수량"].sum():,.0f}')
print(f'낱개수량 총합: {df_pack["낱개수량"].sum():,.0f}')

print(f'\n[ 포장설비별 현황 ]')
pack_equip = df_pack.groupby('포장설비').agg({
    '생산수량': 'sum',
    '팩수량': 'sum',
    '낱개수량': 'sum',
    '포장설비': 'count'
}).rename(columns={'포장설비': '건수'})
print(pack_equip)

print(f'\n[ 구분별 현황 ]')
df_pack_group = df_pack.groupby('구분').agg({
    '생산수량': 'sum',
    '팩수량': 'sum',
    '낱개수량': 'sum',
    '구분': 'count'
}).rename(columns={'구분': '건수'})
print(df_pack_group)

print(f'\n[ 주/야 구분 ]')
shift_group = df_pack.groupby('주/야').agg({
    '생산수량': 'sum',
    '팩수량': 'sum',
    '주/야': 'count'
}).rename(columns={'주/야': '건수'})
print(shift_group)

print(f'\n[ 파손/불량 현황 ]')
print(f'파손/불량 합계: {df_pack["파손/불량"].sum():,.0f}')
print(f'파손/불량 건수: {(df_pack["파손/불량"] > 0).sum()}')

print(f'\n[ 일별 포장 실적 ]')
daily_pack = df_pack.groupby('일마감').agg({
    '생산수량': 'sum',
    '팩수량': 'sum',
    '일마감': 'count'
}).rename(columns={'일마감': '건수'})
print(daily_pack.tail(10))

print(f'\n[ 생산코드별 실적 ]')
code_pack = df_pack.groupby('생산코드').agg({
    '생산수량': 'sum',
    '팩수량': 'sum',
    '생산코드': 'count'
}).rename(columns={'생산코드': '건수'}).sort_values('생산수량', ascending=False)
print(code_pack.head(10))

print(f'\n[ 모델별 TOP10 ]')
model_pack = df_pack.groupby('모델코드').agg({
    '생산수량': 'sum',
    '모델코드': 'count'
}).rename(columns={'모델코드': '건수'}).sort_values('생산수량', ascending=False)
print(model_pack.head(10))
