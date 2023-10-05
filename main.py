#############################################     라이브러리 호출     ###################################################
import gspread as gs
import pandas as pd
import openpyxl
from datetime import datetime

##############################################     필요 함수 정의     ###################################################
# 시트 호출 함수 정의
def load_sheets(url, sheet):
    df_sheet = url.worksheet(sheet).get_all_values()
    df_return = pd.DataFrame(df_sheet[1:], columns=df_sheet[0])
    return df_return

###########################################     입과입력 과정코드 입력     ################################################
# 입과입력 양식 제작을 위한 과정코드 입력
input_month = input('해당월을 숫자로 입력하세요 : ')

#######################################     2023 교육관리 데이터베이스 호출     ############################################
gc = gs.service_account(filename='clever-seat-398104-4fa3ad6684b5.json')
sheets = gc.open_by_url('https://docs.google.com/spreadsheets/d/1KlPcoBQGG0Xr2lRkrQkZukdP7uCpmibBEt7tyJwsCcg/edit#gid=0')
# 시트호출 (영업가족직원관리)
df_fa = load_sheets(sheets, "fa")
# 시트호출 (영업가족관리)
df_branch = load_sheets(sheets, "branch")

# 컬럼명 변경 (영업가족CD -> 영업가족코드)
df_fa.rename(columns={'영업가족CD':'영업가족코드'}, inplace=True)
# 컬럼 재선택
df_fa = df_fa[['성명','사원번호','영업가족코드']]
'''
    # 데이터 삭제 (입문과정 미수료)
    df_fa = df_fa.drop(df_fa[(df_fa.iloc[:,5] == None) & (df_fa.iloc[:,6] == None)].index)
    # 데이터 삭제 (등록여부: 미등록, 손보말소, 생보말소, 손생말소)
    df_fa = df_fa.drop(df_fa[(df_fa.iloc[:,4] != '생보등록') & (df_fa.iloc[:,4] != '손보등록') & (df_fa.iloc[:,4] != '손생등록')].index)
    # 입사일자 및 퇴사일자 데이터 시간 형식으로 변경
    df_fa['입사일자(사원)'] = pd.to_datetime(df_fa['입사일자(사원)'])
    df_fa['퇴사일자(사원)'] = pd.to_datetime(df_fa['퇴사일자(사원)'])
'''
# df_attend: 컬럼 추가 및 데이터 삽입 (입사연차)
df_fa['입사연차'] = (datetime.now().year%100 + 1 - df_fa['사원번호'].astype(str).str[:2].astype(int, errors='ignore')).apply(lambda x: f'{x}년차')

# 컬럼명 변경 (소속부서 -> 파트너)
df_branch.rename(columns={'소속부서':'파트너'}, inplace=True)
# 컬럼 재선택
df_branch = df_branch[['영업가족코드','파트너','영업가족명','코드구분']]
# 컬럼 생성 (소속)
df_branch.insert(loc=1, column='소속부문', value=None)
df_branch.insert(loc=2, column='소속총괄', value=None)
df_branch.insert(loc=3, column='소속부서', value=None)
# 데이터 정리 (소속)
for modify in range(df_branch.shape[0] -1,-1,-1):
    split = df_branch.iloc[modify,4].split(">")
    if len(split) > 4:
        df_branch.iloc[modify,1] = df_branch.iloc[modify,4].split(">")[2]
        df_branch.iloc[modify,2] = df_branch.iloc[modify,4].split(">")[3]
        df_branch.iloc[modify,3] = df_branch.iloc[modify,4].split(">")[4]
    elif split[1] == '다이렉트부문총괄':
        if len(split) > 3:
            df_branch.iloc[modify,1] = df_branch.iloc[modify,4].split(">")[2]
            df_branch.iloc[modify,2] = df_branch.iloc[modify,4].split(">")[3]
            df_branch.iloc[modify,3] = df_branch.iloc[modify,4].split(">")[3]
        else:
            df_branch.iloc[modify,1] = df_branch.iloc[modify,4].split(">")[2]
            df_branch.iloc[modify,2] = df_branch.iloc[modify,4].split(">")[2]
            df_branch.iloc[modify,3] = df_branch.iloc[modify,4].split(">")[2]
    else:
        df_branch.drop([modify], inplace=True)
# 컬럼 삭제 (기존 소속)
df_branch = df_branch.drop(columns='파트너')

# dataframe 병합 (df_fa + df_branch)
df_merge = pd.merge(df_fa, df_branch, on=['영업가족코드'], how='left')

# 소속부문별 재적인원
df_channel = df_merge.groupby(['소속부문'])['사원번호'].count().reset_index(name='재적인원')
# 컬럼명 재설정 (소속부문 -> 항목)
df_channel.rename(columns={'소속부문':'항목'}, inplace=True)
# 구분 컬럼 생성
df_channel['구분'] = '소속부문'
# 입사연차별 재적인원
df_career = df_merge.groupby(['입사연차'])['사원번호'].count().reset_index(name='재적인원')
# 컬럼명 재설정 (입사연차 -> 항목)
df_career.rename(columns={'입사연차':'항목'}, inplace=True)
# 구분 컬럼 생성
df_career['구분'] = '입사연차'
# 소속부문+입사연차
df_registered = pd.concat([df_channel, df_career], axis=0, ignore_index=True)
# 월 값 삽입
df_registered['월'] = input_month + '월'

df_registered = df_registered[['월','구분','항목','재적인원']]
df_registered.to_excel('registered.xlsx', index=False)