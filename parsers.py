"""
광고 CSV 파서 모듈
- 네이버 SA (UTF-8-SIG)
- 구글 SA / DA (UTF-16, tab-separated)
"""

import pandas as pd
import numpy as np
from pathlib import Path
import re


# ────────────────────────────────────────
# 공통 유틸
# ────────────────────────────────────────

def _clean_number(val):
    """'1,234' → 1234.0 / '--' → 0"""
    if pd.isna(val):
        return 0.0
    s = str(val).replace(',', '').replace('%', '').strip()
    if s in ('--', '-', ''):
        return 0.0
    try:
        return float(s)
    except:
        return 0.0


def _get_week_label(date, year_month_first):
    """날짜 → 'N월M주' 형식 (Raw 시트 주차 컬럼용)"""
    month = date.month
    # 해당 월 1일 기준 주차 계산
    first_day = pd.Timestamp(date.year, date.month, 1)
    week_num = (date.day - 1) // 7 + 1
    return f"{month}월{week_num}주"


# ────────────────────────────────────────
# 네이버 SA 파서
# ────────────────────────────────────────

def parse_naver_weekly(filepath: str) -> pd.DataFrame:
    """
    '주간리포트 양식' CSV → Raw 시트 포맷 DataFrame
    컬럼: 월, 주차, 날짜, 매체, 디바이스, 노출, 클릭, 전환, 비용
    매체 매핑: 파워링크→Naver_SA_PL, 쇼핑검색→Naver_SA_SP
    """
    df = pd.read_csv(filepath, encoding='utf-8-sig', skiprows=1)
    df.columns = df.columns.str.strip()

    # ✅ 플랫폼사업 캠페인만 필터링 (리본모바일, 소상공인 등 제외)
    df = df[df['캠페인'].str.contains('플랫폼사업', na=False)].copy()

    # 날짜 파싱
    df['날짜'] = pd.to_datetime(df['일별'], format='%Y.%m.%d.')
    df['월'] = df['날짜'].apply(lambda d: pd.Timestamp(d.year, d.month, 1))
    df['주차'] = df['날짜'].apply(_get_week_label, year_month_first=None)

    # 매체명 매핑
    매체_map = {
        '파워링크': 'Naver_SA_PL',
        '쇼핑검색': 'Naver_SA_SP',
        '브랜드검색/신제품검색': 'Naver_SA_BR',
    }
    df['매체'] = df['캠페인유형'].map(매체_map).fillna(df['캠페인유형'])

    # 디바이스 정규화
    df['디바이스'] = df['PC/모바일 매체'].str.strip()
    df.loc[df['디바이스'] == '모바일', '디바이스'] = '모바일'

    # 숫자 컬럼
    df['노출'] = df['노출수'].apply(_clean_number)
    df['클릭'] = df['클릭수'].apply(_clean_number)
    df['전환'] = df['총 전환수'].apply(_clean_number)
    df['비용'] = df['총비용'].apply(_clean_number)

    # 날짜 × 매체 × 디바이스 집계
    result = df.groupby(['월', '주차', '날짜', '매체', '디바이스'], as_index=False).agg(
        노출=('노출', 'sum'),
        클릭=('클릭', 'sum'),
        전환=('전환', 'sum'),
        비용=('비용', 'sum'),
    )

    return result[['월', '주차', '날짜', '매체', '디바이스', '노출', '클릭', '전환', '비용']]


def parse_naver_keywords(filepath: str) -> pd.DataFrame:
    """
    '키워드별 성과' CSV → PL_키워드 상세 시트 포맷 DataFrame
    """
    df = pd.read_csv(filepath, encoding='utf-8-sig', skiprows=1)
    df.columns = df.columns.str.strip()

    # 기간 헤더에서 날짜 추출
    with open(filepath, encoding='utf-8-sig') as f:
        header = f.readline().strip()
    # "키워드별 성과(2026.05.01.~2026.05.17.),349973"
    dates = re.findall(r'(\d{4}\.\d{2}\.\d{2}\.)', header)
    period = f"{dates[0]}~{dates[1]}" if len(dates) >= 2 else ''

    # ✅ 플랫폼사업 캠페인만 필터링
    df = df[df['캠페인'].str.contains('플랫폼사업', na=False)].copy()

    df['광고비'] = df['총비용'].apply(_clean_number)
    df['노출'] = df['노출수'].apply(_clean_number)
    df['클릭'] = df['클릭수'].apply(_clean_number)
    df['평균노출순위'] = df['평균노출순위'].apply(_clean_number)
    df['기간'] = period

    result = df[['캠페인', '광고그룹', '키워드', '노출', '클릭', '광고비', '평균노출순위', '기간']].copy()
    result.columns = ['캠페인', '광고그룹', '키워드', '노출', '클릭', '광고비(VAT별도)', '평균노출순위', '기간']
    return result


# ────────────────────────────────────────
# 구글 파서 (공통: UTF-16, tab-sep, skiprows=2)
# ────────────────────────────────────────

def _read_google_csv(filepath: str) -> pd.DataFrame:
    df = pd.read_csv(filepath, encoding='utf-16', sep='\t', skiprows=2)
    df.columns = df.columns.str.strip()
    # 마지막 합계 행 제거 (보통 '합계' 또는 비어있음)
    df = df[~df.iloc[:, 0].astype(str).str.startswith('합계')]
    df = df[~df.iloc[:, 0].astype(str).str.startswith('총계')]
    return df


def parse_google_daily(filepath: str, media_type: str = 'Google_SA') -> pd.DataFrame:
    """
    '광고가 게재된 시점' CSV → Raw 시트 포맷
    media_type: 'Google_SA' 또는 'Google_DA'
    """
    df = _read_google_csv(filepath)

    # 헤더에서 기간 파악
    with open(filepath, encoding='utf-16') as f:
        header = f.readline().strip()

    df['날짜'] = pd.to_datetime(df['일'], errors='coerce')
    df = df.dropna(subset=['날짜'])
    df['월'] = df['날짜'].apply(lambda d: pd.Timestamp(d.year, d.month, 1))
    df['주차'] = df['날짜'].apply(_get_week_label, year_month_first=None)
    df['매체'] = media_type
    df['디바이스'] = '*'  # 구글 일별 파일은 디바이스 미구분

    df['노출'] = df['노출수'].apply(_clean_number)
    df['클릭'] = df['클릭수'].apply(_clean_number)
    df['전환'] = df['전환'].apply(_clean_number)
    df['비용'] = df['비용'].apply(_clean_number)

    result = df.groupby(['월', '주차', '날짜', '매체', '디바이스'], as_index=False).agg(
        노출=('노출', 'sum'),
        클릭=('클릭', 'sum'),
        전환=('전환', 'sum'),
        비용=('비용', 'sum'),
    )

    return result[['월', '주차', '날짜', '매체', '디바이스', '노출', '클릭', '전환', '비용']]


def parse_google_keywords(filepath: str) -> pd.DataFrame:
    """'검색 키워드' CSV → GSA_키워드 상세 시트 포맷"""
    df = _read_google_csv(filepath)

    df['노출'] = df['노출수'].apply(_clean_number)
    df['클릭'] = df['클릭수'].apply(_clean_number)
    df['비용'] = df['비용'].apply(_clean_number)
    df['전환'] = df['전환'].apply(_clean_number)
    df['평균CPC'] = df['평균 CPC'].apply(_clean_number)

    cols_map = {
        '검색 키워드': '키워드',
        'Google 검색 키워드 검색 유형': '검색유형',
        '캠페인': '캠페인',
        '광고그룹': '광고그룹',
    }
    df = df.rename(columns=cols_map)

    keep = ['캠페인', '광고그룹', '키워드', '검색유형', '노출', '클릭', '평균CPC', '비용', '전환']
    return df[[c for c in keep if c in df.columns]]


def parse_google_conversions(filepath: str) -> pd.DataFrame:
    """
    '전환' CSV → 전환_상세 시트 포맷
    전환액션별: 페이지조회, 상세보기, 상담예약, 상담신청
    """
    df = _read_google_csv(filepath)

    df['전환수'] = df['전환'].apply(_clean_number)

    # 전환 유형 분류
    def classify(action):
        a = str(action).lower()
        if '상세보기' in a:
            return '상세보기_버튼클릭'
        elif '상담예약' in a:
            return '상담예약_버튼클릭'
        elif '상담신청' in a:
            return '상담신청_버튼클릭'
        elif '페이지' in a or 'page' in a:
            return '페이지조회'
        else:
            return '기타'

    df['전환유형'] = df['전환 액션'].apply(classify)
    df = df[['전환유형', '전환수']].copy()
    df.columns = ['전환유형', '전환수']
    return df


def parse_google_search_terms(filepath: str) -> pd.DataFrame:
    """'검색어' CSV → 검색어 분석 시트 포맷"""
    df = _read_google_csv(filepath)

    df['노출'] = df['노출수'].apply(_clean_number)
    df['클릭'] = df['클릭수'].apply(_clean_number)
    df['비용'] = df['비용'].apply(_clean_number)
    df['전환'] = df['전환'].apply(_clean_number)
    df['평균CPC'] = df['평균 CPC'].apply(_clean_number)

    keep = ['검색어', '클릭수', '노출수', '클릭률(CTR)', '평균CPC', '비용', '전환']
    df = df.rename(columns={'클릭수': '클릭', '노출수': '노출'})
    return df[['검색어', '노출', '클릭', '평균CPC', '비용', '전환']]
