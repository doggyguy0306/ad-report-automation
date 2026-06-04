#!/usr/bin/env python3
"""
광고 성과 장표 자동 업데이트 스크립트

사용법:
    python update_report.py

폴더 구조:
    raw_data/
      naver/        ← 네이버 CSV 여기 넣기
      google_sa/    ← 구글 SA CSV 여기 넣기
      google_da/    ← 구글 DA CSV 여기 넣기
    reports/        ← 생성된 장표 저장됨
"""

import sys
import os
import shutil
import glob
from pathlib import Path
from datetime import datetime
import pandas as pd

# 현재 스크립트 위치 기준으로 경로 설정
BASE_DIR = Path(__file__).parent
RAW_DIR  = BASE_DIR / 'raw_data'
NAVER_DIR   = RAW_DIR / 'naver'
GSA_DIR     = RAW_DIR / 'google_sa'
GDA_DIR     = RAW_DIR / 'google_da'
PROC_DIR    = RAW_DIR / 'processed'
REPORT_DIR  = BASE_DIR / 'reports'

# SNS 관리대장: 상위 폴더(~/Projects)에 있으면 자동 연동
SNS_TRACKER = BASE_DIR.parent / 'SNS_채널_관리대장.xlsx'

# Meta / Instagram: KT-Plaza-Marketing 폴더에서 최신 CSV 자동 감지
KT_PLAZA_DIR = BASE_DIR.parent / 'KT-Plaza-Marketing'

# 파일 이름 패턴 (구글은 이름으로 역할 구분)
NAVER_WEEKLY_KW = '주간리포트'   # 주간리포트 양식
NAVER_KEYWORD_KW = '키워드별'    # 키워드별 성과
GOOGLE_DAILY_KW  = '게재된 시점' # 광고가 게재된 시점
GOOGLE_KW_KW     = '검색 키워드' # 검색 키워드
GOOGLE_CONV_KW   = '전환'        # 전환
GOOGLE_SRCH_KW   = '검색어'      # 검색어


def find_latest_csv(folder: Path, prefix: str):
    """prefix로 시작하는 가장 최신 CSV 반환 (없으면 None)"""
    if not folder.exists():
        return None
    matches = sorted(folder.glob(f'{prefix}*.csv'), reverse=True)
    return matches[0] if matches else None


def find_csv(folder: Path, keyword: str) -> list:
    """폴더에서 keyword를 포함하는 CSV 목록 반환 (macOS NFD/NFC 한글 파일명 대응)"""
    import unicodedata
    if not folder.exists():
        return []
    kw_nfc = unicodedata.normalize('NFC', keyword)
    kw_nfd = unicodedata.normalize('NFD', keyword)
    result = []
    for f in folder.iterdir():
        if f.suffix.lower() != '.csv':
            continue
        name_nfc = unicodedata.normalize('NFC', f.name)
        name_nfd = unicodedata.normalize('NFD', f.name)
        if kw_nfc in name_nfc or kw_nfd in name_nfd:
            result.append(str(f))
    return sorted(result)


def load_or_empty(parse_fn, files, **kwargs):
    """파일 목록에서 파싱 후 concat, 없으면 None"""
    if not files:
        return None
    dfs = []
    for f in files:
        try:
            df = parse_fn(f, **kwargs)
            dfs.append(df)
            print(f'  ✅ {Path(f).name}')
        except Exception as e:
            print(f'  ❌ {Path(f).name}: {e}')
    return pd.concat(dfs, ignore_index=True) if dfs else None


def move_processed(files: list):
    """처리된 파일을 processed/ 폴더로 이동"""
    for f in files:
        dest = PROC_DIR / f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_{Path(f).name}'
        shutil.move(f, dest)


def main():
    from parsers import (
        parse_naver_weekly, parse_naver_keywords,
        parse_google_daily, parse_google_keywords,
        parse_google_conversions, parse_google_search_terms
    )
    from report_builder import build_report
    from html_builder import build_html_report

    print('=' * 55)
    print('  광고 성과 장표 자동 업데이트')
    print(f'  실행 시각: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    print('=' * 55)

    # ── 1. 파일 탐색 ──────────────────────────────────
    naver_weekly_files  = find_csv(NAVER_DIR, NAVER_WEEKLY_KW)
    naver_kw_files      = find_csv(NAVER_DIR, NAVER_KEYWORD_KW)
    gsa_daily_files     = find_csv(GSA_DIR, GOOGLE_DAILY_KW)
    gsa_kw_files        = find_csv(GSA_DIR, GOOGLE_KW_KW)
    gsa_conv_files      = find_csv(GSA_DIR, GOOGLE_CONV_KW)
    gsa_srch_files      = find_csv(GSA_DIR, GOOGLE_SRCH_KW)
    gda_daily_files     = find_csv(GDA_DIR, GOOGLE_DAILY_KW)
    gda_kw_files        = find_csv(GDA_DIR, GOOGLE_KW_KW)
    gda_conv_files      = find_csv(GDA_DIR, GOOGLE_CONV_KW)

    # 파일 없으면 중단
    all_raw_files = naver_weekly_files + gsa_daily_files + gda_daily_files
    if not all_raw_files:
        print()
        print('⚠️  처리할 파일이 없습니다.')
        print(f'   → 네이버 CSV: {NAVER_DIR}')
        print(f'   → 구글 SA CSV: {GSA_DIR}')
        print(f'   → 구글 DA CSV: {GDA_DIR}')
        sys.exit(0)

    print()
    print('[1] 파일 탐색 완료')
    print(f'  네이버: {len(naver_weekly_files)}개 주간리포트, {len(naver_kw_files)}개 키워드')
    print(f'  구글SA: {len(gsa_daily_files)}개 일별, {len(gsa_kw_files)}개 키워드')
    print(f'  구글DA: {len(gda_daily_files)}개 일별, {len(gda_kw_files)}개 키워드')

    # ── 2. 파싱 ──────────────────────────────────────
    print()
    print('[2] CSV 파싱 중...')

    raw_dfs = []

    print('  [네이버 주간리포트]')
    naver_raw = load_or_empty(parse_naver_weekly, naver_weekly_files)
    if naver_raw is not None:
        raw_dfs.append(naver_raw)

    print('  [구글 SA 일별]')
    gsa_raw = load_or_empty(parse_google_daily, gsa_daily_files, media_type='Google_SA')
    if gsa_raw is not None:
        raw_dfs.append(gsa_raw)

    print('  [구글 DA 일별]')
    gda_raw = load_or_empty(parse_google_daily, gda_daily_files, media_type='Google_DA')
    if gda_raw is not None:
        raw_dfs.append(gda_raw)

    if not raw_dfs:
        print('❌ 유효한 Raw 데이터가 없습니다. 종료.')
        sys.exit(1)

    raw_df = pd.concat(raw_dfs, ignore_index=True)
    raw_df = raw_df.sort_values(['날짜', '매체', '디바이스']).reset_index(drop=True)

    print()
    print('  [네이버 키워드]')
    naver_kw = load_or_empty(parse_naver_keywords, naver_kw_files)

    print('  [구글 SA 키워드]')
    gsa_kw = load_or_empty(parse_google_keywords, gsa_kw_files)

    print('  [구글 DA 키워드]')
    gda_kw = load_or_empty(parse_google_keywords, gda_kw_files)

    print('  [구글 SA 전환 상세]')
    sa_conv_df = load_or_empty(parse_google_conversions, gsa_conv_files)
    if sa_conv_df is not None:
        sa_conv_df = sa_conv_df.groupby('전환유형', as_index=False)['전환수'].sum()

    print('  [구글 DA 전환 상세]')
    da_conv_df = load_or_empty(parse_google_conversions, gda_conv_files)
    if da_conv_df is not None:
        da_conv_df = da_conv_df.groupby('전환유형', as_index=False)['전환수'].sum()

    print('  [구글 검색어]')
    srch_df = load_or_empty(parse_google_search_terms, gsa_srch_files)

    # ── 3. 기간 라벨 생성 ────────────────────────────
    dates = pd.to_datetime(raw_df['날짜'])
    period_label = f'{dates.min().strftime("%Y.%m.%d")} ~ {dates.max().strftime("%Y.%m.%d")}'

    # ── 4. SNS 관리대장 자동 감지 ────────────────────
    sns_tracker_path = None
    if SNS_TRACKER.exists():
        sns_tracker_path = str(SNS_TRACKER)
        print()
        print(f'[SNS] 관리대장 감지됨 → {SNS_TRACKER.name}')
    else:
        print()
        print(f'[SNS] 관리대장 없음 (선택사항)  — {SNS_TRACKER} 에 파일을 놓으면 종합 장표에 자동 연동됩니다.')

    # ── 4-1. Meta 광고 CSV 자동 감지 ─────────────────
    print()
    meta_df = None
    meta_csv = find_latest_csv(KT_PLAZA_DIR, 'meta_광고성과_')
    if meta_csv:
        try:
            _m = pd.read_csv(meta_csv, encoding='utf-8-sig')
            _m['날짜'] = pd.to_datetime(_m['날짜'])
            _m['매체'] = 'Meta'
            if '디바이스' not in _m.columns:
                _m['디바이스'] = '전체'
            # raw_df와 같은 컬럼만 유지
            common_cols = [c for c in ['날짜', '매체', '디바이스', '노출', '클릭', '비용', '전환', 'CTR', 'CPC'] if c in _m.columns]
            meta_df = _m[common_cols]
            print(f'[Meta] 광고 데이터 감지됨 → {meta_csv.name} ({len(meta_df)}일치)')
            # raw_df에 합산
            raw_df = pd.concat([raw_df, meta_df], ignore_index=True)
            raw_df = raw_df.sort_values(['날짜', '매체']).reset_index(drop=True)
        except Exception as e:
            print(f'[Meta] 로드 실패: {e}')
    else:
        print(f'[Meta] 광고 CSV 없음 — {KT_PLAZA_DIR}/meta_광고성과_*.csv 를 생성하면 자동 포함됩니다.')

    # ── 4-2. Instagram 유기 CSV 자동 감지 ────────────
    ig_account_df = None
    ig_media_df   = None
    ig_acct_csv   = find_latest_csv(KT_PLAZA_DIR, 'instagram_계정인사이트_')
    ig_media_csv  = find_latest_csv(KT_PLAZA_DIR, 'instagram_게시물성과_')

    if ig_acct_csv:
        try:
            ig_account_df = pd.read_csv(ig_acct_csv, encoding='utf-8-sig')
            ig_account_df['날짜'] = pd.to_datetime(ig_account_df['날짜'])
            print(f'[Instagram] 계정 인사이트 → {ig_acct_csv.name} ({len(ig_account_df)}일치)')
        except Exception as e:
            print(f'[Instagram] 계정 인사이트 로드 실패: {e}')

    if ig_media_csv:
        try:
            ig_media_df = pd.read_csv(ig_media_csv, encoding='utf-8-sig')
            ig_media_df['날짜'] = pd.to_datetime(ig_media_df['날짜'])
            print(f'[Instagram] 게시물 성과   → {ig_media_csv.name} ({len(ig_media_df)}개)')
        except Exception as e:
            print(f'[Instagram] 게시물 성과 로드 실패: {e}')

    # ── 5. 장표 생성 ─────────────────────────────────
    print()
    print('[3] 장표 + HTML 대시보드 생성 중...')
    REPORT_DIR.mkdir(exist_ok=True)
    today = datetime.now().strftime('%Y%m%d')
    xlsx_path = str(REPORT_DIR / f'광고성과_장표_{today}.xlsx')
    html_path = str(REPORT_DIR / f'광고성과_대시보드_{today}.html')

    build_report(
        raw_df=raw_df,
        naver_kw_df=naver_kw,
        google_sa_kw_df=gsa_kw,
        google_da_kw_df=gda_kw,
        sa_conv_df=sa_conv_df,
        da_conv_df=da_conv_df,
        search_terms_df=srch_df,
        period_label=period_label,
        output_path=xlsx_path,
        sns_tracker_path=sns_tracker_path,
    )

    build_html_report(
        raw_df=raw_df,
        sa_conv_df=sa_conv_df,
        da_conv_df=da_conv_df,
        period_label=period_label,
        output_path=html_path,
        sns_tracker_path=sns_tracker_path,
        ig_account_df=ig_account_df,
        ig_media_df=ig_media_df,
    )

    # ── 6. 처리된 파일 이동 ───────────────────────────
    all_files = (naver_weekly_files + naver_kw_files +
                 gsa_daily_files + gsa_kw_files + gsa_conv_files + gsa_srch_files +
                 gda_daily_files + gda_kw_files + gda_conv_files)
    move_processed(all_files)
    print(f'  📦 처리된 CSV {len(all_files)}개 → processed/ 이동 완료')

    print()
    print('=' * 55)
    print(f'  ✅ 완료! 생성 파일:')
    print(f'     📊 Excel: {xlsx_path}')
    print(f'     🌐 HTML:  {html_path}')
    print('=' * 55)


if __name__ == '__main__':
    main()
