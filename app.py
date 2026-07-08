"""
광고 성과 장표 자동화 - Streamlit 웹앱
CSV 파일 업로드 → Excel 장표 + HTML 대시보드 자동 생성
"""

import streamlit as st
import pandas as pd
import tempfile
import shutil
import unicodedata
import traceback
import zipfile
import io
from pathlib import Path
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import sys

# ── 페이지 설정 ─────────────────────────────────────────
st.set_page_config(
    page_title="광고 성과 장표 자동화",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── 헬퍼 함수 ───────────────────────────────────────────
def find_csv(folder: Path, keyword: str) -> list:
    """폴더에서 keyword를 포함하는 CSV 목록 반환 (NFD/NFC 한글 대응)"""
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
    """파일 목록 파싱 후 concat, 없으면 None"""
    if not files:
        return None
    dfs = []
    for f in files:
        try:
            df = parse_fn(f, **kwargs)
            dfs.append(df)
            st.write(f"  ✅ {Path(f).name}")
        except Exception as e:
            st.warning(f"  ⚠️ {Path(f).name}: {e}")
    return pd.concat(dfs, ignore_index=True) if dfs else None


# ── UI 헤더 ─────────────────────────────────────────────
st.title("📊 광고 성과 장표 자동 생성")
st.caption("CSV 파일을 업로드하면 Excel 장표와 HTML 대시보드를 자동으로 생성합니다.")

st.info(
    "📌 **전월 비교 수치를 보려면 전월 데이터도 함께 업로드해야 합니다.**\n\n"
    "당월 CSV만 올리면 전월 대비 증감(▲▼) 수치가 표시되지 않습니다.  \n"
    "**업로드 방법:** 각 매체(네이버·구글 SA)에서 **당월 + 전월** 기간의 CSV를 각각 다운로드하여 함께 업로드하세요.",
    icon="ℹ️"
)
st.divider()

# ── 파일 업로드 섹션 ─────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 🟢 네이버 SA")
    naver_files = st.file_uploader(
        "주간리포트 + 키워드별 성과",
        type="csv",
        accept_multiple_files=True,
        key="naver",
        help="파일명에 '주간리포트' 또는 '키워드별' 이 포함된 CSV"
    )
    if naver_files:
        for f in naver_files:
            st.caption(f"📎 {f.name}")

with col2:
    st.markdown("### 🔵 구글 SA")
    gsa_files = st.file_uploader(
        "게재된 시점 + 검색 키워드 + 전환 + 검색어",
        type="csv",
        accept_multiple_files=True,
        key="gsa",
        help="파일명에 '게재된 시점', '검색 키워드', '전환', '검색어' 가 포함된 CSV"
    )
    if gsa_files:
        for f in gsa_files:
            st.caption(f"📎 {f.name}")

# DA 광고는 중단 — 업로드 섹션 제거됨
gda_files = []

st.divider()

# ── 전월 구글 전환 CSV 업로드 (버튼별 전월 대비 비교용) ───────
with st.expander("📅 전월 구글 SA 전환 CSV 업로드 (버튼별 전월 대비 비교)", expanded=True):
    st.caption(
        "업로드하면 대시보드의 버튼 전환 테이블에 **이번달 / 전월 / 증감(▲▼)** 비교 컬럼이 추가됩니다.  \n"
        "업로드하지 않으면 이번달 수치만 표시됩니다.  \n"
        "**다운로드 방법:** 구글 애즈 → 전환 메뉴 → 기간을 '전월'로 설정 → SA CSV 다운로드"
    )
    st.markdown("**🔵 전월 구글 SA 전환**")
    prev_gsa_conv_files = st.file_uploader(
        "전월 SA_전환_*.csv",
        type="csv",
        accept_multiple_files=True,
        key="prev_gsa_conv",
        help="전월 기간의 구글 SA 전환 CSV (파일명에 '전환' 포함)"
    )
    if prev_gsa_conv_files:
        for f in prev_gsa_conv_files:
            st.caption(f"📎 {f.name}")

# DA 전환 업로드 제거됨
prev_gda_conv_files = []

st.divider()

# ── Meta 광고 + Instagram 유기 데이터 연동 (선택) ─────────
with st.expander("🔷 Meta 광고 + 📸 Instagram 유기 데이터 연동 (선택사항)", expanded=False):
    st.caption(
        "Meta 광고·Instagram 유기 CSV를 업로드하면 보고서에 자동 포함됩니다.  \n"
        "Meta: `meta_scraper.py` 실행 → `meta_광고성과_*.csv`  \n"
        "Instagram: `instagram_scraper.py` 실행 → `instagram_게시물성과_*.csv` 하나만 업로드하면 됩니다."
    )
    mc1, mc2 = st.columns(2)
    with mc1:
        st.markdown("**🔷 Meta 광고 성과**")
        meta_file = st.file_uploader(
            "meta_광고성과_*.csv",
            type="csv",
            key="meta",
            help="meta_scraper.py 로 생성된 Meta 광고 성과 CSV"
        )
        if meta_file:
            st.caption(f"📎 {meta_file.name}")
    with mc2:
        st.markdown("**📸 Instagram 게시물 성과**")
        ig_media_file = st.file_uploader(
            "instagram_게시물성과_*.csv",
            type="csv",
            key="ig_media",
            help="instagram_scraper.py 로 생성된 게시물 성과 CSV — 하나만 올리면 일별 차트도 자동 생성!"
        )
        if ig_media_file:
            st.caption(f"📎 {ig_media_file.name}")

st.divider()

# ── SNS 관리대장 연동 (선택) ──────────────────────────────
with st.expander("📱 SNS 채널 관리대장 연동 (선택사항)", expanded=False):
    st.caption("SNS_채널_관리대장.xlsx 를 업로드하면 종합 장표에 채널 현황이 함께 표시됩니다.")
    sns_file = st.file_uploader(
        "SNS 관리대장 파일",
        type=["xlsx"],
        key="sns",
        help="이전에 생성한 SNS_채널_관리대장.xlsx 파일을 업로드하세요."
    )
    if sns_file:
        st.caption(f"📎 {sns_file.name}")

st.divider()

# ── 보고서 기간 설정 ──────────────────────────────────────
st.markdown("### 📅 보고서 기간 설정")
st.caption("이번 기간과 비교 기간을 각각 직접 선택하세요. 기본값은 이번 달 전체 / 지난 달 전체입니다.")

this_month      = date.today().replace(day=1)
last_day_curr   = this_month + relativedelta(months=1) - relativedelta(days=1)
prev_month_first = this_month - relativedelta(months=1)
prev_month_last  = this_month - relativedelta(days=1)   # 이번달 1일 - 1일 = 전월 말일

period_col1, period_col2 = st.columns(2)

with period_col1:
    st.markdown("**📊 이번 기간**")
    report_start = st.date_input(
        "시작일", value=this_month, key="report_start",
        help="보고서 대상 기간의 첫째 날"
    )
    report_end = st.date_input(
        "종료일", value=last_day_curr, key="report_end",
        help="보고서 대상 기간의 마지막 날"
    )

with period_col2:
    st.markdown("**🔄 비교 기간 (전월)**")
    prev_start = st.date_input(
        "시작일", value=prev_month_first, key="prev_start",
        help="비교할 기간의 첫째 날 (보통 전월 1일)"
    )
    prev_end = st.date_input(
        "종료일", value=prev_month_last, key="prev_end",
        help="비교할 기간의 마지막 날 (보통 전월 말일)"
    )

st.info(
    f"📊 **이번 기간:** {report_start.strftime('%Y.%m.%d')} ~ {report_end.strftime('%Y.%m.%d')}"
    f"　　"
    f"🔄 **비교 기간:** {prev_start.strftime('%Y.%m.%d')} ~ {prev_end.strftime('%Y.%m.%d')}"
)

st.divider()

# ── 생성 버튼 ────────────────────────────────────────────
generate = st.button("🚀 보고서 생성", type="primary", use_container_width=True)

if generate:
    if not naver_files and not gsa_files:
        st.error("❌ 파일을 하나 이상 업로드해주세요.")
    else:
        tmp = None
        try:
            with st.spinner("보고서 생성 중... 잠시만 기다려주세요."):

                # 임시 폴더 구조 생성
                tmp = tempfile.mkdtemp()
                naver_dir = Path(tmp) / "naver"
                gsa_dir   = Path(tmp) / "google_sa"
                for d in [naver_dir, gsa_dir]:
                    d.mkdir()

                # 업로드된 파일 저장
                for f in (naver_files or []):
                    (naver_dir / f.name).write_bytes(f.read())
                for f in (gsa_files or []):
                    (gsa_dir / f.name).write_bytes(f.read())

                # SNS 관리대장 저장 (있을 경우)
                sns_path = None
                if sns_file:
                    sns_save = Path(tmp) / sns_file.name
                    sns_save.write_bytes(sns_file.read())
                    sns_path = str(sns_save)

                # 파서 임포트
                sys.path.insert(0, str(Path(__file__).parent))
                from parsers import (
                    parse_naver_weekly, parse_naver_keywords,
                    parse_google_daily, parse_google_keywords,
                    parse_google_conversions, parse_google_search_terms
                )
                from report_builder import build_report
                from html_builder import build_html_report

                # 파일 탐색
                naver_weekly_files = find_csv(naver_dir, '주간리포트')
                naver_kw_files     = find_csv(naver_dir, '키워드별')
                gsa_daily_files    = find_csv(gsa_dir,   '게재된 시점')
                gsa_kw_files       = find_csv(gsa_dir,   '검색 키워드')
                gsa_conv_files     = find_csv(gsa_dir,   '전환')
                gsa_srch_files     = find_csv(gsa_dir,   '검색어')

                # Raw 데이터 파싱
                raw_dfs = []

                naver_raw = load_or_empty(parse_naver_weekly, naver_weekly_files)
                if naver_raw is not None:
                    raw_dfs.append(naver_raw)

                gsa_raw = load_or_empty(parse_google_daily, gsa_daily_files, media_type='Google_SA')
                if gsa_raw is not None:
                    raw_dfs.append(gsa_raw)

                if not raw_dfs:
                    st.error("❌ 유효한 데이터가 없습니다. 파일명을 확인해주세요.")
                    st.info("파일명에 '주간리포트', '게재된 시점' 키워드가 포함되어야 합니다.")
                    st.stop()

                raw_df_all = pd.concat(raw_dfs, ignore_index=True)
                raw_df_all = raw_df_all.sort_values(['날짜', '매체', '디바이스']).reset_index(drop=True)
                raw_df_all['날짜'] = pd.to_datetime(raw_df_all['날짜'])

                # ── 기간 필터링 ──────────────────────────
                # 보고서 대상 기간
                mask_curr = (
                    (raw_df_all['날짜'] >= pd.Timestamp(report_start)) &
                    (raw_df_all['날짜'] <= pd.Timestamp(report_end))
                )
                raw_df = raw_df_all[mask_curr].copy()

                if raw_df.empty:
                    st.error(
                        f"❌ 선택한 기간({report_start} ~ {report_end})에 해당하는 데이터가 없습니다.\n"
                        "업로드한 파일의 날짜 범위를 확인해주세요."
                    )
                    st.stop()

                # 전월 동기간 자동 추출
                mask_prev = (
                    (raw_df_all['날짜'] >= pd.Timestamp(prev_start)) &
                    (raw_df_all['날짜'] <= pd.Timestamp(prev_end))
                )
                prev_raw_df = raw_df_all[mask_prev].copy()
                if prev_raw_df.empty:
                    prev_raw_df = None
                    st.caption(f"ℹ️ 전월 동기간({prev_start} ~ {prev_end}) 데이터 없음 → 전월 비교 생략")
                else:
                    st.caption(f"✅ 전월 동기간 자동 감지: {prev_start} ~ {prev_end} ({len(prev_raw_df)}행)")

                # ── Meta 광고 CSV 처리 ────────────────────
                if meta_file:
                    try:
                        meta_file.seek(0)
                        _m = pd.read_csv(meta_file, encoding='utf-8-sig')
                        _m['날짜'] = pd.to_datetime(_m['날짜'])
                        _m['매체'] = 'Meta'
                        if '디바이스' not in _m.columns:
                            _m['디바이스'] = '전체'
                        common_cols = [c for c in ['날짜', '매체', '디바이스', '노출', '클릭', '비용', '전환', 'CTR', 'CPC'] if c in _m.columns]
                        _m_filtered = _m[common_cols]
                        # 이번 기간
                        mask_meta = (
                            (_m_filtered['날짜'] >= pd.Timestamp(report_start)) &
                            (_m_filtered['날짜'] <= pd.Timestamp(report_end))
                        )
                        meta_filtered = _m_filtered[mask_meta]
                        if not meta_filtered.empty:
                            raw_df = pd.concat([raw_df, meta_filtered], ignore_index=True)
                            raw_df = raw_df.sort_values(['날짜', '매체']).reset_index(drop=True)
                            st.caption(f"✅ Meta 광고 데이터 {len(meta_filtered)}일치 포함")
                        else:
                            st.caption("ℹ️ Meta CSV: 선택 기간 내 데이터 없음")
                        # 전월 비교 기간 — prev_raw_df에 Meta 전월 데이터 추가
                        mask_meta_prev = (
                            (_m_filtered['날짜'] >= pd.Timestamp(prev_start)) &
                            (_m_filtered['날짜'] <= pd.Timestamp(prev_end))
                        )
                        meta_prev = _m_filtered[mask_meta_prev]
                        if not meta_prev.empty:
                            if prev_raw_df is None:
                                prev_raw_df = meta_prev.copy()
                            else:
                                prev_raw_df = pd.concat([prev_raw_df, meta_prev], ignore_index=True)
                                prev_raw_df = prev_raw_df.sort_values(['날짜', '매체']).reset_index(drop=True)
                            st.caption(f"✅ Meta 전월 데이터 {len(meta_prev)}일치 포함 ({prev_start} ~ {prev_end})")
                        else:
                            st.caption(f"ℹ️ Meta CSV: 비교 기간 내 데이터 없음 ({prev_start} ~ {prev_end})")
                    except Exception as e:
                        st.warning(f"⚠️ Meta CSV 로드 실패: {e}")

                # ── Instagram 유기 CSV 처리 ───────────────
                ig_account_df    = None   # 게시물 데이터에서 자동 계산됨
                ig_media_df      = None
                prev_ig_media_df = None

                if ig_media_file:
                    try:
                        ig_media_file.seek(0)
                        _ig_all = pd.read_csv(ig_media_file, encoding='utf-8-sig')
                        _ig_all['날짜'] = pd.to_datetime(_ig_all['날짜'])
                        # 이번 기간 필터
                        mask_ig_curr = (
                            (_ig_all['날짜'] >= pd.Timestamp(report_start)) &
                            (_ig_all['날짜'] <= pd.Timestamp(report_end))
                        )
                        ig_media_df = _ig_all[mask_ig_curr].copy()
                        # 전월 비교 기간 필터
                        mask_ig_prev = (
                            (_ig_all['날짜'] >= pd.Timestamp(prev_start)) &
                            (_ig_all['날짜'] <= pd.Timestamp(prev_end))
                        )
                        prev_ig_media_df = _ig_all[mask_ig_prev].copy()
                        if prev_ig_media_df.empty:
                            prev_ig_media_df = None
                        st.caption(
                            f"✅ Instagram 게시물: 이번달 {len(ig_media_df)}개 "
                            f"/ 전월 {len(prev_ig_media_df) if prev_ig_media_df is not None else 0}개"
                        )
                    except Exception as e:
                        st.warning(f"⚠️ Instagram 게시물 성과 로드 실패: {e}")

                # 부가 데이터 파싱
                naver_kw   = load_or_empty(parse_naver_keywords, naver_kw_files)
                gsa_kw     = load_or_empty(parse_google_keywords, gsa_kw_files)
                srch_df    = load_or_empty(parse_google_search_terms, gsa_srch_files)

                # ── 구글 SA 전환 CSV 파싱 ──────────────────
                st.caption(f"🔍 SA 전환 파일 탐색: {[Path(f).name for f in gsa_conv_files]}")
                sa_conv_df = load_or_empty(parse_google_conversions, gsa_conv_files)
                if sa_conv_df is not None:
                    sa_conv_df = sa_conv_df.groupby('전환유형', as_index=False)['전환수'].sum()
                    st.caption(f"✅ SA 전환 데이터 파싱 완료: {sa_conv_df.to_dict('records')}")
                else:
                    st.caption("⚠️ SA 전환 파일 없음 또는 파싱 실패")

                da_conv_df = None  # DA 광고 중단

                # 전월 전환 CSV (선택) — 있으면 페이지조회 제외한 정확한 버튼 전환율 계산
                prev_sa_conv_df = None
                if prev_gsa_conv_files:
                    prev_gsa_dir = Path(tmp) / "prev_gsa_conv"
                    prev_gsa_dir.mkdir(exist_ok=True)
                    for f in prev_gsa_conv_files:
                        (prev_gsa_dir / f.name).write_bytes(f.read())
                    prev_gsa_paths = find_csv(prev_gsa_dir, '')   # 전체 csv 탐색
                    prev_sa_conv_df = load_or_empty(parse_google_conversions, prev_gsa_paths)
                    if prev_sa_conv_df is not None:
                        prev_sa_conv_df = prev_sa_conv_df.groupby('전환유형', as_index=False)['전환수'].sum()
                        st.caption(f"✅ 전월 SA 전환 데이터 포함 (버튼전환율 정확도 향상)")

                prev_da_conv_df = None  # DA 광고 중단

                # 기간 라벨
                period_label = f'{report_start.strftime("%Y.%m.%d")} ~ {report_end.strftime("%Y.%m.%d")}'

                # 출력 파일 경로
                today = datetime.now().strftime('%Y%m%d')
                xlsx_path = str(Path(tmp) / f'광고성과_장표_{today}.xlsx')
                html_path = str(Path(tmp) / f'광고성과_대시보드_{today}.html')

                # Excel 장표 생성
                build_report(
                    raw_df=raw_df,
                    naver_kw_df=naver_kw,
                    google_sa_kw_df=gsa_kw,
                    sa_conv_df=sa_conv_df,
                    search_terms_df=srch_df,
                    period_label=period_label,
                    output_path=xlsx_path,
                    sns_tracker_path=sns_path,
                )

                # HTML 대시보드 생성
                build_html_report(
                    raw_df=raw_df,
                    sa_conv_df=sa_conv_df,
                    period_label=period_label,
                    output_path=html_path,
                    sns_tracker_path=sns_path,
                    prev_raw_df=prev_raw_df,
                    ig_account_df=ig_account_df,
                    ig_media_df=ig_media_df,
                    prev_ig_media_df=prev_ig_media_df,
                    prev_sa_conv_df=prev_sa_conv_df,
                )

            # ── 완료 & 다운로드 ──────────────────────────
            st.success(f"✅ 보고서 생성 완료!  |  기간: {period_label}")
            st.balloons()

            # ZIP 파일 생성 (Excel + HTML 동시 다운로드용)
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                zf.write(xlsx_path, arcname=f'광고성과_장표_{today}.xlsx')
                zf.write(html_path, arcname=f'광고성과_대시보드_{today}.html')
            zip_bytes = zip_buf.getvalue()

            col_dl1, col_dl2, col_dl3 = st.columns(3)
            with col_dl1:
                with open(xlsx_path, 'rb') as f:
                    st.download_button(
                        label="📊 Excel 장표",
                        data=f.read(),
                        file_name=f'광고성과_장표_{today}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True,
                    )
            with col_dl2:
                with open(html_path, 'rb') as f:
                    st.download_button(
                        label="🌐 HTML 대시보드",
                        data=f.read(),
                        file_name=f'광고성과_대시보드_{today}.html',
                        mime='text/html',
                        use_container_width=True,
                    )
            with col_dl3:
                st.download_button(
                    label="📦 Excel + HTML 한번에",
                    data=zip_bytes,
                    file_name=f'광고성과_전체_{today}.zip',
                    mime='application/zip',
                    use_container_width=True,
                    type="primary",
                )

        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
            with st.expander("오류 상세 보기"):
                st.code(traceback.format_exc())
        finally:
            if tmp and Path(tmp).exists():
                shutil.rmtree(tmp, ignore_errors=True)

# ── 하단 안내 ────────────────────────────────────────────
st.divider()
with st.expander("📌 파일명 규칙 안내"):
    st.markdown("""
    파일명에 아래 키워드가 **반드시 포함**되어야 자동 인식됩니다.

    | 매체 | 파일 종류 | 파일명에 포함될 키워드 | 생성 방법 |
    |------|----------|----------------------|----------|
    | 네이버 SA | 일별 성과 | `주간리포트` | 네이버 검색광고 다운로드 |
    | 네이버 SA | 키워드 성과 | `키워드별` | 네이버 검색광고 다운로드 |
    | 구글 SA | 일별 성과 | `게재된 시점` | Google Ads 다운로드 |
    | 구글 SA | 키워드 | `검색 키워드` | Google Ads 다운로드 |
    | 구글 SA | 전환 | `전환` | Google Ads 다운로드 |
    | 구글 SA | 검색어 | `검색어` | Google Ads 다운로드 |
    | **Meta 광고** | 광고 성과 | `meta_광고성과_` | `python meta_scraper.py` |
    | **Instagram** | 계정 인사이트 | `instagram_계정인사이트_` | `python instagram_scraper.py` |
    | **Instagram** | 게시물 성과 | `instagram_게시물성과_` | `python instagram_scraper.py` |
    """)
