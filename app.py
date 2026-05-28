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
from datetime import datetime
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
st.divider()

# ── 파일 업로드 섹션 ─────────────────────────────────────
col1, col2, col3 = st.columns(3)

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

with col3:
    st.markdown("### 🔴 구글 DA")
    gda_files = st.file_uploader(
        "게재된 시점 + 검색 키워드 + 전환",
        type="csv",
        accept_multiple_files=True,
        key="gda",
        help="파일명에 '게재된 시점', '검색 키워드', '전환' 이 포함된 CSV"
    )
    if gda_files:
        for f in gda_files:
            st.caption(f"📎 {f.name}")

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

# ── 전월 데이터 업로드 (전월 동기간 대비용, 선택) ─────────────
with st.expander("📅 전월 데이터 업로드 (전월 동기간 대비 비교용, 선택사항)", expanded=False):
    st.caption("전월 광고 CSV를 업로드하면 KPI 카드에 '전월 동기간 대비 +X%' 가 표시됩니다.")
    pcol1, pcol2, pcol3 = st.columns(3)
    with pcol1:
        st.markdown("**🟢 전월 네이버 SA**")
        prev_naver_files = st.file_uploader("전월 주간리포트 + 키워드별", type="csv",
            accept_multiple_files=True, key="prev_naver")
    with pcol2:
        st.markdown("**🔵 전월 구글 SA**")
        prev_gsa_files = st.file_uploader("전월 게재된 시점 + 검색 키워드 + 전환", type="csv",
            accept_multiple_files=True, key="prev_gsa")
    with pcol3:
        st.markdown("**🔴 전월 구글 DA**")
        prev_gda_files = st.file_uploader("전월 게재된 시점 + 전환", type="csv",
            accept_multiple_files=True, key="prev_gda")

st.divider()

# ── 생성 버튼 ────────────────────────────────────────────
generate = st.button("🚀 보고서 생성", type="primary", use_container_width=True)

if generate:
    if not naver_files and not gsa_files and not gda_files:
        st.error("❌ 파일을 하나 이상 업로드해주세요.")
    else:
        tmp = None
        try:
            with st.spinner("보고서 생성 중... 잠시만 기다려주세요."):

                # 임시 폴더 구조 생성
                tmp = tempfile.mkdtemp()
                naver_dir = Path(tmp) / "naver"
                gsa_dir   = Path(tmp) / "google_sa"
                gda_dir   = Path(tmp) / "google_da"
                for d in [naver_dir, gsa_dir, gda_dir]:
                    d.mkdir()

                # 업로드된 파일 저장
                for f in (naver_files or []):
                    (naver_dir / f.name).write_bytes(f.read())
                for f in (gsa_files or []):
                    (gsa_dir / f.name).write_bytes(f.read())
                for f in (gda_files or []):
                    (gda_dir / f.name).write_bytes(f.read())

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
                gda_daily_files    = find_csv(gda_dir,   '게재된 시점')
                gda_kw_files       = find_csv(gda_dir,   '검색 키워드')
                gda_conv_files     = find_csv(gda_dir,   '전환')

                # Raw 데이터 파싱
                raw_dfs = []

                naver_raw = load_or_empty(parse_naver_weekly, naver_weekly_files)
                if naver_raw is not None:
                    raw_dfs.append(naver_raw)

                gsa_raw = load_or_empty(parse_google_daily, gsa_daily_files, media_type='Google_SA')
                if gsa_raw is not None:
                    raw_dfs.append(gsa_raw)

                gda_raw = load_or_empty(parse_google_daily, gda_daily_files, media_type='Google_DA')
                if gda_raw is not None:
                    raw_dfs.append(gda_raw)

                if not raw_dfs:
                    st.error("❌ 유효한 데이터가 없습니다. 파일명을 확인해주세요.")
                    st.info("파일명에 '주간리포트', '게재된 시점' 키워드가 포함되어야 합니다.")
                    st.stop()

                raw_df = pd.concat(raw_dfs, ignore_index=True)
                raw_df = raw_df.sort_values(['날짜', '매체', '디바이스']).reset_index(drop=True)

                # 부가 데이터 파싱
                naver_kw   = load_or_empty(parse_naver_keywords, naver_kw_files)
                gsa_kw     = load_or_empty(parse_google_keywords, gsa_kw_files)
                gda_kw     = load_or_empty(parse_google_keywords, gda_kw_files)
                srch_df    = load_or_empty(parse_google_search_terms, gsa_srch_files)

                sa_conv_df = load_or_empty(parse_google_conversions, gsa_conv_files)
                if sa_conv_df is not None:
                    sa_conv_df = sa_conv_df.groupby('전환유형', as_index=False)['전환수'].sum()

                da_conv_df = load_or_empty(parse_google_conversions, gda_conv_files)
                if da_conv_df is not None:
                    da_conv_df = da_conv_df.groupby('전환유형', as_index=False)['전환수'].sum()

                # 전월 데이터 파싱 (있을 경우)
                prev_raw_df = None
                has_prev_files = prev_naver_files or prev_gsa_files or prev_gda_files
                if has_prev_files:
                    prev_dir_n  = Path(tmp) / "prev_naver"
                    prev_dir_gs = Path(tmp) / "prev_gsa"
                    prev_dir_gd = Path(tmp) / "prev_gda"
                    for d in [prev_dir_n, prev_dir_gs, prev_dir_gd]:
                        d.mkdir(exist_ok=True)
                    for f in (prev_naver_files or []):
                        (prev_dir_n / f.name).write_bytes(f.read())
                    for f in (prev_gsa_files or []):
                        (prev_dir_gs / f.name).write_bytes(f.read())
                    for f in (prev_gda_files or []):
                        (prev_dir_gd / f.name).write_bytes(f.read())

                    prev_dfs = []
                    prev_naver_raw = load_or_empty(parse_naver_weekly, find_csv(prev_dir_n, '주간리포트'))
                    if prev_naver_raw is not None: prev_dfs.append(prev_naver_raw)
                    prev_gsa_raw = load_or_empty(parse_google_daily, find_csv(prev_dir_gs, '게재된 시점'), media_type='Google_SA')
                    if prev_gsa_raw is not None: prev_dfs.append(prev_gsa_raw)
                    prev_gda_raw = load_or_empty(parse_google_daily, find_csv(prev_dir_gd, '게재된 시점'), media_type='Google_DA')
                    if prev_gda_raw is not None: prev_dfs.append(prev_gda_raw)
                    if prev_dfs:
                        prev_raw_df = pd.concat(prev_dfs, ignore_index=True)

                # 기간 라벨 생성
                dates = pd.to_datetime(raw_df['날짜'])
                period_label = f'{dates.min().strftime("%Y.%m.%d")} ~ {dates.max().strftime("%Y.%m.%d")}'

                # 출력 파일 경로
                today = datetime.now().strftime('%Y%m%d')
                xlsx_path = str(Path(tmp) / f'광고성과_장표_{today}.xlsx')
                html_path = str(Path(tmp) / f'광고성과_대시보드_{today}.html')

                # Excel 장표 생성
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
                    sns_tracker_path=sns_path,
                )

                # HTML 대시보드 생성
                build_html_report(
                    raw_df=raw_df,
                    sa_conv_df=sa_conv_df,
                    da_conv_df=da_conv_df,
                    period_label=period_label,
                    output_path=html_path,
                    sns_tracker_path=sns_path,
                    prev_raw_df=prev_raw_df,
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

    | 매체 | 파일 종류 | 파일명에 포함될 키워드 |
    |------|----------|----------------------|
    | 네이버 SA | 일별 성과 | `주간리포트` |
    | 네이버 SA | 키워드 성과 | `키워드별` |
    | 구글 SA | 일별 성과 | `게재된 시점` |
    | 구글 SA | 키워드 | `검색 키워드` |
    | 구글 SA | 전환 | `전환` |
    | 구글 SA | 검색어 | `검색어` |
    | 구글 DA | 일별 성과 | `게재된 시점` |
    | 구글 DA | 키워드 | `검색 키워드` |
    | 구글 DA | 전환 | `전환` |
    """)
