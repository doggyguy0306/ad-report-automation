"""
광고 성과 HTML 대시보드 생성 모듈
인터넷 연결 없이 브라우저에서 바로 열 수 있는 자기완결형 HTML
차트: inline SVG (외부 라이브러리 불필요)
"""

import math
import openpyxl
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

# ── 색상 ────────────────────────────────────
C_NAVER  = '#1B6B3A'
C_GOOGLE = '#4285F4'
C_DA     = '#C0392B'
C_NAVY   = '#1E3A5F'
C_GOLD   = '#F4B400'
C_LIGHT  = '#F4F7FB'
C_BORDER = '#E2E8F0'
C_TEXT   = '#2D3748'
C_MUTED  = '#718096'
C_WHITE  = '#FFFFFF'
C_PCT_BG = '#FFFBEB'

MEDIA_COLORS = {
    'Naver': C_NAVER,
    'Google_SA': C_GOOGLE,
    'Google_DA': C_DA,
}

# ── 포맷 헬퍼 ───────────────────────────────
def _f(n):
    try: return f'{int(float(n)):,}'
    except: return str(n)

def _fp(num, den):
    try:
        d = float(den)
        return f'{float(num)/d*100:.2f}%' if d > 0 else '-'
    except: return '-'

def _pv(num, den):
    try:
        d = float(den)
        return float(num)/d if d > 0 else 0.0
    except: return 0.0

def _cd(conv_df):
    if conv_df is None or conv_df.empty: return {}
    d = {}
    for _, r in conv_df.iterrows():
        d[str(r.get('전환유형',''))] = d.get(str(r.get('전환유형','')), 0) + float(r.get('전환수', 0))
    return d


# ════════════════════════════════════════════
# SVG 차트 생성
# ════════════════════════════════════════════

def _svg_line(dates, values, color, title=''):
    """일별 라인 + 수치 표시 SVG"""
    W, H = 740, 310
    PL, PR, PT, PB = 75, 20, 50, 65
    cw, ch = W - PL - PR, H - PT - PB
    n = len(values)
    if n == 0:
        return f'<svg width="{W}" height="{H}"><text x="50%" y="50%" text-anchor="middle" fill="#aaa">데이터 없음</text></svg>'

    max_v = max(values) * 1.1 if max(values) > 0 else 1
    min_v = max(0, min(values) * 0.9)
    rng = max_v - min_v or 1

    def px(i): return PL + (i / (n - 1)) * cw if n > 1 else PL + cw / 2
    def py(v): return PT + (1 - (v - min_v) / rng) * ch

    # 격자선 + Y레이블 (4단계)
    grid = ''
    for k in range(5):
        gv = min_v + rng * k / 4
        gy = py(gv)
        grid += (f'<line x1="{PL}" y1="{gy:.1f}" x2="{PL+cw}" y2="{gy:.1f}" '
                 f'stroke="#EEEEEE" stroke-width="1"/>'
                 f'<text x="{PL-6}" y="{gy+4:.1f}" text-anchor="end" '
                 f'font-size="10" fill="{C_MUTED}">{_f(gv)}</text>')

    # 면적 채우기
    area_pts = f'{PL:.1f},{PT+ch:.1f}'
    for i, v in enumerate(values): area_pts += f' {px(i):.1f},{py(v):.1f}'
    area_pts += f' {px(n-1):.1f},{PT+ch:.1f}'

    # 선
    line_pts = ' '.join(f'{px(i):.1f},{py(v):.1f}' for i, v in enumerate(values))

    # 포인트 + 수치 레이블
    pts_svg = ''
    for i, v in enumerate(values):
        x, y = px(i), py(v)
        lbl = _f(v)
        pts_svg += (f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4.5" fill="{color}" '
                    f'stroke="white" stroke-width="1.5"/>'
                    f'<text x="{x:.1f}" y="{y-11:.1f}" text-anchor="middle" '
                    f'font-size="10" font-weight="600" fill="{C_TEXT}">{lbl}</text>')

    # X레이블 (최대 8개)
    step = max(1, n // 8)
    x_lbl = ''
    for i in range(0, n, step):
        d_lbl = str(dates[i])[-5:]
        xp = px(i)
        x_lbl += (f'<text x="{xp:.1f}" y="{PT+ch+17}" text-anchor="end" '
                  f'font-size="10" fill="{C_MUTED}" '
                  f'transform="rotate(-35 {xp:.1f} {PT+ch+17})">{d_lbl}</text>')

    axes = (f'<line x1="{PL}" y1="{PT}" x2="{PL}" y2="{PT+ch}" stroke="#CCC" stroke-width="1"/>'
            f'<line x1="{PL}" y1="{PT+ch}" x2="{PL+cw}" y2="{PT+ch}" stroke="#CCC" stroke-width="1"/>')

    return (f'<svg viewBox="0 0 {W} {H}" width="100%" xmlns="http://www.w3.org/2000/svg">'
            f'<rect width="{W}" height="{H}" fill="white" rx="6"/>'
            f'<text x="{W/2}" y="30" text-anchor="middle" font-size="14" font-weight="bold" fill="{C_NAVY}">{title}</text>'
            f'{grid}{axes}'
            f'<polygon points="{area_pts}" fill="{color}" opacity="0.10"/>'
            f'<polyline points="{line_pts}" fill="none" stroke="{color}" stroke-width="2.5" stroke-linejoin="round"/>'
            f'{pts_svg}{x_lbl}</svg>')


def _svg_bar_grouped(group_labels, datasets, title=''):
    """grouped bar chart SVG
    datasets: [(name, [values], color), ...]
    """
    W, H = 740, 310
    PL, PR, PT, PB = 75, 20, 50, 55
    cw, ch = W - PL - PR, H - PT - PB
    n, nd = len(group_labels), len(datasets)
    if n == 0 or nd == 0:
        return ''

    all_vals = [v for _, vals, _ in datasets for v in vals]
    max_v = max(all_vals) * 1.1 if all_vals and max(all_vals) > 0 else 1

    grid = ''
    for k in range(5):
        gv = max_v * k / 4
        gy = PT + ch - (gv / max_v) * ch
        grid += (f'<line x1="{PL}" y1="{gy:.1f}" x2="{PL+cw}" y2="{gy:.1f}" '
                 f'stroke="#EEEEEE" stroke-width="1"/>'
                 f'<text x="{PL-6}" y="{gy+4:.1f}" text-anchor="end" '
                 f'font-size="10" fill="{C_MUTED}">{_f(gv)}</text>')

    group_w = cw / n
    bar_w   = group_w * 0.7 / nd
    pad     = group_w * 0.15
    bars    = ''

    for gi, g_lbl in enumerate(group_labels):
        gx = PL + gi * group_w + pad
        for di, (_, vals, color) in enumerate(datasets):
            v  = vals[gi] if gi < len(vals) else 0
            bh = (v / max_v) * ch if max_v > 0 else 0
            bx = gx + di * bar_w
            by = PT + ch - bh
            bars += (f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bar_w:.1f}" '
                     f'height="{bh:.1f}" fill="{color}" rx="3"/>'
                     f'<text x="{bx + bar_w/2:.1f}" y="{max(PT+5, by-4):.1f}" '
                     f'text-anchor="middle" font-size="9" font-weight="bold" fill="{color}">{_f(v)}</text>')
        # X label
        lx = PL + gi * group_w + group_w / 2
        bars += (f'<text x="{lx:.1f}" y="{PT+ch+18}" text-anchor="middle" '
                 f'font-size="11" fill="{C_TEXT}">{g_lbl}</text>')

    # 범례
    legend = ''
    total_leg_w = nd * 110
    lx0 = (W - total_leg_w) / 2
    for di, (name, _, color) in enumerate(datasets):
        lx = lx0 + di * 110
        legend += (f'<rect x="{lx:.1f}" y="{H-20}" width="12" height="12" fill="{color}" rx="2"/>'
                   f'<text x="{lx+16:.1f}" y="{H-10}" font-size="11" fill="{C_TEXT}">{name}</text>')

    axes = (f'<line x1="{PL}" y1="{PT}" x2="{PL}" y2="{PT+ch}" stroke="#CCC" stroke-width="1"/>'
            f'<line x1="{PL}" y1="{PT+ch}" x2="{PL+cw}" y2="{PT+ch}" stroke="#CCC" stroke-width="1"/>')

    return (f'<svg viewBox="0 0 {W} {H}" width="100%" xmlns="http://www.w3.org/2000/svg">'
            f'<rect width="{W}" height="{H}" fill="white" rx="6"/>'
            f'<text x="{W/2}" y="30" text-anchor="middle" font-size="14" font-weight="bold" fill="{C_NAVY}">{title}</text>'
            f'{grid}{axes}{bars}{legend}</svg>')


def _svg_pie(labels, values, colors, title=''):
    """파이 차트 SVG"""
    W, H = 420, 310
    total = sum(values) if values else 1
    if total == 0: return ''
    cx, cy, r = W * 0.40, H * 0.55, min(W * 0.38, H * 0.42)

    slices, leg = '', ''
    angle = -math.pi / 2

    for i, (lbl, v, color) in enumerate(zip(labels, values, colors)):
        pct  = v / total
        a2   = angle + 2 * math.pi * pct
        x1 = cx + r * math.cos(angle); y1 = cy + r * math.sin(angle)
        x2 = cx + r * math.cos(a2);   y2 = cy + r * math.sin(a2)
        large = 1 if pct > 0.5 else 0
        slices += (f'<path d="M{cx:.1f},{cy:.1f} L{x1:.1f},{y1:.1f} '
                   f'A{r},{r} 0 {large},1 {x2:.1f},{y2:.1f} Z" '
                   f'fill="{color}" stroke="white" stroke-width="2"/>')
        if pct > 0.04:
            mid_a = (angle + a2) / 2
            lx = cx + r * 0.62 * math.cos(mid_a)
            ly = cy + r * 0.62 * math.sin(mid_a)
            slices += (f'<text x="{lx:.1f}" y="{ly:.1f}" text-anchor="middle" '
                       f'font-size="11" font-weight="bold" fill="white">{pct*100:.1f}%</text>')
        # 범례
        ly_l = 55 + i * 24
        leg += (f'<rect x="{W*0.80:.1f}" y="{ly_l}" width="12" height="12" fill="{color}" rx="2"/>'
                f'<text x="{W*0.80+16:.1f}" y="{ly_l+10}" font-size="11" fill="{C_TEXT}">{lbl[:10]}</text>')
        angle = a2

    return (f'<svg viewBox="0 0 {W} {H}" width="100%" xmlns="http://www.w3.org/2000/svg">'
            f'<rect width="{W}" height="{H}" fill="white" rx="6"/>'
            f'<text x="{cx}" y="28" text-anchor="middle" font-size="14" font-weight="bold" fill="{C_NAVY}">{title}</text>'
            f'{slices}{leg}</svg>')


# ════════════════════════════════════════════
# HTML 섹션 빌더
# ════════════════════════════════════════════

def _kpi_card(label, value, sub='', color=C_NAVY):
    return (f'<div class="kpi-card" style="border-top:4px solid {color}">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value" style="color:{color}">{value}</div>'
            f'<div class="kpi-sub">{sub}</div>'
            f'</div>')


def _table_html(headers, rows, stripe=True):
    """범용 HTML 테이블"""
    ths = ''.join(f'<th>{h}</th>' for h in headers)
    trs = ''
    for i, row in enumerate(rows):
        cls = ' class="stripe"' if stripe and i % 2 == 1 else ''
        tds = ''.join(
            f'<td class="pct">{c}</td>' if isinstance(c, str) and c.endswith('%')
            else f'<td>{c}</td>'
            for c in row
        )
        trs += f'<tr{cls}>{tds}</tr>'
    return f'<table><thead><tr>{ths}</tr></thead><tbody>{trs}</tbody></table>'


def _section(title, color, content):
    return (f'<div class="section">'
            f'<div class="section-header" style="background:{color}">{title}</div>'
            f'<div class="section-body">{content}</div>'
            f'</div>')


def _chart_block(title, svg):
    return (f'<div class="chart-block">'
            f'<div class="chart-title">{title}</div>'
            f'{svg}'
            f'</div>')


# ════════════════════════════════════════════
# SNS 채널 데이터 읽기 / HTML 섹션
# ════════════════════════════════════════════

# 채널별 메타 정보
_SNS_META = {
    '📱 인스타그램': {
        'color': '#E1306C', 'emoji': '📱',
        'table_cols': ['날짜', '팔로워수', '게시물수\n(당일)', '좋아요수', '댓글수', '도달수'],
        'col_idx':    [0,      1,           3,                  5,         6,        7],
        'main_col': 1,  # 팔로워수
    },
    '🧵 쓰레드': {
        'color': '#444444', 'emoji': '🧵',
        'table_cols': ['날짜', '팔로워수', '게시물수\n(당일)', '조회수', '좋아요수', '댓글수'],
        'col_idx':    [0,      1,           3,                  5,        6,          7],
        'main_col': 1,
    },
    '📝 티스토리': {
        'color': '#FF6600', 'emoji': '📝',
        'table_cols': ['날짜', '방문자수(UV)', '페이지뷰(PV)', '게시물수\n(당일)', '댓글수', '공감수'],
        'col_idx':    [0,      1,              2,               3,                  5,        6],
        'main_col': 1,  # 방문자수
    },
    '💬 카카오톡채널': {
        'color': '#C8A800', 'emoji': '💬',
        'table_cols': ['날짜', '친구수(누적)', '신규친구수', '메시지발송수', '메시지열람수', '열람율(%)'],
        'col_idx':    [0,      1,              2,             4,              5,              6],
        'main_col': 1,  # 친구수
    },
    '▶ 유튜브': {
        'color': '#CC0000', 'emoji': '▶',
        'table_cols': ['날짜', '구독자수', '조회수', '시청시간(시간)', '좋아요수', '댓글수'],
        'col_idx':    [0,      1,           3,        4,                5,          6],
        'main_col': 1,  # 구독자수
    },
}

_SNS_CHANNEL_ORDER = [
    '📱 인스타그램', '🧵 쓰레드', '📝 티스토리', '💬 카카오톡채널', '▶ 유튜브'
]


def _read_sns_channel_data(sns_path: str) -> dict:
    """SNS 관리대장 xlsx 에서 각 채널의 입력 데이터를 읽어 반환.
    반환:
        { channel_name: {'headers': [...], 'rows': [[date, v1, v2, ...], ...]} }
    """
    result = {}
    try:
        wb = openpyxl.load_workbook(sns_path, data_only=True)
        for ch in _SNS_CHANNEL_ORDER:
            if ch not in wb.sheetnames:
                continue
            ws = wb[ch]
            # 헤더 row4
            raw_hdrs = [ws.cell(4, c).value for c in range(1, 12)]
            headers  = [str(h).replace('\n', '') if h else '' for h in raw_hdrs]
            # 데이터 row5~94: 날짜(col A)가 있는 행만
            rows = []
            for r in range(5, 95):
                date_val = ws.cell(r, 1).value
                if date_val is None:
                    continue
                row_vals = [ws.cell(r, c).value for c in range(1, 12)]
                rows.append(row_vals)
            result[ch] = {'headers': headers, 'rows': rows}
    except Exception:
        pass
    return result


def _sns_summary_cards(channel_data: dict) -> str:
    """채널별 메인 지표 KPI 카드 5개 (한 줄)"""
    cards = ''
    for ch in _SNS_CHANNEL_ORDER:
        meta  = _SNS_META[ch]
        color = meta['color']
        rows  = channel_data.get(ch, {}).get('rows', [])
        main_col = meta['main_col']

        if rows:
            # 마지막으로 실제 값이 입력된 행을 찾음 (빈 날짜 행 무시)
            filled = [r for r in rows if r[main_col] is not None]
            latest = filled[-1][main_col] if filled else None
            prev   = filled[-2][main_col] if len(filled) >= 2 else None
            val_str  = _f(latest) if latest is not None else '-'
            if prev is not None and latest is not None:
                try:
                    diff = float(latest) - float(prev)
                    sign = '+' if diff >= 0 else ''
                    delta_str = f'{sign}{_f(diff)}'
                    delta_color = '#16A34A' if diff >= 0 else '#DC2626'
                except Exception:
                    delta_str, delta_color = '-', C_MUTED
            else:
                delta_str, delta_color = '입력 대기', C_MUTED
        else:
            val_str, delta_str, delta_color = '미입력', '-', C_MUTED

        # 지표 이름 (메인 컬럼 헤더)
        headers = channel_data.get(ch, {}).get('headers', [])
        metric_name = headers[main_col] if headers and len(headers) > main_col else '주요지표'
        ch_name = ch[2:].strip()  # 이모지 제거

        cards += (
            f'<div class="sns-kpi" style="border-top:4px solid {color}">'
            f'<div class="sns-ch-name" style="color:{color}">{ch}</div>'
            f'<div class="sns-metric-label">{metric_name}</div>'
            f'<div class="sns-value" style="color:{color}">{val_str}</div>'
            f'<div class="sns-delta" style="color:{delta_color}">전일대비 {delta_str}</div>'
            f'</div>'
        )
    return f'<div class="sns-kpi-grid">{cards}</div>'


def _sns_channel_detail(ch: str, data: dict) -> str:
    """채널 1개의 최근 데이터 테이블 + 미니 트렌드 차트"""
    meta    = _SNS_META[ch]
    color   = meta['color']
    rows    = data.get('rows', [])
    headers = data.get('headers', [])
    col_idx = meta['col_idx']
    tbl_hdrs = meta['table_cols']
    main_col = meta['main_col']

    if not rows:
        placeholder = (
            f'<div style="text-align:center;padding:20px;color:{C_MUTED};font-size:12px">'
            f'아직 데이터가 입력되지 않았습니다.<br>'
            f'<b>SNS_채널_관리대장.xlsx</b> 의 {ch} 탭에 날짜와 수치를 입력하세요.</div>'
        )
        return placeholder

    # 값이 입력된 행만 필터링 후 최근 14일치 표시
    filled_rows = [r for r in rows if any(r[ci] is not None for ci in col_idx[1:])]
    recent = filled_rows[-14:] if filled_rows else rows[-14:]

    # 미니 트렌드 SVG (메인 지표)
    trend_vals = []
    for r in recent:
        v = r[main_col] if main_col < len(r) else None
        try:
            trend_vals.append(float(v)) if v is not None else trend_vals.append(None)
        except Exception:
            trend_vals.append(None)

    # None 값을 건너뛰고 유효 데이터만
    valid_pairs = [(i, v) for i, v in enumerate(trend_vals) if v is not None]
    if len(valid_pairs) >= 2:
        n = len(recent)
        W, H = 700, 120
        PL, PR, PT, PB = 10, 10, 20, 20
        cw, ch_h = W-PL-PR, H-PT-PB
        vmin = min(v for _, v in valid_pairs)
        vmax = max(v for _, v in valid_pairs)
        rng  = vmax - vmin or 1
        def px(i): return PL + (i/(n-1))*cw if n > 1 else PL+cw/2
        def py(v): return PT + (1-(v-vmin)/rng)*ch_h
        pts = ' '.join(f'{px(i):.1f},{py(v):.1f}' for i,v in valid_pairs)
        area_pts = f'{px(valid_pairs[0][0]):.1f},{PT+ch_h:.1f} '
        area_pts += pts
        area_pts += f' {px(valid_pairs[-1][0]):.1f},{PT+ch_h:.1f}'
        # 최신값 레이블
        li, lv = valid_pairs[-1]
        mini_svg = (
            f'<svg viewBox="0 0 {W} {H}" width="100%" style="display:block;margin-bottom:8px">'
            f'<rect width="{W}" height="{H}" fill="white"/>'
            f'<polygon points="{area_pts}" fill="{color}" opacity="0.12"/>'
            f'<polyline points="{pts}" fill="none" stroke="{color}" stroke-width="2.5" stroke-linejoin="round"/>'
            f'<circle cx="{px(li):.1f}" cy="{py(lv):.1f}" r="5" fill="{color}" stroke="white" stroke-width="2"/>'
            f'<text x="{px(li):.1f}" y="{max(PT+2, py(lv)-8):.1f}" text-anchor="middle" '
            f'font-size="11" font-weight="bold" fill="{color}">{_f(lv)}</text>'
            f'</svg>'
        )
    else:
        mini_svg = ''

    # 테이블
    tbl_rows = []
    for r in reversed(recent):  # 최신순
        tbl_row = []
        for ci in col_idx:
            v = r[ci] if ci < len(r) else None
            if ci == 0 and v is not None:  # 날짜
                try:
                    tbl_row.append(pd.Timestamp(v).strftime('%Y-%m-%d'))
                except Exception:
                    tbl_row.append(str(v))
            elif isinstance(v, float) and v == int(v):
                tbl_row.append(f'{int(v):,}')
            elif isinstance(v, (int, float)):
                tbl_row.append(_f(v))
            else:
                tbl_row.append(str(v) if v is not None else '-')
        tbl_rows.append(tbl_row)

    tbl_hdrs_clean = [h.replace('\n', ' ') for h in tbl_hdrs]
    tbl_html = _table_html(tbl_hdrs_clean, tbl_rows)

    return mini_svg + tbl_html


def _build_sns_section(sns_path: str) -> str:
    """SNS 전체 섹션 HTML 반환"""
    channel_data = _read_sns_channel_data(sns_path)

    # 요약 카드
    summary_html = _sns_summary_cards(channel_data)

    # 채널별 상세 (2열 그리드)
    detail_html = '<div class="sns-grid">'
    for ch in _SNS_CHANNEL_ORDER:
        meta  = _SNS_META[ch]
        color = meta['color']
        data  = channel_data.get(ch, {'rows': [], 'headers': []})
        rows  = data.get('rows', [])
        count = len(rows)
        badge = f'<span class="sns-badge">{count}일 입력</span>' if count else \
                '<span class="sns-badge sns-badge-empty">미입력</span>'
        detail = _sns_channel_detail(ch, data)
        detail_html += (
            f'<div class="sns-channel-card">'
            f'<div class="sns-ch-header" style="background:{color}">'
            f'  {ch} {badge}'
            f'</div>'
            f'<div class="sns-ch-body">{detail}</div>'
            f'</div>'
        )
    detail_html += '</div>'

    return summary_html + detail_html


# ════════════════════════════════════════════
# 메인 빌더
# ════════════════════════════════════════════

def build_html_report(raw_df, sa_conv_df=None, da_conv_df=None,
                      period_label='', output_path='report.html',
                      sns_tracker_path=None):

    sa_cd = _cd(sa_conv_df)
    da_cd = _cd(da_conv_df)
    today = datetime.now().strftime('%Y-%m-%d %H:%M')

    # ── 전체 집계 ──────────────────────────
    total_노출 = raw_df['노출'].sum()
    total_클릭 = raw_df['클릭'].sum()
    total_비용 = raw_df['비용'].sum()

    grp = raw_df.groupby('매체').agg(
        노출=('노출','sum'), 클릭=('클릭','sum'),
        전환=('전환','sum'), 비용=('비용','sum')
    ).reset_index()

    naver_df  = raw_df[raw_df['매체'].str.startswith('Naver')]
    gsa_df    = raw_df[raw_df['매체'] == 'Google_SA']
    gda_df    = raw_df[raw_df['매체'] == 'Google_DA']
    google_df = raw_df[raw_df['매체'].str.startswith('Google')]

    naver_클릭  = naver_df['클릭'].sum()
    gsa_클릭    = gsa_df['클릭'].sum()
    gda_클릭    = gda_df['클릭'].sum()
    google_클릭 = google_df['클릭'].sum()

    # 네이버: 별도 전환파일 없음 → raw 데이터 '총 전환수' 사용
    naver_버튼전환 = naver_df['전환'].sum()
    sa_버튼전환    = sum(sa_cd.values()) - sa_cd.get('페이지조회', 0)
    da_버튼전환    = sum(da_cd.values()) - da_cd.get('페이지조회', 0)

    # ── KPI 카드 ───────────────────────────
    kpi_html = (
        _kpi_card('전체 노출', _f(total_노출), '기간 합계', C_NAVY) +
        _kpi_card('전체 클릭', _f(total_클릭), f'CTR {_fp(total_클릭, total_노출)}', C_GOOGLE) +
        _kpi_card('전체 광고비', f'₩{_f(total_비용)}', f'CPC ₩{_f(total_비용/total_클릭 if total_클릭 else 0)}', C_DA) +
        _kpi_card('네이버 클릭', _f(naver_클릭), f'전체 클릭비중 {_fp(naver_클릭, total_클릭)}', C_NAVER) +
        _kpi_card('구글 SA 클릭', _f(gsa_클릭), f'전체 클릭비중 {_fp(gsa_클릭, total_클릭)}', C_GOOGLE) +
        _kpi_card('구글 DA 클릭', _f(gda_클릭), f'전체 클릭비중 {_fp(gda_클릭, total_클릭)}', C_DA)
    )
    kpi_section = f'<div class="kpi-grid">{kpi_html}</div>'

    # 매체별 버튼 전환 매핑
    _btn_map = {}
    for _, r in grp.iterrows():
        m = str(r['매체'])
        if m.startswith('Naver'):
            _btn_map[m] = naver_버튼전환
        elif m == 'Google_SA':
            _btn_map[m] = sa_버튼전환
        elif m == 'Google_DA':
            _btn_map[m] = da_버튼전환
        else:
            _btn_map[m] = 0

    # ── 플랫폼 비교 테이블 ─────────────────
    comp_rows = []
    for _, r in grp.iterrows():
        노출 = float(r['노출']); 클릭 = float(r['클릭']); 비용 = float(r['비용'])
        btn  = _btn_map.get(str(r['매체']), 0)
        comp_rows.append([
            r['매체'],
            _f(노출), _f(클릭), _fp(클릭, 노출),
            f'₩{_f(비용/클릭 if 클릭 else 0)}', f'₩{_f(비용)}',
            _fp(노출, total_노출), _fp(클릭, total_클릭), _fp(비용, total_비용),
            f'{_f(btn)}건 ({_fp(btn, 클릭)})',
        ])
    comp_tbl = _table_html(
        ['매체','노출','클릭','CTR','CPC','비용','노출비중','클릭비중','비용비중','버튼 전환'],
        comp_rows
    )

    # ── 플랫폼 비교 차트 ───────────────────
    media_labels = list(grp['매체'])
    click_vals   = [float(v) for v in grp['클릭']]
    cost_vals    = [float(v) for v in grp['비용']]
    media_colors = [MEDIA_COLORS.get(m.split('_')[0] if '_' in m else m, C_NAVY) for m in media_labels]

    svg_cmp_click = _svg_bar_grouped(
        media_labels,
        [('클릭수', click_vals, C_GOOGLE)],
        title='매체별 클릭수 비교'
    )
    svg_cost_pie = _svg_pie(
        media_labels, cost_vals, media_colors,
        title='매체별 비용 비중'
    )

    compare_content = (
        f'<div class="two-col">'
        f'<div class="col-main">{comp_tbl}</div>'
        f'</div>'
        f'<div class="chart-row">'
        f'<div class="chart-item">{_chart_block("매체별 클릭수 비교", svg_cmp_click)}</div>'
        f'<div class="chart-item chart-narrow">{_chart_block("매체별 비용 비중", svg_cost_pie)}</div>'
        f'</div>'
    )

    # ── 매체별 일별 트렌드 ─────────────────
    def _daily_block(df, media_label, color):
        if df.empty: return ''
        daily = df.groupby('날짜', as_index=False).agg(
            노출=('노출','sum'), 클릭=('클릭','sum'), 비용=('비용','sum')
        ).sort_values('날짜')
        daily['날짜'] = pd.to_datetime(daily['날짜']).dt.strftime('%Y-%m-%d')
        dates  = list(daily['날짜'])
        clicks = [float(v) for v in daily['클릭']]
        costs  = [float(v) for v in daily['비용']]

        svg_c = _svg_line(dates, clicks, color, title='일별 클릭수 추이')
        svg_v = _svg_line(dates, costs,  color, title='일별 비용 추이')

        # 일별 테이블
        rows = []
        for _, r in daily.iterrows():
            클릭 = float(r['클릭']); 노출 = float(r['노출']); 비용 = float(r['비용'])
            rows.append([r['날짜'], _f(노출), _f(클릭), _fp(클릭, 노출),
                         f'₩{_f(비용/클릭 if 클릭 else 0)}', f'₩{_f(비용)}'])
        tbl = _table_html(['날짜','노출','클릭','CTR','CPC','비용'], rows)

        return (
            f'<div class="chart-row">'
            f'<div class="chart-item">{_chart_block("일별 클릭수 추이", svg_c)}</div>'
            f'<div class="chart-item">{_chart_block("일별 비용 추이", svg_v)}</div>'
            f'</div>'
            f'<div class="sub-title">일별 성과 상세</div>{tbl}'
        )

    naver_content  = _daily_block(naver_df,  '네이버 SA',  C_NAVER)
    gsa_content    = _daily_block(gsa_df,    '구글 SA',    C_GOOGLE)
    gda_content    = _daily_block(gda_df,    '구글 DA',    C_DA)

    # 네이버 버튼 전환 블록 (총 전환수 기준)
    naver_conv_html = (
        f'<div class="sub-title">네이버 SA 버튼 전환 (총 전환수 기준)</div>'
        + _table_html(
            ['전환 기준', '전환수', '전환율(클릭 대비)'],
            [['총 전환수 (상담예약·신청·상세보기 통합)',
              f'{_f(naver_버튼전환)}건',
              _fp(naver_버튼전환, naver_클릭)]],
            stripe=False
        )
        + '<p style="font-size:11px;color:#718096;margin-top:6px">'
          '※ 네이버는 버튼별 세분화 미설정으로 총 전환수를 버튼 전환으로 사용합니다.</p>'
    )

    # ── 버튼 전환 섹션 ─────────────────────
    def _conv_block(cd, 클릭수, label, color):
        if not cd: return ''
        상세 = cd.get('상세보기_버튼클릭', 0)
        예약 = cd.get('상담예약_버튼클릭', 0)
        신청 = cd.get('상담신청_버튼클릭', 0)
        rows = [
            ['상세보기 버튼클릭', _f(상세), _fp(상세, 클릭수)],
            ['상담예약 버튼클릭', _f(예약), _fp(예약, 클릭수)],
            ['상담신청 버튼클릭', _f(신청), _fp(신청, 클릭수)],
            ['버튼클릭 합계',     _f(상세+예약+신청), _fp(상세+예약+신청, 클릭수)],
        ]
        tbl = _table_html(['전환 유형', '전환수', f'전환율({label} 클릭 대비)'], rows)
        return f'<div class="sub-title">{label} 버튼 전환</div>{tbl}'

    sa_conv_content = _conv_block(sa_cd, gsa_클릭, 'SA', C_GOOGLE)
    da_conv_content = _conv_block(da_cd, gda_클릭, 'DA', C_DA)

    # SA vs DA 비교 바차트
    btn_labels = ['상세보기', '상담예약', '상담신청']
    sa_vals = [sa_cd.get('상세보기_버튼클릭',0), sa_cd.get('상담예약_버튼클릭',0), sa_cd.get('상담신청_버튼클릭',0)]
    da_vals = [da_cd.get('상세보기_버튼클릭',0), da_cd.get('상담예약_버튼클릭',0), da_cd.get('상담신청_버튼클릭',0)]
    svg_conv_bar = _svg_bar_grouped(
        btn_labels,
        [('Google SA', sa_vals, C_GOOGLE), ('Google DA', da_vals, C_DA)],
        title='SA vs DA 버튼별 전환 비교'
    )

    conv_content = (
        f'{_chart_block("SA vs DA 버튼별 전환 비교", svg_conv_bar)}'
        f'<div class="two-col-conv">'
        f'<div>{sa_conv_content}</div>'
        f'<div>{da_conv_content}</div>'
        f'</div>'
    )

    # ══════════════════════════════════════
    # CSS
    # ══════════════════════════════════════
    css = f"""
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
    font-family: 'Apple SD Gothic Neo','Malgun Gothic','맑은 고딕',Arial,sans-serif;
    background: {C_LIGHT}; color: {C_TEXT}; font-size: 13px;
}}
.page {{ max-width: 1200px; margin: 0 auto; padding: 24px 16px; }}

/* 헤더 */
.header {{ background: {C_NAVY}; color: white; border-radius: 10px;
    padding: 20px 28px; margin-bottom: 20px; }}
.header h1 {{ font-size: 20px; font-weight: 700; }}
.header .meta {{ font-size: 12px; opacity: 0.7; margin-top: 6px; }}

/* KPI 카드 */
.kpi-grid {{ display: grid; grid-template-columns: repeat(6,1fr); gap: 12px; margin-bottom: 20px; }}
.kpi-card {{ background: white; border-radius: 8px; padding: 14px 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
.kpi-label {{ font-size: 11px; color: {C_MUTED}; font-weight: 500; margin-bottom: 6px; }}
.kpi-value {{ font-size: 20px; font-weight: 700; }}
.kpi-sub {{ font-size: 11px; color: {C_MUTED}; margin-top: 4px; }}

/* 섹션 */
.section {{ background: white; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,.08);
    margin-bottom: 20px; overflow: hidden; }}
.section-header {{ color: white; font-weight: 700; font-size: 14px;
    padding: 12px 20px; }}
.section-body {{ padding: 20px; }}

/* 테이블 */
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
thead tr {{ background: {C_NAVY}; color: white; }}
th {{ padding: 9px 10px; text-align: center; font-weight: 600; white-space: nowrap; }}
td {{ padding: 8px 10px; text-align: center; border-bottom: 1px solid {C_BORDER}; }}
tr.stripe {{ background: #F8FAFC; }}
td.pct {{ background: #FFFBEB; font-weight: 600; color: #92400E; }}
.sub-title {{ font-size: 13px; font-weight: 700; color: {C_NAVY};
    margin: 16px 0 8px; padding-left: 10px; border-left: 4px solid {C_NAVY}; }}

/* 차트 */
.chart-row {{ display: flex; gap: 16px; margin: 16px 0; }}
.chart-item {{ flex: 1; min-width: 0; background: white;
    border: 1px solid {C_BORDER}; border-radius: 8px; overflow: hidden; }}
.chart-narrow {{ max-width: 380px; flex: 0 0 380px; }}
.chart-block {{ width: 100%; padding: 12px; }}
.chart-title {{ font-size: 12px; font-weight: 600; color: {C_MUTED};
    text-align: center; margin-bottom: 8px; }}
.two-col-conv {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-top: 16px; }}

/* SNS 섹션 */
.sns-kpi-grid {{ display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin-bottom:20px; }}
.sns-kpi {{ background:white; border-radius:8px; padding:14px 16px;
    box-shadow:0 1px 4px rgba(0,0,0,.08); }}
.sns-ch-name {{ font-size:12px; font-weight:700; margin-bottom:4px; }}
.sns-metric-label {{ font-size:11px; color:{C_MUTED}; margin-bottom:6px; }}
.sns-value {{ font-size:22px; font-weight:800; margin-bottom:4px; }}
.sns-delta {{ font-size:11px; font-weight:600; }}
.sns-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
.sns-channel-card {{ background:white; border-radius:8px;
    box-shadow:0 1px 4px rgba(0,0,0,.08); overflow:hidden; }}
.sns-ch-header {{ color:white; font-weight:700; font-size:13px;
    padding:10px 16px; display:flex; align-items:center; justify-content:space-between; }}
.sns-ch-body {{ padding:14px 16px; }}
.sns-badge {{ background:rgba(255,255,255,0.25); border-radius:10px;
    padding:2px 10px; font-size:11px; font-weight:600; }}
.sns-badge-empty {{ background:rgba(0,0,0,0.20); }}

/* 반응형 */
@media (max-width: 900px) {{
    .kpi-grid {{ grid-template-columns: repeat(3,1fr); }}
    .chart-row {{ flex-direction: column; }}
    .chart-narrow {{ max-width: 100%; flex: 1; }}
    .two-col-conv {{ grid-template-columns: 1fr; }}
    .sns-kpi-grid {{ grid-template-columns: repeat(2,1fr); }}
    .sns-grid {{ grid-template-columns: 1fr; }}
}}
"""

    # ══════════════════════════════════════
    # HTML 조립
    # ══════════════════════════════════════
    html_body = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>광고 성과 대시보드 | {period_label}</title>
<style>{css}</style>
</head>
<body>
<div class="page">

  <div class="header">
    <h1>📊 광고 성과 대시보드</h1>
    <div class="meta">기간: {period_label} &nbsp;|&nbsp; 생성: {today}</div>
  </div>

  {kpi_section}

  {_section('🔍 전체 매체 비교', C_NAVY, compare_content)}

  {_section('🟢 네이버 SA', C_NAVER, naver_content + naver_conv_html)}

  {_section('🔵 구글 SA (검색)', C_GOOGLE, gsa_content + sa_conv_content)}

  {_section('🔴 구글 DA (실적 최대화)', C_DA, gda_content + da_conv_content)}

  {_section('🎯 버튼 전환 상세 (SA / DA 비교)', C_GOOGLE, conv_content)}

  {_section('📱 SNS 채널 현황', '#6C3483',
      _build_sns_section(sns_tracker_path) if sns_tracker_path else
      '<div style="text-align:center;padding:24px;color:#888;font-size:13px">'
      '광고 보고서 생성 시 <b>SNS_채널_관리대장.xlsx</b>를 함께 제공하면 이 섹션이 자동으로 표시됩니다.</div>'
  )}

</div>
</body>
</html>"""

    Path(output_path).write_text(html_body, encoding='utf-8')
    print(f'✅ HTML 대시보드 저장: {output_path}')
    return output_path
