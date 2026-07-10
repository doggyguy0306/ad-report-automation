// KT M&S 양식 기반 광고 성과 보고서 PPT 생성
// 2026년 6월 | 채널마케팅팀

const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "광고 성과 보고서 2026년 6월";
pres.author = "kt m&s 채널마케팅팀";

// ── 색상 팔레트 (KT M&S 공식 템플릿) ──────────────────
const BG_DARK   = "1C1C1E";   // 다크 배경
const BG_LIGHT  = "F0F0F2";   // 라이트 배경
const KT_RED    = "E4002B";   // KT 레드
const NAVY      = "1A1F36";   // 다크 네이비
const DIVIDER   = "4B5563";   // 섹션 구분선 (회색)
const WHITE     = "FFFFFF";
const GRAY_TXT  = "6B7280";   // 본문 회색
const CARD_BG   = "FFFFFF";   // 카드 흰 배경
const PINK_BG   = "FEE2E2";   // 경고/노트 연한 핑크
const PINK_TXT  = "E4002B";
const FONT      = "Pretendard";  // KT M&S 공식 폰트

const TOTAL_PAGES = 11;

// ── 공통 Helper ────────────────────────────────────────

// 하단 푸터
function addFooter(s, pageNum) {
  s.addText("kt m&s · 채널마케팅팀", {
    x: 0.45, y: 5.33, w: 3, h: 0.22,
    fontSize: 9, fontFace: FONT, color: GRAY_TXT,
    align: "left", margin: 0
  });
  s.addText("사내용 · Internal Use Only", {
    x: 3.5, y: 5.33, w: 3, h: 0.22,
    fontSize: 9, fontFace: FONT, color: GRAY_TXT,
    align: "center", margin: 0
  });
  s.addText(`${String(pageNum).padStart(2,"0")} / ${String(TOTAL_PAGES).padStart(2,"0")}`, {
    x: 7.05, y: 5.33, w: 2.5, h: 0.22,
    fontSize: 9, fontFace: FONT, color: GRAY_TXT,
    align: "right", margin: 0
  });
}

// 콘텐츠 슬라이드 섹션 뱃지 + 제목
function addContentHeader(s, badge, title, pageNum) {
  // 섹션 뱃지
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: 0.28, w: 1.8, h: 0.30,
    fill: { color: CARD_BG },
    line: { color: "D1D5DB", width: 1 },
    rectRadius: 0.04
  });
  s.addText(badge, {
    x: 0.45, y: 0.28, w: 1.8, h: 0.30,
    fontSize: 10, fontFace: FONT, color: GRAY_TXT,
    align: "center", valign: "middle", margin: 0
  });
  // 제목
  s.addText(title, {
    x: 0.45, y: 0.65, w: 9.1, h: 0.62,
    fontSize: 30, fontFace: FONT, bold: true, color: NAVY,
    align: "left", margin: 0
  });
  addFooter(s, pageNum);
}

// 섹션 구분 슬라이드 (다크)
function makeSectionDivider(num, title) {
  const s = pres.addSlide();
  s.background = { color: BG_DARK };
  // 구분선
  s.addShape(pres.shapes.LINE, {
    x: 3.5, y: 2.60, w: 3.0, h: 0,
    line: { color: DIVIDER, width: 2 }
  });
  // 번호 + 제목
  s.addText(`${num}    ${title}`, {
    x: 2, y: 2.72, w: 6, h: 0.75,
    fontSize: 34, fontFace: FONT, bold: true, color: WHITE,
    align: "center", margin: 0
  });
  return s;
}

// 카드 (컬러 헤더 + 흰 바디)
function addCard(s, x, y, w, h, headerColor, headerText, bodyLines) {
  // 바디
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: CARD_BG },
    line: { color: "E5E7EB", width: 1 },
    rectRadius: 0.06
  });
  // 헤더
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 0.40,
    fill: { color: headerColor },
    rectRadius: 0.06,
    line: { color: headerColor, width: 0 }
  });
  s.addText(headerText, {
    x: x + 0.14, y, w: w - 0.28, h: 0.40,
    fontSize: 12, fontFace: FONT, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0
  });
  // 바디 텍스트
  bodyLines.forEach(([label, value, sub], i) => {
    const by = y + 0.50 + i * 0.50;
    s.addText(label, {
      x: x + 0.14, y: by, w: 1.0, h: 0.28,
      fontSize: 10, fontFace: FONT, color: GRAY_TXT, align: "left", margin: 0
    });
    s.addText(value, {
      x: x + 1.12, y: by - 0.01, w: w - 1.35, h: 0.32,
      fontSize: 12, fontFace: FONT, bold: true, color: NAVY, align: "left", margin: 0
    });
    if (sub) {
      s.addText(sub, {
        x: x + 1.12, y: by + 0.28, w: w - 1.35, h: 0.20,
        fontSize: 9, fontFace: FONT, color: GRAY_TXT, align: "left", margin: 0
      });
    }
    if (i < bodyLines.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: x + 0.14, y: by + 0.42, w: w - 0.28, h: 0,
        line: { color: "E5E7EB", width: 0.5 }
      });
    }
  });
}

// KPI 미니카드
function addKpiMini(s, x, y, w, label, value, delta, positive) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 1.10,
    fill: { color: CARD_BG },
    line: { color: "E5E7EB", width: 1 },
    rectRadius: 0.06
  });
  s.addText(label, {
    x: x + 0.14, y: y + 0.10, w: w - 0.28, h: 0.25,
    fontSize: 10, fontFace: FONT, color: GRAY_TXT, align: "left", margin: 0
  });
  s.addText(value, {
    x: x + 0.14, y: y + 0.35, w: w - 0.28, h: 0.42,
    fontSize: 22, fontFace: FONT, bold: true, color: NAVY, align: "left", margin: 0
  });
  const deltaColor = positive === null ? GRAY_TXT : positive ? "059669" : KT_RED;
  s.addText(delta, {
    x: x + 0.14, y: y + 0.78, w: w - 0.28, h: 0.22,
    fontSize: 9.5, fontFace: FONT, color: deltaColor, align: "left", margin: 0
  });
}

// Shape 기반 묶음 세로 막대 차트 (PPT 호환성 100%)
function addShapeBarChart(s, x, y, w, h, categories, seriesArr, title) {
  const TITLE_H  = 0.28;
  const LEGEND_H = 0.24;
  const LABEL_H  = 0.28;
  const chartH   = h - TITLE_H - LEGEND_H - LABEL_H;
  const chartY   = y + TITLE_H + LEGEND_H;
  const maxVal   = Math.max(...seriesArr.flatMap(s => s.values));

  // 배경 카드
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: CARD_BG }, line: { color: "E5E7EB", width: 1 }, rectRadius: 0.05
  });

  // 제목
  s.addText(title, {
    x: x + 0.08, y: y + 0.06, w: w - 0.16, h: TITLE_H,
    fontSize: 11, fontFace: FONT, bold: true, color: NAVY,
    align: "center", margin: 0
  });

  // 레전드
  const totalLegW = seriesArr.length * 1.05;
  const lStartX   = x + (w - totalLegW) / 2;
  seriesArr.forEach((ser, si) => {
    const lx = lStartX + si * 1.05;
    s.addShape(pres.shapes.RECTANGLE, {
      x: lx, y: y + TITLE_H + 0.05, w: 0.15, h: 0.11,
      fill: { color: ser.color }, line: { color: ser.color, width: 0 }
    });
    s.addText(ser.name, {
      x: lx + 0.18, y: y + TITLE_H + 0.04, w: 0.80, h: 0.16,
      fontSize: 9, fontFace: FONT, color: GRAY_TXT, align: "left", margin: 0
    });
  });

  // 0 기준선
  const baseY = chartY + chartH;
  s.addShape(pres.shapes.LINE, {
    x: x + 0.06, y: baseY, w: w - 0.12, h: 0,
    line: { color: "D1D5DB", width: 0.8 }
  });

  // 막대 그리기
  const numCats   = categories.length;
  const numSeries = seriesArr.length;
  const groupW    = (w - 0.12) / numCats;
  const BAR_W     = groupW * 0.3;
  const BAR_GAP   = groupW * 0.05;
  const groupPad  = (groupW - (BAR_W * numSeries + BAR_GAP * (numSeries - 1))) / 2;

  categories.forEach((cat, ci) => {
    const gx = x + 0.06 + ci * groupW;

    seriesArr.forEach((ser, si) => {
      const val  = ser.values[ci];
      const barH = Math.max((val / maxVal) * chartH * 0.90, 0.04);
      const bx   = gx + groupPad + si * (BAR_W + BAR_GAP);
      const by   = baseY - barH;

      // 막대
      s.addShape(pres.shapes.RECTANGLE, {
        x: bx, y: by, w: BAR_W, h: barH,
        fill: { color: ser.color }, line: { color: ser.color, width: 0 }
      });

      // 수치 레이블
      const label = val >= 1000 ? val.toLocaleString() : String(val);
      s.addText(label, {
        x: bx - 0.06, y: by - 0.20, w: BAR_W + 0.12, h: 0.19,
        fontSize: 8, fontFace: FONT, bold: si === 1,
        color: si === 1 ? KT_RED : "5C6370",
        align: "center", margin: 0
      });
    });

    // X축 카테고리 레이블
    s.addText(cat, {
      x: gx, y: baseY + 0.05, w: groupW, h: LABEL_H - 0.05,
      fontSize: 9, fontFace: FONT, color: GRAY_TXT,
      align: "center", margin: 0
    });
  });
}

// ══════════════════════════════════════════════════════
// SLIDE 1 — 표지 (다크)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_DARK };

  // kt m&s 로고 (우상단)
  s.addText("kt m&s", {
    x: 8.3, y: 0.22, w: 1.5, h: 0.35,
    fontSize: 16, fontFace: FONT, bold: true, color: KT_RED,
    align: "right", margin: 0
  });

  // 레드 대시 (카테고리 위)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.50, y: 2.42, w: 0.50, h: 0.06,
    fill: { color: KT_RED }, line: { color: KT_RED, width: 0 }
  });

  // 카테고리
  s.addText("채널마케팅팀  ·  광고 성과 보고", {
    x: 0.50, y: 2.52, w: 5, h: 0.32,
    fontSize: 13, fontFace: FONT, color: GRAY_TXT,
    align: "left", margin: 0
  });

  // 메인 타이틀
  s.addText("광고 성과 보고서", {
    x: 0.50, y: 2.88, w: 8.5, h: 0.78,
    fontSize: 52, fontFace: FONT, bold: true, color: WHITE,
    align: "left", margin: 0
  });

  // 날짜
  s.addText("2026.06.01 – 06.30  ·  Monthly Review", {
    x: 0.50, y: 3.72, w: 6, h: 0.32,
    fontSize: 14, fontFace: FONT, color: GRAY_TXT,
    align: "left", margin: 0
  });

  // 하단 슬로건
  s.addText("신나게！  당당하게！  무한도전！", {
    x: 0.50, y: 5.10, w: 5, h: 0.32,
    fontSize: 13, fontFace: FONT, bold: true, color: WHITE,
    align: "left", margin: 0
  });
}

// ══════════════════════════════════════════════════════
// SLIDE 2 — 목차 (라이트)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };

  // 헤더
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: 0.28, w: 1.0, h: 0.30,
    fill: { color: CARD_BG }, line: { color: "D1D5DB", width: 1 }, rectRadius: 0.04
  });
  s.addText("목차", {
    x: 0.45, y: 0.28, w: 1.0, h: 0.30,
    fontSize: 10, fontFace: FONT, color: GRAY_TXT, align: "center", valign: "middle", margin: 0
  });
  s.addText("Agenda", {
    x: 0.45, y: 0.65, w: 4, h: 0.62,
    fontSize: 36, fontFace: FONT, bold: true, color: NAVY, align: "left", margin: 0
  });

  // 아젠다 항목 (2 x 3)
  const items = [
    { num: "01", color: KT_RED,   title: "핵심 성과 요약",    desc: "노출·클릭·비용·CTR·CPC·전환율 KPI" },
    { num: "02", color: NAVY,     title: "매체별 성과 비교",  desc: "5월 vs 6월 전월 대비 · 매체별 성과표" },

    { num: "03", color: KT_RED,   title: "SA 전환 상세",      desc: "버튼 전환수·전환율·CPA 분석" },
    { num: "04", color: NAVY,     title: "Instagram 유기",    desc: "@ktplaza_story 도달·조회·TOP 게시물" },
    { num: "05", color: KT_RED,   title: "인사이트 & 액션",   desc: "이번 달 우선 실행 과제 TOP 3" },
    { num: "06", color: NAVY,     title: "다음 달 운영 방향", desc: "7월 광고 운영 및 예산 배분 방향" },
  ];

  const cols = [0.45, 5.10];
  const startY = 1.38;
  const cardH = 1.10;
  const cardW = 4.45;
  const gapY = 0.18;

  items.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = cols[col];
    const cy = startY + row * (cardH + gapY);

    // 카드
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cy, w: cardW, h: cardH,
      fill: { color: CARD_BG }, line: { color: "E5E7EB", width: 1 }, rectRadius: 0.06
    });
    // 번호 박스
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cy, w: 0.70, h: cardH,
      fill: { color: item.color }, rectRadius: 0.06
    });
    s.addText(item.num, {
      x: cx, y: cy + 0.35, w: 0.70, h: 0.40,
      fontSize: 16, fontFace: FONT, bold: true, color: WHITE,
      align: "center", margin: 0
    });
    // 제목
    s.addText(item.title, {
      x: cx + 0.85, y: cy + 0.18, w: cardW - 1.0, h: 0.35,
      fontSize: 14, fontFace: FONT, bold: true, color: NAVY,
      align: "left", margin: 0
    });
    // 설명
    s.addText(item.desc, {
      x: cx + 0.85, y: cy + 0.58, w: cardW - 1.0, h: 0.35,
      fontSize: 11, fontFace: FONT, color: GRAY_TXT,
      align: "left", margin: 0
    });
  });

  addFooter(s, 2);
}

// ══════════════════════════════════════════════════════
// SLIDE 3 — 섹션 01 구분 슬라이드
// ══════════════════════════════════════════════════════
makeSectionDivider("01", "핵심 성과 요약");

// ══════════════════════════════════════════════════════
// SLIDE 4 — 핵심 성과 요약 (라이트)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };
  addContentHeader(s, "핵심 성과 요약", "6월 광고 성과 KPI  ·  SA 2매체 기준", 4);

  // 6개 KPI 미니카드 (2행 3열) — SA 2매체 기준 / 전월 대비
  const kpis = [
    { label: "전체 노출",   value: "238,271",     delta: "5월 210,968  →  6월 238,271  ▲ +13.0%", pos: true  },
    { label: "전체 클릭",   value: "8,105",       delta: "5월 7,720  →  6월 8,105  ▲ +5.0%",      pos: true  },
    { label: "총 광고비",   value: "₩4,031,165",  delta: "5월 ₩4,156,547  →  6월 ₩4,031,165  ▼ -3.0%", pos: null },
    { label: "전체 CTR",    value: "3.40%",        delta: "5월 3.66%  →  6월 3.40%",              pos: false },
    { label: "평균 CPC",    value: "₩497",         delta: "5월 ₩538  →  6월 ₩497  ▼ -7.6%",      pos: true  },
    { label: "버튼 전환",   value: "1,862건",      delta: "5월 1,766건  →  6월 1,862건  ▲ +5.4%", pos: true  },
  ];

  const cw = 2.93, ch = 1.10, gap = 0.16, lm = 0.45;
  const r1y = 1.42, r2y = r1y + ch + gap;

  kpis.forEach((k, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const kx = lm + col * (cw + gap);
    const ky = row === 0 ? r1y : r2y;
    addKpiMini(s, kx, ky, cw, k.label, k.value, k.delta, k.pos);
  });

  // 하단 요약 노트 (핑크 배경)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: r2y + ch + 0.14, w: 9.1, h: 0.50,
    fill: { color: PINK_BG }, line: { color: "FCA5A5", width: 1 }, rectRadius: 0.06
  });
  s.addText(
    "네이버 SA · 구글 SA  2매체 기준  |  SA 전환율 22.97% 양호  ·  구글 SA CPC ₩225 고효율 확인  →  구글 SA 중심 예산 집중 검토 권장",
    {
      x: 0.60, y: r2y + ch + 0.14, w: 8.8, h: 0.50,
      fontSize: 10.5, fontFace: FONT, color: PINK_TXT,
      align: "left", valign: "middle", margin: 0
    }
  );
}

// ══════════════════════════════════════════════════════
// SLIDE 5 — 전월 대비 SA 채널 성과 비교 (신규)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };
  addContentHeader(s, "전월 대비 비교", "2026년 5월 vs 6월  ·  SA 채널 성과 비교", 5);

  const COL_5 = "B0BEC5";   // 5월 — 회색 (과거)
  const COL_6 = "E4002B";   // 6월 — KT 레드 (현재)
  const cats  = ["네이버 SA", "구글 SA"];

  const chartDefs = [
    { title: "노출수 비교",  data5: [125166, 85802],  data6: [132932, 105339], cx: 0.45 },
    { title: "클릭수 비교",  data5: [2164,   5556],   data6: [1891,   6214],   cx: 3.40 },
    { title: "전환수 비교",  data5: [634,    1132],   data6: [720,    1142],   cx: 6.35 },
  ];

  // Shape 기반 막대 차트 (PPT 호환성)
  chartDefs.forEach(cd => {
    addShapeBarChart(
      s, cd.cx, 1.38, 2.85, 3.20,
      cats,
      [
        { name: "5월", values: cd.data5, color: COL_5 },
        { name: "6월", values: cd.data6, color: COL_6 },
      ],
      cd.title
    );
  });

  // 하단 노트 (광고비 절감 & Meta 신규)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: 4.68, w: 9.1, h: 0.44,
    fill: { color: "EFF6FF" }, line: { color: "93C5FD", width: 1 }, rectRadius: 0.05
  });
  s.addText(
    "SA 광고비  5월 ₩4,156,547 → 6월 ₩4,031,165  (▼ -3.0% 절감)   |   SA 전환  5월 1,766건 → 6월 1,862건  ▲ +5.4%   |   CTR  5월 3.66% → 6월 3.40%",
    {
      x: 0.60, y: 4.68, w: 8.8, h: 0.44,
      fontSize: 10, fontFace: FONT, color: "1E40AF",
      align: "left", valign: "middle", margin: 0
    }
  );

  addFooter(s, 5);
}

// ══════════════════════════════════════════════════════
// SLIDE 6 — 섹션 02 구분
// ══════════════════════════════════════════════════════
makeSectionDivider("02", "매체별 성과 비교");

// ══════════════════════════════════════════════════════
// SLIDE 7 — 매체별 성과 비교 (라이트)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };
  addContentHeader(s, "매체별 성과 비교", "네이버 SA  ·  구글 SA  2매체  |  2026.06.01 – 06.30", 7);

  const NAVER_C  = "059669";   // 그린
  const GOOGLE_C = "2563EB";   // 블루
  const META_C   = "6366F1";   // 인디고

  // 헤더 행
  const cols  = ["매체", "광고비", "노출", "클릭", "CTR", "CPC", "버튼전환"];
  const colW  = [1.65, 1.60, 1.20, 1.00, 0.80, 0.95, 1.40];
  const tableY = 1.40;
  let xPos = 0.45;

  cols.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: tableY, w: colW[ci], h: 0.40,
      fill: { color: NAVY }, line: { color: WHITE, width: 0.5 }
    });
    s.addText(h, {
      x: xPos, y: tableY, w: colW[ci], h: 0.40,
      fontSize: 10.5, fontFace: FONT, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    xPos += colW[ci];
  });

  // 데이터 행 (SA 2매체 기준)
  const rows = [
    { cells: ["네이버 SA", "₩2,629,373\n(65.2%)", "132,932", "1,891", "1.42%", "₩1,390", "720건\n(38.08%)"], color: NAVER_C,  bg: "F0FDF4" },
    { cells: ["구글 SA",   "₩1,401,792\n(34.8%)", "105,339", "6,214", "5.90%", "₩225",   "1,142건\n(18.38%)"], color: GOOGLE_C, bg: "EFF6FF" },
  ];

  const rowH = 0.68;
  rows.forEach((row, ri) => {
    let rx = 0.45;
    const ry = tableY + 0.40 + ri * rowH;
    row.cells.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: rx, y: ry, w: colW[ci], h: rowH,
        fill: { color: ri % 2 === 0 ? CARD_BG : row.bg },
        line: { color: "E5E7EB", width: 0.5 }
      });
      if (ci === 0) {
        // 좌측 색상 바
        s.addShape(pres.shapes.RECTANGLE, {
          x: rx, y: ry, w: 0.06, h: rowH,
          fill: { color: row.color }, line: { color: row.color, width: 0 }
        });
      }
      s.addText(cell, {
        x: rx + (ci === 0 ? 0.12 : 0), y: ry,
        w: colW[ci] - (ci === 0 ? 0.12 : 0), h: rowH,
        fontSize: ci === 0 ? 11 : 10,
        fontFace: FONT, bold: ci === 0,
        color: ci === 0 ? NAVY : "374151",
        align: ci === 0 ? "left" : "center",
        valign: "middle", margin: ci === 0 ? 4 : 0
      });
      rx += colW[ci];
    });
  });

  // 합계 행
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: tableY + 0.40 + rows.length * rowH, w: colW.reduce((a, b) => a + b, 0), h: 0.45,
    fill: { color: "F3F4F6" }, line: { color: "D1D5DB", width: 0.5 }
  });
  const totalCells = ["합계", "₩4,031,165\n(100%)", "238,271", "8,105", "3.40%", "₩497", "1,862건\n(22.97%)"];
  let txPos = 0.45;
  totalCells.forEach((cell, ci) => {
    s.addText(cell, {
      x: txPos + (ci === 0 ? 0.12 : 0), y: tableY + 0.40 + rows.length * rowH,
      w: colW[ci] - (ci === 0 ? 0.12 : 0), h: 0.45,
      fontSize: ci === 0 ? 11 : 10, fontFace: FONT, bold: true,
      color: NAVY,
      align: ci === 0 ? "left" : "center",
      valign: "middle", margin: ci === 0 ? 4 : 0
    });
    txPos += colW[ci];
  });
}

// ══════════════════════════════════════════════════════
// SLIDE 8 — SA 전환 상세 (섹션 03)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };
  addContentHeader(s, "SA 전환 상세", "버튼 전환수  ·  전환율  ·  CPA  비교  |  2026.06", 8);

  // 차트: 전환수 비교 (Shape 기반)
  addShapeBarChart(
    s, 0.45, 1.40, 4.60, 3.35,
    ["네이버 SA", "구글 SA"],
    [{ name: "6월 전환수 (건)", values: [720, 1142], color: "059669" }],
    "매체별 전환수 비교 (건)"
  );

  // 우측 지표 카드 4개
  const metrics = [
    { label: "네이버 SA  전환율",  val: "38.08%",  sub: "클릭 1,891  →  전환 720건", color: "059669" },
    { label: "구글 SA  전환율",    val: "18.38%",  sub: "클릭 6,214  →  전환 1,142건", color: "2563EB" },
    { label: "CPA  (구글 SA)",     val: "₩1,228",  sub: "₩1,401,792  ÷  1,142건", color: "2563EB" },
    { label: "CPA  (네이버 SA)",   val: "₩3,652",  sub: "₩2,629,373  ÷  720건", color: "059669" },
  ];

  metrics.forEach((m, i) => {
    const my = 1.40 + i * 0.87;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.35, y: my, w: 4.2, h: 0.78,
      fill: { color: CARD_BG }, line: { color: "E5E7EB", width: 1 }, rectRadius: 0.06
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.35, y: my, w: 0.06, h: 0.78,
      fill: { color: m.color }, rectRadius: 0.06
    });
    s.addText(m.label, {
      x: 5.52, y: my + 0.07, w: 3.90, h: 0.25,
      fontSize: 10, fontFace: FONT, color: GRAY_TXT, align: "left", margin: 0
    });
    s.addText(m.val, {
      x: 5.52, y: my + 0.31, w: 3.90, h: 0.32,
      fontSize: 20, fontFace: FONT, bold: true, color: NAVY, align: "left", margin: 0
    });
    s.addText(m.sub, {
      x: 5.52, y: my + 0.58, w: 3.90, h: 0.20,
      fontSize: 9, fontFace: FONT, color: GRAY_TXT, align: "left", margin: 0
    });
  });
}

// ══════════════════════════════════════════════════════
// SLIDE 9 — 섹션 04 구분
// ══════════════════════════════════════════════════════
makeSectionDivider("04", "Instagram 유기 성과");

// ══════════════════════════════════════════════════════
// SLIDE 10 — Instagram 성과 (라이트)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };
  addContentHeader(s, "Instagram 유기", "@ktplaza_story  ·  2026년 6월", 10);

  // 상단 KPI 4개 (가로)
  const igKPIs = [
    { label: "팔로워 수",    val: "1,876",   delta: "현재 기준" },
    { label: "게시물 (6월)", val: "33개",    delta: "5월 29개  →  +4" },
    { label: "총 도달수",    val: "20,028",  delta: "5월 32,782  ▼" },
    { label: "총 조회수",    val: "51,608",  delta: "5월 46,573  ▲" },
  ];
  const kw = 2.20, ky = 1.42;
  igKPIs.forEach((k, i) => {
    const kx = 0.45 + i * (kw + 0.18);
    addKpiMini(s, kx, ky, kw, k.label, k.val, k.delta, null);
  });

  // TOP 3 게시물 테이블
  s.addText("6월 상위 게시물  TOP 3  (도달수 기준)", {
    x: 0.45, y: 2.62, w: 9.1, h: 0.28,
    fontSize: 12, fontFace: FONT, bold: true, color: NAVY, align: "left", margin: 0
  });

  const hdr = ["날짜", "유형", "내용", "도달수", "좋아요", "댓글", "저장"];
  const hw  = [0.80, 0.72, 4.25, 0.82, 0.72, 0.65, 0.60];
  let hx = 0.45;
  hdr.forEach((h, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: hx, y: 2.95, w: hw[i], h: 0.35,
      fill: { color: NAVY }, line: { color: WHITE, width: 0.5 }
    });
    s.addText(h, {
      x: hx, y: 2.95, w: hw[i], h: 0.35,
      fontSize: 10, fontFace: FONT, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0
    });
    hx += hw[i];
  });

  const posts = [
    ["06-09", "동영상", "KT플라자 직원이면 친절하기만 하면 돼", "2,835", "15", "0", "2"],
    ["06-17", "동영상", "핸드폰 바꾸고 탭/워치/이어폰 받는법",  "2,376", "22", "2", "6"],
    ["06-16", "이미지", "KT플라자 축구공 찾기 EVENT",            "2,207", "1,561", "3,327", "42"],
  ];

  posts.forEach((row, ri) => {
    let rx = 0.45;
    const ry = 3.30 + ri * 0.50;
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: rx, y: ry, w: hw[ci], h: 0.48,
        fill: { color: ri % 2 === 0 ? CARD_BG : "F9FAFB" },
        line: { color: "E5E7EB", width: 0.5 }
      });
      s.addText(cell, {
        x: rx, y: ry, w: hw[ci], h: 0.48,
        fontSize: ci === 2 ? 10 : 10.5, fontFace: FONT,
        color: NAVY,
        align: ci === 2 ? "left" : "center",
        valign: "middle", margin: ci === 2 ? 5 : 0
      });
      rx += hw[ci];
    });
  });

  // 인사이트 노트
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.45, y: 4.82, w: 9.1, h: 0.40,
    fill: { color: PINK_BG }, line: { color: "FCA5A5", width: 1 }, rectRadius: 0.05
  });
  s.addText(
    "이벤트 게시물(축구공 찾기)  —  좋아요 1,561 · 댓글 3,327  →  이벤트 콘텐츠의 참여도가 일반 콘텐츠 대비 압도적으로 높음",
    {
      x: 0.60, y: 4.82, w: 8.8, h: 0.40,
      fontSize: 10, fontFace: FONT, color: PINK_TXT,
      align: "left", valign: "middle", margin: 0
    }
  );
}

// ══════════════════════════════════════════════════════
// SLIDE 11 — 인사이트 & 액션 플랜 (라이트)
// ══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: BG_LIGHT };
  addContentHeader(s, "인사이트 & 액션", "이번 달 우선 실행 과제  TOP 3", 11);

  const actions = [
    {
      num: "01",
      color: KT_RED,
      title: "네이버 SA CTR 개선  —  광고 소재 즉시 교체",
      desc: "네이버 SA CTR 1.42%로 구글 SA 5.90% 대비 저조. 소재 노후화 또는 타겟 비매칭이 원인으로 추정됩니다. 네이버 SA 소재를 신규 크리에이티브로 교체하고 A/B 테스트를 실시하세요."
    },
    {
      num: "02",
      color: NAVY,
      title: "구글 SA 중심 예산 재배분  —  네이버 일부 이전",
      desc: "구글 SA  CPC ₩225 · CPA ₩1,228  vs  네이버 CPC ₩1,390 · CPA ₩3,652. 구글 SA가 전환 효율 압도적 우위. 네이버 예산 일부를 구글 SA로 이동하는 시뮬레이션을 권장합니다."
    },
    {
      num: "03",
      color: KT_RED,
      title: "리타겟팅 광고 도입  —  SA 전환 지속 확대",
      desc: "SA 전환율 22.97% (네이버 38.1% · 구글 18.4%) 양호하나 구글 SA CPA ₩1,228 개선 여지 존재. 웹사이트 방문자 대상 리타겟팅 캠페인 추가 시 구매 의향 사용자 재접촉 및 전환 확대 기대."
    },
  ];

  actions.forEach((a, i) => {
    const ay = 1.40 + i * 1.18;
    // 카드
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.45, y: ay, w: 9.1, h: 1.08,
      fill: { color: CARD_BG }, line: { color: "E5E7EB", width: 1 }, rectRadius: 0.06
    });
    // 번호 박스
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.45, y: ay, w: 0.72, h: 1.08,
      fill: { color: a.color }, rectRadius: 0.06
    });
    s.addText(a.num, {
      x: 0.45, y: ay + 0.33, w: 0.72, h: 0.42,
      fontSize: 20, fontFace: FONT, bold: true, color: WHITE,
      align: "center", margin: 0
    });
    // 제목
    s.addText(a.title, {
      x: 1.30, y: ay + 0.10, w: 8.1, h: 0.34,
      fontSize: 13, fontFace: FONT, bold: true, color: NAVY,
      align: "left", margin: 0
    });
    // 설명
    s.addText(a.desc, {
      x: 1.30, y: ay + 0.46, w: 8.1, h: 0.58,
      fontSize: 10.5, fontFace: FONT, color: GRAY_TXT,
      align: "left", margin: 0, wrap: true
    });
  });
}

// ══════════════════════════════════════════════════════
// WRITE
// ══════════════════════════════════════════════════════
pres.writeFile({ fileName: "KT플라자_광고성과_202606_ktmns.pptx" })
  .then(() => console.log("✅ 완료: KT플라자_광고성과_202606_ktmns.pptx"))
  .catch(err => { console.error("❌", err); process.exit(1); });
