#!/bin/bash
# ─────────────────────────────────────────────
#  광고 성과 자동화 — 서버 시작 스크립트 (macOS)
#  보고서를 만들 날에 이 파일을 더블클릭하세요.
#  창을 닫으면 서버가 종료됩니다.
# ─────────────────────────────────────────────
cd "$(dirname "$0")"

echo ""
echo "================================================"
echo "  광고 성과 자동화  |  서버 시작"
echo "================================================"
echo ""

# 현재 IP 주소 확인 (Wi-Fi → 유선 순서로 시도)
IP=$(ipconfig getifaddr en0 2>/dev/null)
if [ -z "$IP" ]; then
    IP=$(ipconfig getifaddr en1 2>/dev/null)
fi
if [ -z "$IP" ]; then
    IP=$(ipconfig getifaddr en2 2>/dev/null)
fi
if [ -z "$IP" ]; then
    IP="(IP 확인 실패 — 직접 확인하세요)"
fi

echo "  📡 팀원 접속 주소:"
echo ""
echo "       http://$IP:8501"
echo ""
echo "  ⚠️  이 창을 닫으면 서버가 종료됩니다."
echo "================================================"
echo ""

# Streamlit 서버 시작 (사내망 전체에 노출)
streamlit run app.py \
    --server.address 0.0.0.0 \
    --server.port 8501 \
    --server.headless true \
    --browser.gatherUsageStats false
