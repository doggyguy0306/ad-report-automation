#!/bin/bash
# ─────────────────────────────────────────────
#  광고 성과 자동화 — 최초 설치 스크립트 (macOS)
#  이 파일을 더블클릭해서 한 번만 실행하세요.
# ─────────────────────────────────────────────
cd "$(dirname "$0")"

echo ""
echo "================================================"
echo "  광고 성과 자동화  |  최초 설치"
echo "================================================"
echo ""

# Python3 확인
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3이 설치되어 있지 않습니다."
    echo "   https://www.python.org/downloads/ 에서 설치 후 다시 실행하세요."
    echo ""
    read -p "엔터를 누르면 창이 닫힙니다..."
    exit 1
fi

echo "✅ Python3 확인: $(python3 --version)"
echo ""
echo "📦 필요한 패키지를 설치합니다..."
echo ""

pip3 install -r requirements.txt

echo ""
echo "================================================"
echo "  ✅ 설치 완료!"
echo ""
echo "  이제 '서버_시작.command' 를 더블클릭하면"
echo "  서버가 시작됩니다."
echo "================================================"
echo ""
read -p "엔터를 누르면 창이 닫힙니다..."
