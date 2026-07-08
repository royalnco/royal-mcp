#!/bin/bash
# setup.sh — 가상환경 생성 및 패키지 설치 (macOS / Linux)
# 사용법: bash setup.sh
# 개발자(HTTP 모드 포함): bash setup.sh --dev

set -e

DEV=false
for arg in "$@"; do
  [ "$arg" = "--dev" ] && DEV=true
done

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo ""
echo "=== royal-mcp 환경 설정 ==="

# Python 설치 상태 확인
if ! command -v python3 &>/dev/null; then
  echo ""
  echo "❌ python3 를 찾을 수 없습니다."
  echo ""
  echo "   해결 방법:"
  echo "   1. https://www.python.org/ 에서 Python 3.11 이상을 설치하세요."
  echo "   2. 설치 완료 후 터미널을 새로 열고 bash setup.sh 를 다시 실행하세요."
  exit 1
fi

PYTHON_VER=$(python3 -c "import sys; v=sys.version_info; print(str(v.major)+'.'+str(v.minor))" 2>/dev/null) || true
if [ -z "$PYTHON_VER" ]; then
  echo ""
  echo "❌ Python 버전을 확인할 수 없습니다. python3 가 정상적으로 설치되었는지 확인하세요."
  echo ""
  echo "   해결 방법:"
  echo "   1. https://www.python.org/ 에서 Python 3.11 이상을 설치하세요."
  echo "   2. 설치 완료 후 터미널을 새로 열고 bash setup.sh 를 다시 실행하세요."
  exit 1
fi

PYTHON_MAJOR=$(echo "$PYTHON_VER" | cut -d. -f1)
PYTHON_MINOR=$(echo "$PYTHON_VER" | cut -d. -f2)
if [ "$PYTHON_MAJOR" -lt 3 ] || { [ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 11 ]; }; then
  echo ""
  echo "❌ Python $PYTHON_VER 이 설치되어 있으나 3.11 이상이 필요합니다."
  echo ""
  echo "   해결 방법:"
  echo "   1. https://www.python.org/ 에서 Python 3.11 이상을 설치하세요."
  echo "   2. 설치 완료 후 터미널을 새로 열고 bash setup.sh 를 다시 실행하세요."
  exit 1
fi

# 1. 가상환경 생성
if [ -d "$SCRIPT_DIR/.venv" ]; then
  echo "[1/3] .venv 이미 존재 — 재사용합니다."
else
  echo "[1/3] 가상환경 생성 중..."
  python3 -m venv "$SCRIPT_DIR/.venv"
  echo "      완료"
fi

PIP="$SCRIPT_DIR/.venv/bin/pip"
PYTHON="$SCRIPT_DIR/.venv/bin/python"

# 2. pip 업그레이드
echo "[2/3] pip 업그레이드 중..."
"$PYTHON" -m pip install --upgrade pip --quiet
echo "      완료"

# 3. 패키지 설치
if [ "$DEV" = true ]; then
  echo "[3/3] 패키지 설치 중 (개발자 모드)..."
  "$PIP" install -r "$SCRIPT_DIR/requirements-dev.txt"
else
  echo "[3/3] 패키지 설치 중..."
  "$PIP" install -r "$SCRIPT_DIR/requirements.txt"
fi
echo "      완료"

echo ""
echo "✅ 설치 완료!"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo " Claude Code MCP 등록 안내"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
echo "⚠  TENANT_ID · CLIENT_ID 는 정보관리팀에 문의하여 발급받으세요."
echo ""
echo " 환경변수(--env) 설명:"
echo "  MCP_SERVER_NAME         MCP 서버 표시 이름 (기본값 그대로 사용)"
echo "  TENANT_ID               Azure AD 테넌트 ID"
echo "  CLIENT_ID               Azure AD 앱 클라이언트 ID"
echo "  MS365_TOKEN_CACHE_PATH  토큰 캐시 파일 경로 (project scope 사용 시 권장, 생략 시 기본값)"
echo "  MS365_AUTO_AUTH         true: 서버 시작 시 자동 로그인 시도 / false: 수동 로그인(ms365 로그인해줘)"
echo ""
echo "[1단계] 기존 등록 제거  (처음 설치라면 건너뜁니다)"
echo "  claude mcp remove royal-mcp --scope project"
echo ""
echo "[2단계 — project] 프로젝트 폴더에서 실행  (프로젝트 전용 · 기본 권장)"
echo ""
echo '  claude mcp add --scope project royal-mcp \'
echo "    \"$PYTHON\" \"$SCRIPT_DIR/server.py\" \\"
echo '    --env "MCP_SERVER_NAME=royal-mcp" \'
echo '    --env "TENANT_ID=여기에_테넌트ID_입력" \'
echo '    --env "CLIENT_ID=여기에_클라이언트ID_입력" \'
echo '    --env "MS365_TOKEN_CACHE_PATH=$(pwd)/.royal-mcp/token_cache.bin" \'
echo '    --env "MS365_AUTO_AUTH=false"'
echo ""
echo "[2단계 — user]  어느 폴더에서나 실행 가능  (모든 프로젝트 공용)"
echo ""
echo '  claude mcp add --scope user royal-mcp \'
echo "    \"$PYTHON\" \"$SCRIPT_DIR/server.py\" \\"
echo '    --env "MCP_SERVER_NAME=royal-mcp" \'
echo '    --env "TENANT_ID=여기에_테넌트ID_입력" \'
echo '    --env "CLIENT_ID=여기에_클라이언트ID_입력" \'
echo '    --env "MS365_AUTO_AUTH=false"'
echo "  (MS365_TOKEN_CACHE_PATH 생략 시 기본값: ~/.royal-mcp/token_cache.bin)"
echo ""
