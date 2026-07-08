# server.py  ─  로얄앤컴퍼니 MCP 서버
# ─────────────────────────────────────────────────────────────────
# [STDIO 모드] Claude Desktop / Copilot / Cursor  (단일 사용자)
#   python server.py
#
# [HTTP 모드]  Claude.ai 웹 / 사내 전 직원 공용  (다중 사용자)
#   python server.py --http
#   python server.py --http --port 8080 --host 0.0.0.0
#
# [HTTP 모드 인증 흐름]
#   클라이언트 → /.well-known/oauth-protected-resource
#              → Azure AD OAuth (사용자 @iroyal.kr 계정 로그인)
#              → Bearer {Azure AD 토큰} 으로 /mcp 접근
#   서버는 요청마다 JWT를 검증하고 사용자를 식별합니다.
#   ms365_login 도구 호출 불필요.
# ─────────────────────────────────────────────────────────────────

import os
import sys
import json as _json
import logging
import argparse

from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

load_dotenv()

from ms365 import register_ms365_tools, register_ms365_read_tools
from ms365 import initiate_device_code_login, logout, get_login_status, restore_session_from_cache, fetch_user_profile
from erp import register_erp_tools

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stderr,
)
log = logging.getLogger("royal-mcp")

# ── 인자 파싱 ──────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="로얄앤컴퍼니 MCP 서버")
parser.add_argument("--http",  action="store_true", help="HTTP 모드로 실행 (기본: STDIO)")
parser.add_argument("--host",  default="0.0.0.0",   help="바인딩 IP (기본: 0.0.0.0)")
parser.add_argument("--port",  type=int, default=8080, help="포트 (기본: 8080)")
args, _ = parser.parse_known_args()

# ── 서버 초기화 ────────────────────────────────────────────────
mcp = FastMCP(
    os.getenv("MCP_SERVER_NAME", "royal-mcp"),
    host=args.host,
    port=args.port,
    streamable_http_path="/mcp",
)

# ── MS365 Graph API 도구 등록 ────────────────────────────────
register_ms365_tools(mcp)
register_ms365_read_tools(mcp)

# ── 기간계 DB ERP 도구 등록 ──────────────────────────────────
register_erp_tools(mcp)

@mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True})
async def ms365_login() -> dict:
    """MS365 로그인 상태를 확인하거나 로그인을 시작합니다.

    HTTP 모드: Azure AD OAuth로 자동 인증됩니다. 이 도구를 호출하면 현재 인증 상태를 반환합니다.
    STDIO 모드: Device Code와 인증 URL이 반환됩니다. 브라우저에서 본인 MS365 계정으로 로그인하세요.
    """
    return initiate_device_code_login()

@mcp.tool(annotations={"readOnlyHint": False, "idempotentHint": True})
async def ms365_logout() -> dict:
    """MS365 로그아웃합니다.

    HTTP 모드: Claude.ai에서 MCP 연결을 해제하거나 Azure Portal에서 앱 액세스를 취소하세요.
    STDIO 모드: 세션 인증 정보를 초기화합니다. 다른 계정으로 전환 시 재로그인이 필요합니다.
    """
    return logout()

@mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
async def ms365_status() -> dict:
    """현재 MS365 로그인 상태와 사용자 정보를 확인합니다."""
    await fetch_user_profile()
    return get_login_status()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HTTP 모드: Azure AD JWT 미들웨어 (순수 ASGI — contextvars 안전)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class _AzureADMiddleware:
    """MCP 2025-03-26 OAuth 스펙 준수 Azure AD JWT 검증 미들웨어.

    요청별로 Azure AD JWT를 검증하고 contextvars에 사용자 정보를 저장합니다.
    BaseHTTPMiddleware 대신 순수 ASGI 방식을 사용하여 contextvars가 도구까지
    올바르게 전파됩니다.
    """

    _BYPASS = frozenset([
        "/health", "/register",
        "/.well-known/oauth-protected-resource",
        "/.well-known/oauth-authorization-server",
        "/.well-known/openid-configuration",
    ])

    def __init__(self, app):
        self.app = app

    async def _send_401(self, send, message: str, resource_metadata_url: str):
        body = _json.dumps(
            {"error": "Unauthorized", "message": message},
            ensure_ascii=False,
        ).encode()
        await send({
            "type": "http.response.start",
            "status": 401,
            "headers": [
                [b"content-type", b"application/json; charset=utf-8"],
                [b"www-authenticate",
                 f'Bearer realm="royal-mcp", resource_metadata="{resource_metadata_url}"'.encode()],
            ],
        })
        await send({"type": "http.response.body", "body": body})

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        path   = scope.get("path", "")
        method = scope.get("method", "")

        if (method == "OPTIONS"
                or path in self._BYPASS
                or path.startswith("/.well-known/")):
            await self.app(scope, receive, send)
            return

        headers = dict(scope.get("headers", []))
        host    = headers.get(b"host", b"localhost").decode()
        scheme  = "https" if scope.get("scheme") == "https" else "http"
        resource_meta_url = f"{scheme}://{host}/.well-known/oauth-protected-resource"

        auth  = headers.get(b"authorization", b"").decode()
        token = auth[7:].strip() if auth.lower().startswith("bearer ") else ""

        if not token:
            await self._send_401(
                send,
                "Bearer 토큰이 없습니다. Azure AD(@iroyal.kr)로 로그인해주세요.",
                resource_meta_url,
            )
            return

        try:
            from ms365.auth import validate_azure_token, set_http_request_context
            claims = validate_azure_token(token)
            set_http_request_context(token, claims)
            log.debug("Azure AD 인증 성공: %s", claims.get("upn") or claims.get("preferred_username"))
        except Exception as e:
            log.warning("Azure AD 토큰 검증 실패: %s", e)
            await self._send_401(send, f"토큰 검증 실패: {e}", resource_meta_url)
            return

        await self.app(scope, receive, send)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HTTP 모드 추가 라우트 구성
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _get_configured_app():
    """HTTP 모드 실행 시 사용할 ASGI 앱을 구성합니다."""
    from starlette.routing import Route
    from starlette.responses import JSONResponse

    inner = mcp.streamable_http_app()
    server_name = os.getenv("MCP_SERVER_NAME", "royal-mcp")
    tenant_id   = os.getenv("TENANT_ID", "common")

    # ── Health check ───────────────────────────────────────────
    async def health(request):
        return JSONResponse({"status": "ok", "server": server_name})

    # ── MCP OAuth 발견 엔드포인트 ──────────────────────────────
    # MCP 2025-03-26 클라이언트가 401 수신 후 자동 탐색합니다.
    # authorization_servers 에 Azure AD를 등록하면 클라이언트가
    # 사용자에게 @iroyal.kr 계정으로 OAuth 로그인을 안내합니다.
    async def oauth_protected_resource(request):
        base = str(request.base_url).rstrip("/")
        return JSONResponse({
            "resource": f"{base}/mcp",
            "authorization_servers": [
                f"https://login.microsoftonline.com/{tenant_id}/v2.0"
            ],
            "bearer_methods_supported": ["header"],
            "scopes_supported": ["https://graph.microsoft.com/.default"],
        })

    async def oauth_register(request):
        return JSONResponse(
            {
                "error": "registration_not_supported",
                "error_description": (
                    "Dynamic client registration is not supported. "
                    "Authenticate via Azure AD OAuth using the authorization_server in "
                    "/.well-known/oauth-protected-resource."
                ),
            },
            status_code=400,
        )

    inner.routes.append(Route("/health", health))
    inner.routes.append(Route("/.well-known/oauth-protected-resource", oauth_protected_resource))
    inner.routes.append(Route("/.well-known/oauth-protected-resource/{path:path}", oauth_protected_resource))
    inner.routes.append(Route("/register", oauth_register, methods=["POST"]))

    # Azure AD JWT 미들웨어로 전체 앱 래핑 (순수 ASGI)
    return _AzureADMiddleware(inner)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서버 실행
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if __name__ == "__main__":
    log.info("MS365 도구 전용 MCP 서버 시작")
    if os.getenv("MS365_AUTO_AUTH", "true").lower() == "true":
        restore_session_from_cache()
    if args.http:
        import uvicorn
        app = _get_configured_app()
        log.info("HTTP 모드 시작 — http://%s:%d/mcp", args.host, args.port)
        print("=" * 60)
        print("  로얄앤컴퍼니 MCP 서버 [HTTP 모드 — Azure AD OAuth]")
        print(f"  URL    : http://{args.host}:{args.port}/mcp")
        print(f"  Health : http://{args.host}:{args.port}/health")
        print(f"  Auth   : Azure AD ({os.getenv('TENANT_ID','?')})")
        print("=" * 60)
        uvicorn.run(app, host=args.host, port=args.port)
    else:
        log.info("STDIO 모드 시작 — 로얄앤컴퍼니 MCP 서버 (Claude Desktop / Copilot / Cursor 연결용)")
        mcp.run()
