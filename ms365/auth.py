# ms365/auth.py — MS365 Graph API 인증 모듈
# ──────────────────────────────────────────────────────────────
# [HTTP 모드 — 다중 사용자 Azure AD OAuth]
#   MCP 2025-03-26 OAuth 스펙을 따릅니다.
#   클라이언트가 Azure AD로 직접 OAuth 인증 → MS365 액세스 토큰 취득.
#   서버는 요청마다 JWT를 검증하고 contextvar에 저장합니다.
#   ms365_login / ms365_logout 도구 호출 불필요.
#
# [STDIO 모드 — 단일 사용자 Device Code Flow]
#   Claude Desktop / Cursor 등 로컬 클라이언트용.
#   ms365_login 도구를 호출하면 Device Code가 발급됩니다.
#
# [Azure AD 앱 설정]
#   인증 → 퍼블릭 클라이언트 흐름 허용 = 예 (STDIO용)
#   인증 → 리디렉션 URI: Claude.ai OAuth 콜백 URL 추가 (HTTP용)
#   API 사용 권한 → Delegated 권한 목록 참고
# ──────────────────────────────────────────────────────────────

import os
import sys
import logging
import asyncio
import threading
import contextvars
from typing import Optional

log = logging.getLogger("royal-mcp.ms365-auth")

# ── Graph API에 필요한 Delegated 스코프 ─────────────────────
MS365_SCOPES = [
    "User.Read",
    "User.Read.All",            # 임직원 전체 조회
    "OrgContact.Read.All",      # 조직 연락처 조회
    "Mail.ReadWrite",           # 이메일 읽기 + 범주/플래그 등 수정 (Mail.Read superset)
    "Mail.Send",
    "Calendars.ReadWrite",
    "ChannelMessage.Send",
    # "ChannelMessage.Read.All" — 관리자 동의 필요, 로그인 오류 원인이 될 수 있어 제외
    # "ChannelMessage.ReadWrite" — update_teams_channel_message/delete_teams_channel_message 도구에 필요.
    #   관리자 동의 필요. Azure AD 앱에 권한 추가 + 관리자 동의를 먼저 완료한 뒤 이 줄의 주석을 해제할 것.
    #   동의 전에 주석 해제하면 전체 사용자의 로그인(Device Code Flow)이 실패할 수 있음.
    "Chat.ReadWrite",           # Teams 그룹·1:1 채팅 발송/읽기
    "Team.ReadBasic.All",       # 가입 팀 목록 조회
    "Sites.ReadWrite.All",
    "Files.ReadWrite.All",
    "Reports.Read.All",         # 사용량 보고서 조회 (org_usage_report). 로그인 계정이 Reports Reader 이상이어야 함
]

# ── STDIO 모드 전역 상태 (단일 사용자) ──────────────────────
_msal_app = None                # MSAL PublicClientApplication
_graph_client = None            # 인증된 GraphServiceClient
_user_email: Optional[str] = None
_user_display_name: Optional[str] = None
_authenticated = False
_login_in_progress = False      # 로그인 진행 중 여부
_pending_flow: dict | None = None  # Device Code Flow 대기 객체

# ── HTTP 모드: 요청별 컨텍스트 변수 (다중 사용자) ────────────
_http_token: contextvars.ContextVar[str] = contextvars.ContextVar("http_token", default="")
_http_email: contextvars.ContextVar[str] = contextvars.ContextVar("http_email", default="")
_http_name:  contextvars.ContextVar[str] = contextvars.ContextVar("http_name",  default="")

# ── PyJWT JWKS 클라이언트 (캐시됨) ──────────────────────────
_jwks_client = None

# ── 허용된 Graph 리소스 식별자 (완전 일치, 부분 문자열 매칭 금지) ──
# v1 리소스 URI(https://graph.microsoft.com)와 v2 앱 ID(GUID)를 모두 허용.
_GRAPH_AUDIENCES = {
    "https://graph.microsoft.com",
    "00000003-0000-0000-c000-000000000000",  # Microsoft Graph 앱 ID
}


def _get_jwks_client():
    """Azure AD JWKS 클라이언트를 초기화하거나 캐시된 인스턴스를 반환합니다."""
    global _jwks_client
    if _jwks_client is None:
        from jwt import PyJWKClient
        tenant_id = os.getenv("TENANT_ID", "common")
        _jwks_client = PyJWKClient(
            f"https://login.microsoftonline.com/{tenant_id}/discovery/v2.0/keys",
            cache_keys=True,
            max_cached_keys=16,
        )
    return _jwks_client


def validate_azure_token(token: str) -> dict:
    """Azure AD JWT를 검증하고 claims를 반환합니다.

    검증 항목: 서명(JWKS), 만료(exp), 테넌트(tid), 대상(aud, 허용 목록과 완전 일치).
    실패 시 예외를 발생시킵니다.
    """
    import jwt as pyjwt

    tenant_id = os.getenv("TENANT_ID")
    if not tenant_id:
        raise ValueError("TENANT_ID 환경변수가 설정되지 않았습니다.")

    jwks = _get_jwks_client()
    signing_key = jwks.get_signing_key_from_jwt(token)

    payload = pyjwt.decode(
        token,
        signing_key.key,
        algorithms=["RS256"],
        options={"verify_aud": False},  # aud는 수동 검증 (v1/v2 issuers 모두 허용)
    )

    # aud 검증: 허용 목록과 완전 일치해야 함 (부분 문자열 포함 매칭 금지 —
    # "graph.microsoft.com" in a 방식은 fake-graph.microsoft.com.evil.com 같은
    # 조작된 aud도 통과시키는 취약점이 있었음)
    aud = payload.get("aud", "")
    aud_values = aud if isinstance(aud, list) else [aud]
    normalized_aud = {str(a).rstrip("/") for a in aud_values}
    if not normalized_aud & _GRAPH_AUDIENCES:
        raise ValueError(f"허용되지 않은 토큰 대상(aud): {aud}")

    # 테넌트 검증
    if payload.get("tid") != tenant_id:
        raise ValueError(f"허용되지 않은 테넌트: {payload.get('tid')}")

    return payload


def set_http_request_context(token: str, claims: dict) -> None:
    """HTTP 요청별 인증 컨텍스트를 설정합니다. 미들웨어에서 호출됩니다."""
    _http_token.set(token)
    email = (claims.get("upn")
             or claims.get("unique_name")
             or claims.get("preferred_username")
             or "")
    name = claims.get("name") or ""
    _http_email.set(email)
    _http_name.set(name)


def get_request_token() -> str:
    """현재 HTTP 요청의 Azure AD 토큰을 반환합니다. STDIO 모드이면 빈 문자열."""
    return _http_token.get()


def _build_http_graph_client(access_token: str):
    """HTTP 모드용: Azure AD 토큰으로 GraphServiceClient를 생성합니다 (요청별 생성)."""
    from azure.core.credentials import AccessToken
    from msgraph import GraphServiceClient
    import time

    class _BearerCredential:
        def __init__(self, t: str):
            self._t = t
        def get_token(self, *scopes, **kwargs) -> AccessToken:
            return AccessToken(self._t, int(time.time()) + 3600)

    return GraphServiceClient(
        _BearerCredential(access_token),
        scopes=["https://graph.microsoft.com/.default"],
    )


def _build_token_cache():
    """플랫폼별 영구 토큰 캐시를 생성합니다. 실패 시 None 반환 (메모리 캐시 사용)."""
    cache_path = os.getenv("MS365_TOKEN_CACHE_PATH") or os.path.join(
        os.path.expanduser("~"), ".royal-mcp", "token_cache.bin"
    )
    try:
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        from msal_extensions import PersistedTokenCache, FilePersistence
        if sys.platform == "win32":
            from msal_extensions import FilePersistenceWithDataProtection
            persistence = FilePersistenceWithDataProtection(cache_path)
        elif sys.platform == "darwin":
            try:
                from msal_extensions import KeychainPersistence
                persistence = KeychainPersistence(cache_path, "royal-mcp", "token_cache")
            except Exception:
                persistence = FilePersistence(cache_path)
        else:
            try:
                from msal_extensions import LibsecretPersistence
                persistence = LibsecretPersistence(
                    cache_path, label="royal-mcp",
                    attributes={"application": "royal-mcp"},
                )
            except Exception:
                persistence = FilePersistence(cache_path)
        log.info("토큰 캐시 초기화: %s", cache_path)
        return PersistedTokenCache(persistence)
    except Exception as e:
        log.warning("영구 토큰 캐시 초기화 실패 (메모리 캐시 사용): %s", e)
        return None


def _ensure_msal_app():
    """MSAL PublicClientApplication을 초기화합니다."""
    global _msal_app
    if _msal_app:
        return True

    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    if not all([tenant_id, client_id]):
        log.warning("MS365 인증 불가 — TENANT_ID 또는 CLIENT_ID 환경변수 누락")
        return False

    try:
        import msal
        _msal_app = msal.PublicClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            token_cache=_build_token_cache(),
        )
        return True
    except ImportError:
        log.warning("MS365 인증 불가 — 'pip install msal' 를 실행하세요.")
        return False


def is_authenticated() -> bool:
    """현재 요청(HTTP) 또는 세션(STDIO) 인증 여부."""
    if _http_token.get():
        return True
    return _authenticated


def get_graph_client():
    """현재 컨텍스트의 GraphServiceClient를 반환합니다.
    HTTP 모드: 요청별 Azure AD 토큰으로 생성.
    STDIO 모드: 세션 전역 클라이언트.
    """
    token = _http_token.get()
    if token:
        return _build_http_graph_client(token)
    return _graph_client


def get_user_email() -> Optional[str]:
    """현재 사용자의 이메일을 반환합니다."""
    if _http_token.get():
        return _http_email.get() or None
    return _user_email


def get_user_display_name() -> Optional[str]:
    """현재 사용자의 표시 이름을 반환합니다."""
    if _http_token.get():
        return _http_name.get() or None
    return _user_display_name


def get_login_status() -> dict:
    """현재 로그인 상태를 반환합니다."""
    if _http_token.get():
        return {
            "status": "authenticated",
            "mode": "http_azure_ad",
            "user": _http_name.get() or None,
            "email": _http_email.get() or None,
        }
    if _authenticated:
        return {
            "status": "authenticated",
            "mode": "stdio_device_code",
            "user": _user_display_name,
            "email": _user_email,
        }
    if _login_in_progress:
        return {"status": "in_progress", "message": "로그인 대기 중입니다. 브라우저에서 인증을 완료해주세요."}
    return {"status": "not_authenticated"}


def initiate_device_code_login() -> dict:
    """Device Code Flow를 시작합니다 (STDIO 모드 전용).

    HTTP 모드에서는 Azure AD OAuth가 인증을 처리하므로 이 함수 호출이 불필요합니다.

    Returns:
        dict: verification_uri, user_code 등 로그인 안내 정보
    """
    global _pending_flow, _login_in_progress

    # HTTP 모드: Azure AD OAuth로 이미 인증됨
    if _http_token.get():
        return {
            "success": True,
            "already_authenticated": True,
            "mode": "http_azure_ad",
            "user": _http_name.get() or None,
            "email": _http_email.get() or None,
            "message": "HTTP 모드에서는 Azure AD OAuth로 자동 인증됩니다. 이미 로그인 상태입니다.",
        }

    if _authenticated:
        return {
            "success": True,
            "already_authenticated": True,
            "user": _user_display_name,
            "email": _user_email,
            "message": f"이미 {_user_display_name} ({_user_email}) 계정으로 로그인되어 있습니다.",
        }

    if _login_in_progress and _pending_flow:
        return {
            "success": True,
            "already_in_progress": True,
            "verification_uri": _pending_flow.get("verification_uri"),
            "user_code": _pending_flow.get("user_code"),
            "message": "로그인이 이미 진행 중입니다. 아래 정보로 브라우저에서 인증을 완료해주세요.",
        }

    if not _ensure_msal_app():
        return {"success": False, "message": "MS365 환경변수(TENANT_ID, CLIENT_ID)가 설정되지 않았습니다."}

    # Device Code Flow 시작
    scopes = [f"https://graph.microsoft.com/{s}" for s in MS365_SCOPES]
    flow = _msal_app.initiate_device_flow(scopes=scopes)

    if "user_code" not in flow:
        error = flow.get("error_description", flow.get("error", "알 수 없는 오류"))
        return {"success": False, "message": f"Device Code 발급 실패: {error}"}

    _pending_flow = flow
    _login_in_progress = True

    # 백그라운드에서 로그인 완료 대기
    thread = threading.Thread(target=_poll_device_code_completion, daemon=True)
    thread.start()

    log.info("Device Code Flow 시작 — 코드: %s", flow["user_code"])
    return {
        "success": True,
        "verification_uri": flow["verification_uri"],
        "user_code": flow["user_code"],
        "expires_in_seconds": flow.get("expires_in", 900),
        "message": (
            f"MS365 로그인이 필요합니다.\n"
            f"1. 브라우저에서 {flow['verification_uri']} 접속\n"
            f"2. 코드 입력: {flow['user_code']}\n"
            f"3. 본인 MS365 계정으로 로그인\n"
            f"로그인 완료 후 MS365 도구를 다시 사용하실 수 있습니다."
        ),
    }


def _poll_device_code_completion():
    """백그라운드 스레드에서 로그인 완료를 폴링합니다."""
    global _graph_client, _user_email, _user_display_name
    global _authenticated, _login_in_progress, _pending_flow

    if not _msal_app or not _pending_flow:
        _login_in_progress = False
        return

    # _pending_flow를 로컬 변수로 캡처 (finally에서 None 처리되기 전에 사용)
    flow = _pending_flow
    try:
        result = _msal_app.acquire_token_by_device_flow(flow)
        log.debug("acquire_token_by_device_flow 완료. 키: %s", list(result.keys()))

        if "access_token" in result:
            log.info("Device Code 로그인 성공 — 클라이언트 생성 시작")
            _build_graph_client(result["access_token"])  # 실패 시 예외 전파
            _authenticated = True
            log.info("MS365 인증 완료 — _authenticated = True")
        else:
            error = result.get("error_description", result.get("error", "알 수 없는 오류"))
            log.error("Device Code 로그인 실패: %s", error)
    except Exception as e:
        log.error("Device Code 폴링 오류: %s", e, exc_info=True)
    finally:
        _login_in_progress = False
        _pending_flow = None


def _build_graph_client(access_token: str):
    """액세스 토큰으로 GraphServiceClient를 생성합니다.

    실패 시는 예외를 전파하여 _poll_device_code_completion에서 _authenticated가
    잘못 설정되는 것을 방지합니다.
    """
    global _graph_client

    from azure.core.credentials import AccessToken, TokenCredential  # 올바른 경로
    from msgraph import GraphServiceClient

    class _StaticTokenCredential:
        """MSAL에서 받은 토큰을 azure-core 호환 credential로 래핑."""
        def __init__(self, token: str):
            self._token = token

        def get_token(self, *scopes, **kwargs) -> AccessToken:
            import time
            # exp=0 이면 msgraph-sdk가 매 호출마다 refresh 시도하므로
            # 현재 시간 + 1시간을 만료시간으로 설정
            return AccessToken(self._token, int(time.time()) + 3600)

    # 초기 토큰으로 클라이언트 생성
    credential = _StaticTokenCredential(access_token)
    _graph_client = GraphServiceClient(
        credential,
        scopes=["https://graph.microsoft.com/.default"],
    )
    log.info("GraphServiceClient 생성 완료 (StaticToken)")

    # 이후 토큰 만료 시 MSAL 캐시에서 자동 갱신으로 전환 시도
    try:
        refresh_cred = _MsalCredential(_msal_app)
        refresh_cred.get_token("https://graph.microsoft.com/.default")  # 캐시 확인
        _graph_client = GraphServiceClient(
            refresh_cred,
            scopes=["https://graph.microsoft.com/.default"],
        )
        log.info("GraphServiceClient MSAL 자동갱신 credential로 전환 완료")
    except Exception:
        log.info("MSAL 캐시 전환 실패 — StaticToken 클라이언트 유지 (토큰 유효 시간: 1시간)")


class _MsalCredential:
    """MSAL PublicClientApplication을 azure-identity 호환 credential로 래핑.

    토큰 만료 시 MSAL 캐시에서 자동 갱신합니다.
    """
    def __init__(self, app):
        self._app = app

    def get_token(self, *scopes, **kwargs):
        from azure.core.credentials import AccessToken

        # 캐시에서 토큰 조회 (자동 갱신 포함)
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(
                list(scopes) if scopes else ["https://graph.microsoft.com/.default"],
                account=accounts[0],
            )
            if result and "access_token" in result:
                return AccessToken(result["access_token"], 0)

        raise Exception("MS365 토큰이 만료되었습니다. ms365_login 도구로 다시 로그인해주세요.")


async def fetch_user_profile():
    """로그인 사용자 프로필을 캐시합니다.
    HTTP 모드: JWT claims에서 이미 추출됨 → API 호출 없음.
    STDIO 모드: Graph API 호출.
    """
    if _http_token.get():
        return  # HTTP 모드: set_http_request_context()에서 이미 설정됨
    global _user_email, _user_display_name
    if _user_email or not _graph_client:
        return
    try:
        me = await _graph_client.me.get()
        _user_email = me.mail or me.user_principal_name
        _user_display_name = me.display_name
        log.info("MS365 사용자 확인: %s (%s)", _user_display_name, _user_email)
    except Exception as e:
        log.error("사용자 프로필 조회 실패: %s", e)


def restore_session_from_cache() -> bool:
    """캐시된 토큰으로 자동 인증을 시도합니다 (server.py 시작 시 호출).

    MS365_AUTO_AUTH=true 일 때 server.py에서 호출됩니다.

    Returns:
        True: 자동 인증 성공
        False: 캐시 없음 또는 갱신 실패 (기존 device code 흐름 필요)
    """
    global _authenticated

    if _authenticated:
        return True

    if not _ensure_msal_app():
        return False

    accounts = _msal_app.get_accounts()
    if not accounts:
        log.info("토큰 캐시: 저장된 계정 없음 — ms365_login 필요")
        return False

    scopes = [f"https://graph.microsoft.com/{s}" for s in MS365_SCOPES]
    try:
        result = _msal_app.acquire_token_silent(scopes, account=accounts[0])
    except Exception as e:
        log.warning("캐시 토큰 조회 오류: %s", e)
        return False

    if result and "access_token" in result:
        try:
            _build_graph_client(result["access_token"])
            _authenticated = True
            log.info("캐시 토큰 자동 인증 완료 — %s", accounts[0].get("username", "?"))
            return True
        except Exception as e:
            log.warning("캐시 토큰으로 GraphServiceClient 생성 실패: %s", e)
            return False

    error = (result.get("error_description", result.get("error", "?")) if result else "no result")
    log.info("캐시 토큰 만료 또는 갱신 실패 (%s) — ms365_login 필요", error)
    return False


def logout():
    """로그아웃합니다.
    HTTP 모드: 서버에서 할 수 있는 로그아웃이 없습니다. Azure AD에서 토큰을 취소하거나
               Claude.ai에서 MCP 연결을 해제하세요.
    STDIO 모드: 세션 인증 정보를 초기화합니다.
    """
    if _http_token.get():
        email = _http_email.get()
        return {
            "success": True,
            "mode": "http_azure_ad",
            "message": (
                f"{email} 계정 세션 정보를 확인했습니다. "
                "완전한 로그아웃을 위해 Claude.ai에서 MCP 서버 연결을 해제하거나 "
                "Azure Portal(portal.azure.com)에서 앱 액세스를 취소하세요."
            ),
        }

    global _graph_client, _user_email, _user_display_name
    global _authenticated, _login_in_progress, _pending_flow

    old_user = _user_email
    _graph_client = None
    _user_email = None
    _user_display_name = None
    _authenticated = False
    _login_in_progress = False
    _pending_flow = None

    # MSAL 캐시도 클리어
    if _msal_app:
        for account in _msal_app.get_accounts():
            _msal_app.remove_account(account)

    # 토큰 캐시 파일 삭제
    cache_path = os.getenv("MS365_TOKEN_CACHE_PATH") or os.path.join(
        os.path.expanduser("~"), ".royal-mcp", "token_cache.bin"
    )
    try:
        if os.path.exists(cache_path):
            os.remove(cache_path)
    except OSError as e:
        log.warning("토큰 캐시 파일 삭제 실패: %s", e)

    log.info("MS365 로그아웃: %s", old_user)
    return {"success": True, "message": f"{old_user} 계정에서 로그아웃되었습니다."}


def require_auth() -> dict | None:
    """인증 여부를 확인하고, 미인증 시 안내 메시지를 반환합니다.

    Returns:
        None: 인증됨 (계속 진행 가능)
        dict: 미인증 안내 메시지 (도구에서 이 값을 바로 return)
    """
    if _http_token.get():
        return None  # HTTP 모드: 미들웨어가 이미 검증 완료

    if _authenticated:
        return None  # STDIO 모드: 인증됨

    if _login_in_progress:
        return {
            "success": False,
            "auth_required": True,
            "message": "MS365 로그인이 진행 중입니다. 브라우저에서 인증을 완료해주세요.",
        }

    return {
        "success": False,
        "auth_required": True,
        "message": "MS365 로그인이 필요합니다. ms365_login 도구를 먼저 호출해주세요.",
    }


def get_access_token() -> str | None:
    """현재 요청(HTTP) 또는 세션(STDIO)의 MS365 액세스 토큰을 반환합니다.

    HTTP 모드: Azure AD 토큰을 Graph API에 직접 사용합니다.
    STDIO 모드: MSAL 캐시에서 조용히(silent) 토큰을 갱신합니다.
    """
    token = _http_token.get()
    if token:
        return token  # HTTP 모드: Azure AD 토큰 = Graph API 토큰 (aud 동일)

    if not _authenticated or not _msal_app:
        return None
    try:
        accounts = _msal_app.get_accounts()
        if not accounts:
            return None
        scopes = [f"https://graph.microsoft.com/{s}" for s in MS365_SCOPES]
        result = _msal_app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
        log.warning("get_access_token: 토큰 갱신 실패 %s", result.get("error_description") if result else "no result")
        return None
    except Exception as e:
        log.error("get_access_token 오류: %s", e)
        return None
