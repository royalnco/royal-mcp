# ms365/graph_read.py — MS365 Graph API 읽기 도구 모듈
# ──────────────────────────────────────────────────────────────
# Claude.ai MS365 MCP의 7개 읽기 도구를 Graph API로 구현
# (Delegated 권한 / DeviceCodeCredential — 사용자 본인 로그인)
#
# [Delegated 권한 필요 목록]
#   Mail.Read              — 이메일 읽기
#   Calendars.Read         — 캘린더/일정 읽기 + getSchedule
#   Sites.Read.All         — SharePoint 사이트 검색
#   Files.Read.All         — 파일 내용 읽기
#   Reports.Read.All       — 사용량 보고서 조회 (org_usage_report)
#                            로그인 계정이 Reports Reader 이상 관리자 역할이어야 함
#
# [chat_message_search]
#   Delegated 인증 전환으로 향후 구현 가능.
#   현재는 Graph Search API 제약으로 미구현 상태 유지.
# ──────────────────────────────────────────────────────────────

import os
import logging
from typing import Annotated, List, Literal
from urllib.parse import urlparse

from pydantic import Field

log = logging.getLogger("royal-mcp.ms365-read")

_MAX_HOSTED_CONTENT_BYTES = 3 * 1024 * 1024  # Teams 인라인 이미지(hostedContent) base64 반환 한도


def _extract_id_from_uri(uri: str) -> str:
    """Graph API URL 또는 raw ID 문자열에서 리소스 ID를 추출합니다."""
    parsed = urlparse(uri)
    if parsed.scheme in ("https", "http") and "graph.microsoft.com" in (parsed.netloc or ""):
        parts = [p for p in parsed.path.split("/") if p]
        if parts:
            return parts[-1]
    return uri


def _parse_drive_and_item(uri: str):
    """Graph API URL에서 (drive_id, item_id) 쌍을 반환합니다.
    형식: .../drives/{driveId}/items/{itemId}
    파싱 실패 시 (None, uri) 반환.
    """
    parsed = urlparse(uri)
    if parsed.scheme in ("https", "http"):
        parts = [p for p in parsed.path.split("/") if p]
        try:
            drive_idx = parts.index("drives")
            item_idx = parts.index("items", drive_idx)
            return parts[drive_idx + 1], parts[item_idx + 1]
        except (ValueError, IndexError):
            pass
    return None, uri


def _to_kst_date(utc_str: str) -> str:
    """UTC datetime 문자열을 KST 날짜(YYYY-MM-DD)로 변환합니다. 파싱 실패 시 빈 문자열 반환."""
    if not utc_str:
        return ""
    try:
        from datetime import datetime, timezone, timedelta
        s = utc_str.replace("Z", "+00:00")
        # 하위 호환: 소수점 이하 7자리 이상이면 6자리로 잘라냄 (Python 3.10 이하 대비)
        if "." in s:
            head, tail = s.split(".", 1)
            digits, i = "", 0
            while i < len(tail) and tail[i].isdigit():
                digits += tail[i]; i += 1
            s = f"{head}.{digits[:6]}{tail[i:]}"
        dt = datetime.fromisoformat(s)
        return dt.astimezone(timezone(timedelta(hours=9))).strftime("%Y-%m-%d")
    except Exception:
        return ""


def _parse_graph_dt(utc_str: str):
    """Graph API datetime 문자열을 timezone-aware datetime으로 변환합니다. 실패 시 None 반환."""
    if not utc_str:
        return None
    try:
        from datetime import datetime, timezone
        s = utc_str.replace("Z", "+00:00").replace("z", "+00:00")
        if "." in s:
            head, tail = s.split(".", 1)
            digits, i = "", 0
            while i < len(tail) and tail[i].isdigit():
                digits += tail[i]; i += 1
            s = f"{head}.{digits[:6]}{tail[i:]}"
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def _to_kst_str(utc_str: str) -> str:
    """Graph API datetime 문자열을 KST 'YYYY-MM-DD HH:MM'으로 변환합니다. 실패 시 원문 반환."""
    dt = _parse_graph_dt(utc_str)
    if not dt:
        return utc_str or ""
    from datetime import timezone, timedelta
    return dt.astimezone(timezone(timedelta(hours=9))).strftime("%Y-%m-%d %H:%M")


def _html_to_text(html: str) -> str:
    """Teams 메시지 HTML 본문을 읽기 쉬운 평문으로 변환합니다."""
    import re
    import html as _html
    if not html:
        return ""
    s = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</li>", "\n", html)
    s = re.sub(r"<[^>]+>", "", s)
    s = _html.unescape(s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _message_sender(msg: dict) -> tuple[str, str]:
    """chatMessage의 (발신자 이름, 발신자 사용자 ID)를 반환합니다."""
    sender_info = msg.get("from") or {}
    user = sender_info.get("user") or {}
    app = sender_info.get("application") or {}
    name = user.get("displayName") or app.get("displayName") or "알 수 없음"
    return name, user.get("id") or ""


def _encode_sharing_url(url: str) -> str:
    """SharePoint/OneDrive URL을 Graph Shares API 토큰으로 인코딩합니다."""
    import base64
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).rstrip(b"=").decode("utf-8")
    return f"u!{b64}"


def _file_type_from_meta(mime_type: str, filename: str) -> str:
    """MIME 타입과 파일명 확장자로 파일 유형을 반환합니다."""
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if mime_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" or ext == "xlsx":
        return "xlsx"
    if mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document" or ext == "docx":
        return "docx"
    if mime_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation" or ext == "pptx":
        return "pptx"
    if mime_type == "application/pdf" or ext == "pdf":
        return "pdf"
    return "text"


def _extract_file_content(content_bytes: bytes, mime_type: str, filename: str, max_chars: int) -> dict:
    """파일 바이트를 메모리 내에서 텍스트로 추출합니다 (디스크 저장 없음)."""
    from io import BytesIO
    file_type = _file_type_from_meta(mime_type, filename)
    raw_text = ""
    if file_type == "xlsx":
        try:
            import openpyxl
            wb = openpyxl.load_workbook(BytesIO(content_bytes), data_only=True, read_only=True)
            parts = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                parts.append(f"=== Sheet: {sheet_name} ===")
                for row in ws.iter_rows(values_only=True):
                    row_vals = [str(c) if c is not None else "" for c in row]
                    if any(v.strip() for v in row_vals):
                        parts.append("\t".join(row_vals))
            raw_text = "\n".join(parts)
        except Exception as e:
            return {"content": f"[Excel 추출 실패: {e}]", "file_type": "xlsx", "truncated": False, "original_length": 0}
    elif file_type == "docx":
        try:
            from docx import Document
            doc = Document(BytesIO(content_bytes))
            raw_text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            return {"content": f"[Word 추출 실패: {e}]", "file_type": "docx", "truncated": False, "original_length": 0}
    elif file_type == "pptx":
        try:
            from pptx import Presentation
            prs = Presentation(BytesIO(content_bytes))
            parts = []
            for i, slide in enumerate(prs.slides, 1):
                parts.append(f"=== Slide {i} ===")
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            txt = para.text.strip()
                            if txt:
                                parts.append(txt)
            raw_text = "\n".join(parts)
        except Exception as e:
            return {"content": f"[PowerPoint 추출 실패: {e}]", "file_type": "pptx", "truncated": False, "original_length": 0}
    elif file_type == "pdf":
        try:
            from pypdf import PdfReader
            reader = PdfReader(BytesIO(content_bytes))
            parts = []
            for i, page in enumerate(reader.pages, 1):
                txt = page.extract_text() or ""
                if txt.strip():
                    parts.append(f"=== Page {i} ===\n{txt}")
            raw_text = "\n".join(parts)
        except Exception as e:
            return {"content": f"[PDF 추출 실패: {e}]", "file_type": "pdf", "truncated": False, "original_length": 0}
    else:
        try:
            raw_text = content_bytes.decode("utf-8")
        except UnicodeDecodeError:
            return {
                "content": f"[바이너리 파일: {len(content_bytes):,} bytes — 텍스트 변환 불가]",
                "file_type": "binary", "truncated": False, "original_length": len(content_bytes),
            }
    original_length = len(raw_text)
    truncated = original_length > max_chars
    content = raw_text[:max_chars] if truncated else raw_text
    return {"content": content, "file_type": file_type, "truncated": truncated, "original_length": original_length}


async def _resolve_file_meta(uri: str, access_token: str) -> dict:
    """Graph API URL 또는 웹 URL을 드라이브 메타데이터(drive_id, item_id, name, size, mime_type)로 해석합니다."""
    import httpx
    headers = {"Authorization": f"Bearer {access_token}"}
    drive_id, item_id = _parse_drive_and_item(uri)
    if drive_id:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.get(
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}",
                params={"$select": "id,name,size,file"},
                headers=headers,
            )
        if resp.status_code != 200:
            return {"error": f"파일 메타데이터 조회 실패: HTTP {resp.status_code}"}
        data = resp.json()
        return {
            "drive_id": drive_id,
            "item_id": data.get("id", item_id),
            "name": data.get("name", ""),
            "size": data.get("size") or 0,
            "mime_type": (data.get("file") or {}).get("mimeType", ""),
        }
    encoded = _encode_sharing_url(uri)
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(
            f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem",
            params={"$select": "id,name,size,file,parentReference"},
            headers=headers,
        )
    if resp.status_code != 200:
        return {"error": f"URL 해석 실패 (HTTP {resp.status_code}). 지원 형식: Graph API URL, SharePoint 파일 URL, OneDrive 공유 링크."}
    data = resp.json()
    parent = data.get("parentReference") or {}
    return {
        "drive_id": parent.get("driveId"),
        "item_id": data.get("id"),
        "name": data.get("name", ""),
        "size": data.get("size") or 0,
        "mime_type": (data.get("file") or {}).get("mimeType", ""),
    }


# ── Tool R17: 사용량 보고서(Reports API) 조회 지원 ───────────────

_USAGE_REPORT_ENDPOINTS = {
    "office365_active_users": "getOffice365ActiveUserDetail",   # 서비스별(Exchange/OneDrive/SharePoint/Teams) 최근 활동일
    "mailbox_usage":          "getMailboxUsageDetail",           # 메일박스 용량·건수
    "onedrive_usage":         "getOneDriveUsageAccountDetail",   # OneDrive 계정별 저장 용량
    "sharepoint_usage":       "getSharePointSiteUsageDetail",    # SharePoint 사이트별 저장 용량
    "teams_user_activity":    "getTeamsUserActivityUserDetail",  # 사용자별 Teams 채팅/통화/회의 활동
    "teams_device_usage":     "getTeamsDeviceUsageUserDetail",   # 사용자별 Teams 접속 기기 유형
    "email_activity":         "getEmailActivityUserDetail",      # 사용자별 이메일 송수신 활동
}


def _parse_usage_report_csv(text: str) -> list[dict]:
    """Graph Reports API가 반환하는 CSV 본문을 dict 목록으로 파싱합니다."""
    import csv
    from io import StringIO
    reader = csv.DictReader(StringIO(text))
    return [dict(row) for row in reader]


def register_ms365_read_tools(mcp):
    """MS365 Graph API 읽기 도구 7개를 MCP 서버에 등록합니다.

    도구는 항상 등록되지만, 각 도구 호출 시 로그인 여부를 확인합니다.
    미로그인 시 ms365_login 도구를 먼저 호출하도록 안내합니다.
    """
    from ms365.auth import get_graph_client, get_user_email, fetch_user_profile, require_auth, get_access_token

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R1: Teams 채팅 메시지 키워드 검색
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": False})
    async def chat_message_search(
        query: Annotated[str, Field(min_length=1, description="검색 키워드")],
        team_id: Annotated[str, Field(description="Teams 팀 ID (지정 시 해당 채널에서만 검색)")] = "",
        channel_id: Annotated[str, Field(description="Teams 채널 ID (team_id와 함께 지정)")] = "",
        limit: Annotated[int, Field(ge=1, le=50, description="최대 반환 건수 (기본 10)")] = 10,
    ) -> dict:
        """Teams 채널 및 채팅 메시지를 키워드로 검색합니다.
        - team_id + channel_id 지정: 해당 채널 메시지에서 키워드 필터링 (ChannelMessage.Read.All 필요)
        - 미지정: Graph Search API로 접근 가능한 전체 chatMessage 검색 (Chat.ReadWrite 필요)
        본문(body)은 잘라내지 않고 Graph API가 반환한 원문(HTML) 전체를 반환합니다.
        단, 전체 검색(Search API) 결과의 본문은 검색 엔진이 제공하는 값이라 원문 일부일 수 있으므로,
        원문 전체가 필요하면 반환된 team_id/channel_id로 teams_channel_messages 또는
        teams_channel_message_replies를 다시 조회하세요.
        ms365_login으로 로그인 후 사용하세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx
        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        # ── 특정 채널 지정 검색 ─────────────────────────────────
        if team_id and channel_id:
            try:
                async with httpx.AsyncClient(timeout=30) as client:
                    resp = await client.get(
                        f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages",
                        params={"$top": min(limit * 3, 50)},
                        headers=headers,
                    )
                    resp.raise_for_status()
                    data = resp.json()

                kw = query.lower()
                messages_out = []
                for msg in data.get("value", []):
                    body = (msg.get("body") or {}).get("content") or ""
                    if kw not in body.lower():
                        continue
                    sender = (msg.get("from") or {})
                    messages_out.append({
                        "id": msg.get("id"),
                        "sender": (sender.get("user") or {}).get("displayName"),
                        "created_at": msg.get("createdDateTime"),
                        "body": body,
                        "message_type": msg.get("messageType"),
                        "web_url": msg.get("webUrl"),
                    })
                    if len(messages_out) >= limit:
                        break

                return {
                    "success": True,
                    "query": query,
                    "team_id": team_id,
                    "channel_id": channel_id,
                    "count": len(messages_out),
                    "messages": messages_out,
                }
            except httpx.HTTPStatusError as e:
                log.error("chat_message_search(channel) HTTP 오류: %s", e.response.text)
                return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
            except Exception as e:
                log.error("chat_message_search(channel) 실패: %s", e)
                return {"success": False, "message": str(e)}

        # ── Graph Search API — chatMessage 전체 검색 ────────────
        try:
            search_body = {
                "requests": [{
                    "entityTypes": ["chatMessage"],
                    "query": {"queryString": query},
                    "from": 0,
                    "size": limit,
                }]
            }
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(
                    "https://graph.microsoft.com/v1.0/search/query",
                    json=search_body,
                    headers=headers,
                )
                resp.raise_for_status()
                data = resp.json()

            messages_out = []
            for response_item in data.get("value", []):
                for hits_container in response_item.get("hitsContainers", []):
                    for hit in hits_container.get("hits", []):
                        res = hit.get("resource", {})
                        messages_out.append({
                            "id": res.get("id"),
                            "sender": (res.get("from") or {}).get("user", {}).get("displayName"),
                            "created_at": res.get("createdDateTime"),
                            "body": (res.get("body") or {}).get("content", ""),
                            "channel_id": res.get("channelIdentity", {}).get("channelId"),
                            "team_id": res.get("channelIdentity", {}).get("teamId"),
                            "web_url": res.get("webUrl"),
                        })
                        if len(messages_out) >= limit:
                            break

            return {
                "success": True,
                "query": query,
                "search_scope": "전체 (Graph Search API)",
                "count": len(messages_out),
                "messages": messages_out,
            }
        except httpx.HTTPStatusError as e:
            log.error("chat_message_search(search) HTTP 오류: %s", e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("chat_message_search(search) 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R2: 회의 가능 시간 조회 (getSchedule)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def find_meeting_availability(
        attendees: Annotated[List[str], Field(description="스케줄 조회 대상 이메일 목록")],
        start: Annotated[str, Field(
            pattern=r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$",
            description="조회 시작 시간 ISO8601 (예: 2026-04-23T09:00:00, KST 기준)",
        )],
        end: Annotated[str, Field(
            pattern=r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$",
            description="조회 종료 시간 ISO8601 (예: 2026-04-23T18:00:00, KST 기준)",
        )],
        interval_minutes: Annotated[int, Field(ge=5, le=60, description="가용 시간 슬롯 간격(분). 기본 30")] = 30,
        calendar_owner: Annotated[str, Field(description="getSchedule 호출 주체 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """참석자들의 캘린더 가용 시간을 조회합니다 (Graph API getSchedule 사용).
        availability_view 필드의 각 문자: 0=Free, 1=Tentative, 2=Busy, 3=OOF, 4=WorkingElsewhere.
        Azure AD 앱에 Calendars.Read Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        owner = calendar_owner or get_user_email()
        try:
            from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
            from msgraph.generated.users.item.calendar.get_schedule.get_schedule_post_request_body import (
                GetSchedulePostRequestBody,
            )

            start_dt = DateTimeTimeZone()
            start_dt.date_time = start
            start_dt.time_zone = "Asia/Seoul"

            end_dt = DateTimeTimeZone()
            end_dt.date_time = end
            end_dt.time_zone = "Asia/Seoul"

            body = GetSchedulePostRequestBody()
            body.schedules = attendees
            body.start_time = start_dt
            body.end_time = end_dt
            body.availability_view_interval = interval_minutes

            if calendar_owner:
                result = await graph.users.by_user_id(owner).calendar.get_schedule.post(body)
            else:
                result = await graph.me.calendar.get_schedule.post(body)

            schedules_out = []
            for s in (result.value or []):
                items_out = []
                for item in (s.schedule_items or []):
                    items_out.append({
                        "status": item.status.value if item.status else None,
                        "start": item.start.date_time if item.start else None,
                        "end": item.end.date_time if item.end else None,
                        "subject": item.subject,
                    })
                schedules_out.append({
                    "email": s.schedule_id,
                    "availability_view": s.availability_view,
                    "schedule_items": items_out,
                    "error": s.error.message if s.error else None,
                })

            return {
                "success": True,
                "start": start,
                "end": end,
                "interval_minutes": interval_minutes,
                "schedules": schedules_out,
            }
        except Exception as e:
            log.error("find_meeting_availability 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R3: Outlook 캘린더 일정 검색
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def outlook_calendar_search(
        query: Annotated[str, Field(description="검색 키워드 (비우면 날짜 범위 전체 조회)")] = "",
        start: Annotated[str, Field(description="조회 시작 ISO8601 (예: 2026-04-23T00:00:00). 비우면 오늘 자정")] = "",
        end: Annotated[str, Field(description="조회 종료 ISO8601 (예: 2026-04-30T23:59:59). 비우면 start+7일")] = "",
        limit: Annotated[int, Field(ge=1, le=50, description="최대 반환 건수 (기본 10)")] = 10,
        calendar_owner: Annotated[str, Field(description="캘린더 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 캘린더 일정을 검색합니다.
        날짜 범위 지정 시 calendarView, 키워드만 지정 시 events?$search를 사용합니다.
        둘 다 지정 시 날짜 범위로 가져온 후 Python에서 제목 키워드 필터링합니다.
        Azure AD 앱에 Calendars.Read Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        owner = calendar_owner or get_user_email()
        try:
            from datetime import datetime, timedelta
            from zoneinfo import ZoneInfo

            KST = ZoneInfo("Asia/Seoul")
            now = datetime.now(KST)
            events_raw = []

            if start or (not query):
                # calendarView: 날짜 범위 기반 조회
                if not start:
                    start = now.strftime("%Y-%m-%dT00:00:00")
                if not end:
                    start_dt_obj = datetime.fromisoformat(start)
                    end = (start_dt_obj + timedelta(days=7)).strftime("%Y-%m-%dT23:59:59")

                from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import (
                    CalendarViewRequestBuilder,
                )

                # calendarView requires ISO 8601 with UTC offset; append KST offset if absent
                def _kst(dt_str: str) -> str:
                    return dt_str if ("+" in dt_str or "Z" in dt_str) else dt_str + "+09:00"

                qp = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
                    start_date_time=_kst(start),
                    end_date_time=_kst(end),
                    top=limit,
                    select=["id", "subject", "start", "end", "location",
                            "attendees", "bodyPreview", "isAllDay"],
                    orderby=["start/dateTime asc"],
                )
                config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                if calendar_owner:
                    result = await graph.users.by_user_id(owner).calendar_view.get(
                        request_configuration=config
                    )
                else:
                    result = await graph.me.calendar_view.get(
                        request_configuration=config
                    )
                events_raw = result.value or []

                # 키워드가 함께 지정된 경우 Python에서 제목 필터링
                if query:
                    events_raw = [ev for ev in events_raw if query.lower() in (ev.subject or "").lower()]

            else:
                # events?$search: 키워드 전용 검색 ($search와 $orderby 동시 사용 불가)
                from msgraph.generated.users.item.events.events_request_builder import (
                    EventsRequestBuilder,
                )

                qp = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
                    search=f'"{query}"',
                    top=limit,
                    select=["id", "subject", "start", "end", "location",
                            "bodyPreview", "isAllDay"],
                )
                config = EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                if calendar_owner:
                    result = await graph.users.by_user_id(owner).events.get(
                        request_configuration=config
                    )
                else:
                    result = await graph.me.events.get(
                        request_configuration=config
                    )
                events_raw = result.value or []

            events_out = []
            for ev in events_raw:
                attendee_list = []
                for att in (ev.attendees or []):
                    if att.email_address:
                        attendee_list.append(att.email_address.address)
                events_out.append({
                    "id": ev.id,
                    "subject": ev.subject,
                    "start": ev.start.date_time if ev.start else None,
                    "end": ev.end.date_time if ev.end else None,
                    "location": ev.location.display_name if ev.location else None,
                    "body_preview": ev.body_preview,
                    "is_all_day": ev.is_all_day,
                    "attendees": attendee_list,
                })

            return {
                "success": True,
                "owner": owner,
                "count": len(events_out),
                "events": events_out,
            }
        except Exception as e:
            log.error("outlook_calendar_search 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R4: Outlook 이메일 검색
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def outlook_email_search(
        query: Annotated[str, Field(min_length=1, description="KQL 검색 쿼리 (예: '회의' 또는 'subject:보고서 from:boss@royal.co.kr')")],
        folder: Annotated[str, Field(description="메일 폴더명 (inbox / sentitems / drafts). 비우면 전체 메일함")] = "",
        limit: Annotated[int, Field(ge=1, le=50, description="최대 반환 건수 (기본 10)")] = 10,
        mailbox_owner: Annotated[str, Field(description="메일박스 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 이메일을 KQL 키워드로 검색합니다.
        $search와 $orderby는 동시 사용이 불가능하므로 결과 순서는 Graph API 기본값을 따릅니다.
        Azure AD 앱에 Mail.Read Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        select_fields = ["id", "subject", "from", "receivedDateTime",
                         "bodyPreview", "isRead", "hasAttachments", "webLink"]
        try:
            if folder:
                from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
                    MessagesRequestBuilder as FolderMsgBuilder,
                )
                qp = FolderMsgBuilder.MessagesRequestBuilderGetQueryParameters(
                    search=f'"{query}"',
                    top=limit,
                    select=select_fields,
                )
                config = FolderMsgBuilder.MessagesRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                if mailbox_owner:
                    result = await (
                        graph.users.by_user_id(owner)
                        .mail_folders.by_mail_folder_id(folder)
                        .messages.get(request_configuration=config)
                    )
                else:
                    result = await (
                        graph.me
                        .mail_folders.by_mail_folder_id(folder)
                        .messages.get(request_configuration=config)
                    )
            else:
                from msgraph.generated.users.item.messages.messages_request_builder import (
                    MessagesRequestBuilder,
                )
                qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                    search=f'"{query}"',
                    top=limit,
                    select=select_fields,
                )
                config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                if mailbox_owner:
                    result = await graph.users.by_user_id(owner).messages.get(
                        request_configuration=config
                    )
                else:
                    result = await graph.me.messages.get(
                        request_configuration=config
                    )

            messages_out = []
            for msg in (result.value or []):
                messages_out.append({
                    "id": msg.id,
                    "subject": msg.subject,
                    "from": (
                        msg.from_.email_address.address
                        if msg.from_ and msg.from_.email_address else None
                    ),
                    "received_at": (
                        msg.received_date_time.isoformat()
                        if msg.received_date_time else None
                    ),
                    "body_preview": msg.body_preview,
                    "is_read": msg.is_read,
                    "has_attachments": msg.has_attachments,
                    "web_link": msg.web_link,
                })

            return {
                "success": True,
                "owner": owner,
                "folder": folder or "all",
                "count": len(messages_out),
                "messages": messages_out,
            }
        except Exception as e:
            log.error("outlook_email_search 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R5: 리소스 URI 읽기 (메일 본문 / 일정 상세 / 파일 내용)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def read_resource(
        uri: Annotated[str, Field(description=(
            "읽을 리소스 식별자. "
            "메일: outlook_email_search 반환 'id' 값. "
            "일정: outlook_calendar_search 반환 'id' 값. "
            "파일: Graph API URL (sharepoint_search 반환 'graph_url'), "
            "SharePoint 브라우저 URL, OneDrive 공유 링크, Teams 파일 링크 모두 지원."
        ))],
        resource_type: Annotated[str, Field(description="리소스 유형: mail / calendar_event / file. 비우면 자동 추론")] = "",
        owner: Annotated[str, Field(description="리소스 소유자 이메일 (mail/calendar_event 유형에 사용. 비우면 로그인 계정)")] = "",
        max_chars: Annotated[int, Field(ge=1000, le=200000, description="파일 유형에서 반환할 최대 문자 수 (기본 50000)")] = 50000,
    ) -> dict:
        """메일 본문·일정 상세·파일 내용을 읽어 반환합니다.
        파일은 xlsx/docx/pptx/pdf/텍스트를 메모리 내에서 추출합니다(디스크 저장 없음).
        SharePoint 브라우저 URL, OneDrive 공유 링크, Teams 파일 링크도 직접 입력 가능합니다.
        Azure AD 앱에 Mail.Read / Calendars.Read / Files.Read.All Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        resource_owner = owner or get_user_email()
        uri_lower = uri.lower()
        try:
            # 리소스 유형 자동 추론
            rtype = resource_type.lower()
            if not rtype:
                if "calendar" in uri_lower or "/events/" in uri_lower:
                    rtype = "calendar_event"
                elif "/drives/" in uri_lower or "/items/" in uri_lower:
                    rtype = "file"
                else:
                    rtype = "mail"

            if rtype == "mail":
                msg_id = _extract_id_from_uri(uri)
                from msgraph.generated.users.item.messages.item.message_item_request_builder import (
                    MessageItemRequestBuilder,
                )
                qp = MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
                    select=["id", "subject", "from", "toRecipients", "ccRecipients",
                            "receivedDateTime", "body", "isRead", "hasAttachments", "webLink"]
                )
                config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                if owner:
                    msg = await (
                        graph.users.by_user_id(resource_owner)
                        .messages.by_message_id(msg_id)
                        .get(request_configuration=config)
                    )
                else:
                    msg = await (
                        graph.me
                        .messages.by_message_id(msg_id)
                        .get(request_configuration=config)
                    )
                to_list = [
                    r.email_address.address
                    for r in (msg.to_recipients or [])
                    if r.email_address
                ]
                cc_list = [
                    r.email_address.address
                    for r in (msg.cc_recipients or [])
                    if r.email_address
                ]
                return {
                    "success": True,
                    "type": "mail",
                    "id": msg.id,
                    "subject": msg.subject,
                    "from": (
                        msg.from_.email_address.address
                        if msg.from_ and msg.from_.email_address else None
                    ),
                    "to": to_list,
                    "cc": cc_list,
                    "received_at": (
                        msg.received_date_time.isoformat()
                        if msg.received_date_time else None
                    ),
                    "body": msg.body.content if msg.body else None,
                    "body_type": (
                        msg.body.content_type.value
                        if msg.body and msg.body.content_type else None
                    ),
                    "is_read": msg.is_read,
                    "has_attachments": msg.has_attachments,
                    "web_link": msg.web_link,
                }

            elif rtype == "calendar_event":
                event_id = _extract_id_from_uri(uri)
                from msgraph.generated.users.item.events.item.event_item_request_builder import (
                    EventItemRequestBuilder,
                )
                qp = EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters(
                    select=["id", "subject", "start", "end", "location", "attendees",
                            "body", "isAllDay", "isOnlineMeeting", "onlineMeetingUrl"]
                )
                config = EventItemRequestBuilder.EventItemRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                if owner:
                    ev = await (
                        graph.users.by_user_id(resource_owner)
                        .events.by_event_id(event_id)
                        .get(request_configuration=config)
                    )
                else:
                    ev = await (
                        graph.me
                        .events.by_event_id(event_id)
                        .get(request_configuration=config)
                    )
                attendee_list = []
                for att in (ev.attendees or []):
                    if att.email_address:
                        attendee_list.append({
                            "email": att.email_address.address,
                            "name": att.email_address.name,
                            "status": (
                                att.status.response.value
                                if att.status and att.status.response else None
                            ),
                        })
                return {
                    "success": True,
                    "type": "calendar_event",
                    "id": ev.id,
                    "subject": ev.subject,
                    "start": ev.start.date_time if ev.start else None,
                    "end": ev.end.date_time if ev.end else None,
                    "location": ev.location.display_name if ev.location else None,
                    "body": ev.body.content if ev.body else None,
                    "body_type": (
                        ev.body.content_type.value
                        if ev.body and ev.body.content_type else None
                    ),
                    "is_all_day": ev.is_all_day,
                    "is_online_meeting": ev.is_online_meeting,
                    "online_meeting_url": ev.online_meeting_url,
                    "attendees": attendee_list,
                }

            elif rtype == "file":
                import httpx as _httpx
                token = get_access_token()
                if not token:
                    return {"success": False, "message": "액세스 토큰을 가져올 수 없습니다. ms365_login으로 다시 로그인해주세요."}

                meta = await _resolve_file_meta(uri, token)
                if "error" in meta:
                    return {"success": False, "message": meta["error"]}

                file_size = meta.get("size", 0)
                MAX_BYTES = 20 * 1024 * 1024  # 20MB
                if file_size > MAX_BYTES:
                    return {
                        "success": False,
                        "message": f"파일 크기({file_size / 1024 / 1024:.1f}MB)가 20MB 제한을 초과합니다.",
                        "filename": meta.get("name"),
                        "size": file_size,
                    }

                drive_id = meta.get("drive_id")
                item_id = meta.get("item_id")
                if not drive_id or not item_id:
                    return {"success": False, "message": "파일 드라이브 정보를 확인할 수 없습니다."}

                async with _httpx.AsyncClient(timeout=60) as client:
                    dl_resp = await client.get(
                        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
                        headers={"Authorization": f"Bearer {token}"},
                        follow_redirects=True,
                    )
                if dl_resp.status_code != 200:
                    return {"success": False, "message": f"파일 다운로드 실패: HTTP {dl_resp.status_code}"}

                content_bytes = dl_resp.content
                if not content_bytes:
                    return {"success": True, "type": "file", "uri": uri, "filename": meta.get("name"), "content": ""}

                extracted = _extract_file_content(
                    content_bytes,
                    meta.get("mime_type", ""),
                    meta.get("name", ""),
                    max_chars,
                )
                return {
                    "success": True,
                    "type": "file",
                    "uri": uri,
                    "filename": meta.get("name"),
                    "size": file_size,
                    "file_type": extracted["file_type"],
                    "content": extracted["content"],
                    "truncated": extracted["truncated"],
                    "original_length": extracted["original_length"],
                }

            else:
                return {"success": False, "message": f"지원하지 않는 리소스 유형: '{resource_type}'. mail / calendar_event / file 중 하나를 입력하세요."}

        except Exception as e:
            log.error("read_resource 실패 (uri=%s): %s", uri, e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R6: SharePoint / OneDrive 파일 검색
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def sharepoint_search(
        query: Annotated[str, Field(min_length=1, description="검색 키워드 (예: 'aa', '보고서')")],
        site_id: Annotated[str, Field(description="SharePoint 사이트 ID. 비우면 로그인 사용자의 OneDrive + SharePoint 전체에서 검색합니다.")] = "",
        limit: Annotated[int, Field(ge=1, le=50, description="최대 반환 건수 (기본 10)")] = 10,
    ) -> dict:
        """파일과 폴더를 키워드로 검색합니다.

        - site_id 없음: Graph 전역 Search API(/search/query) + OneDrive(/me/drive) 병렬 검색.
          로그인 사용자가 접근 가능한 SharePoint 및 OneDrive 전체를 탐색합니다.
        - site_id 있음: 해당 사이트의 드라이브에서만 검색합니다.

        반환된 'graph_url'을 read_resource 도구에 전달하면 파일 내용을 읽을 수 있습니다.
        Azure AD 앱에 Files.Read.All, Sites.Read.All Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()

        # ── site_id 있음: 기존 사이트별 드라이브 검색 ─────────────
        if site_id:
            try:
                drive = await graph.sites.by_site_id(site_id).drive.get()
                drive_id = drive.id
                from msgraph.generated.drives.item.search_with_q.search_with_q_request_builder import (
                    SearchWithQRequestBuilder,
                )
                qp = SearchWithQRequestBuilder.SearchWithQRequestBuilderGetQueryParameters(
                    top=limit,
                    select=["id", "name", "size", "webUrl", "lastModifiedDateTime",
                            "folder", "file", "parentReference"],
                )
                cfg = SearchWithQRequestBuilder.SearchWithQRequestBuilderGetRequestConfiguration(
                    query_parameters=qp
                )
                result = await graph.drives.by_drive_id(drive_id).search_with_q(query).get(
                    request_configuration=cfg
                )
                items_out = []
                for item in (result.value or []):
                    items_out.append({
                        "id": item.id,
                        "name": item.name,
                        "size": item.size,
                        "web_url": item.web_url,
                        "graph_url": f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item.id}" if item.id else None,
                        "last_modified": item.last_modified_date_time.isoformat() if item.last_modified_date_time else None,
                        "is_folder": item.folder is not None,
                        "mime_type": item.file.mime_type if item.file else None,
                        "source": "site",
                        "site_id": site_id,
                    })
                    if len(items_out) >= limit:
                        break
                return {"success": True, "query": query, "site_id": site_id, "count": len(items_out), "items": items_out}
            except Exception as e:
                log.error("sharepoint_search(site) 실패: %s", e)
                return {"success": False, "message": str(e)}

        # ── site_id 없음: 전역 Search API + OneDrive httpx 직접 호출 ───
        import httpx

        items_out = []
        errors = []

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        # 1) Graph 전역 Search API — driveItem 엔티티
        try:
            search_body = {
                "requests": [{
                    "entityTypes": ["driveItem"],
                    "query": {"queryString": query},
                    "fields": ["id", "name", "size", "webUrl", "lastModifiedDateTime",
                               "parentReference", "file", "folder"],
                    "from": 0,
                    "size": limit,
                }]
            }
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(
                    "https://graph.microsoft.com/v1.0/search/query",
                    json=search_body,
                    headers=headers,
                )
                resp.raise_for_status()
                data = resp.json()

            for response_item in data.get("value", []):
                for hit in response_item.get("hitsContainers", []):
                    for h in hit.get("hits", []):
                        res = h.get("resource", {})
                        parent = res.get("parentReference", {})
                        drive_id = parent.get("driveId")
                        item_id = res.get("id")
                        items_out.append({
                            "id": item_id,
                            "name": res.get("name"),
                            "size": res.get("size"),
                            "web_url": res.get("webUrl"),
                            "graph_url": f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}" if drive_id and item_id else None,
                            "last_modified": res.get("lastModifiedDateTime"),
                            "is_folder": "folder" in res,
                            "site_id": parent.get("siteId"),
                            "drive_id": drive_id,
                            "source": "global_search",
                        })
                        if len(items_out) >= limit:
                            break
            log.info("전역 Search API: %d건 반환", len(items_out))
        except Exception as e:
            log.warning("전역 Search API 실패: %s", e)
            errors.append(f"전역검색: {e}")

        # 2) OneDrive (/me/drive/root/search) — httpx 직접 호출
        if len(items_out) < limit:
            try:
                remain = limit - len(items_out)
                select = "id,name,size,webUrl,lastModifiedDateTime,folder,file,parentReference"
                async with httpx.AsyncClient(timeout=30) as client:
                    resp = await client.get(
                        f"https://graph.microsoft.com/v1.0/me/drive/root/search(q='{query}')",
                        params={"$top": remain, "$select": select},
                        headers=headers,
                    )
                    resp.raise_for_status()
                    me_data = resp.json()

                seen_ids = {i["id"] for i in items_out if i.get("id")}
                for item in me_data.get("value", []):
                    if item.get("id") in seen_ids:
                        continue
                    parent = item.get("parentReference", {})
                    drive_id = parent.get("driveId")
                    item_id = item.get("id")
                    items_out.append({
                        "id": item_id,
                        "name": item.get("name"),
                        "size": item.get("size"),
                        "web_url": item.get("webUrl"),
                        "graph_url": f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}" if drive_id and item_id else None,
                        "last_modified": item.get("lastModifiedDateTime"),
                        "is_folder": "folder" in item,
                        "mime_type": item.get("file", {}).get("mimeType"),
                        "site_id": parent.get("siteId"),
                        "drive_id": drive_id,
                        "source": "onedrive",
                    })
                    if len(items_out) >= limit:
                        break
                log.info("OneDrive 검색: %d건 추가", len([i for i in items_out if i.get("source") == "onedrive"]))
            except Exception as e:
                log.warning("OneDrive 검색 실패: %s", e)
                errors.append(f"OneDrive: {e}")

        return {
            "success": True,
            "query": query,
            "search_scope": "global (SharePoint + OneDrive)" if not site_id else f"site:{site_id}",
            "count": len(items_out),
            "items": items_out,
            **(({"warnings": errors}) if errors else {}),
        }

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R7: SharePoint 폴더 검색
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def sharepoint_folder_search(
        query: Annotated[str, Field(min_length=1, description="검색 키워드")],
        site_id: Annotated[str, Field(description="SharePoint 사이트 ID. 비우면 OneDrive + SharePoint 전체에서 폴더를 검색합니다.")] = "",
        limit: Annotated[int, Field(ge=1, le=50, description="최대 반환 건수 (기본 10)")] = 10,
    ) -> dict:
        """폴더만 키워드로 검색합니다. sharepoint_search와 동일 로직이나 폴더 항목만 반환합니다.
        site_id 없으면 전역 검색, 있으면 해당 사이트에서만 검색합니다.
        Azure AD 앱에 Files.Read.All, Sites.Read.All Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        # sharepoint_search(폴더 포함 전체) 결과에서 폴더만 필터링
        result = await sharepoint_search(query=query, site_id=site_id, limit=limit * 5)
        if not result.get("success"):
            return result

        folders = [i for i in result.get("items", []) if i.get("is_folder")][:limit]
        return {
            "success": True,
            "query": query,
            "search_scope": result.get("search_scope"),
            "count": len(folders),
            "folders": folders,
        }

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R8: 임직원 / 조직 연락처 조회 (활성 사용자만)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def org_users_search(
        query: Annotated[str, Field(description="검색어 (이름·부서·이메일). 비우면 활성 사용자 전체 목록을 반환합니다.")] = "",
        include_contacts: Annotated[bool, Field(description="조직 연락처(/contacts)도 함께 조회할지 여부 (기본 False)")] = False,
        limit: Annotated[int, Field(ge=1, le=200, description="최대 반환 건수 (기본 50)")] = 50,
    ) -> dict:
        """활성(Active) 임직원 목록을 조회하거나 이름·부서·이메일로 검색합니다.

        - 비활성 계정(퇴직자 등)은 자동으로 제외됩니다 (accountEnabled = true).
        - query 없음: 전체 활성 사용자 목록 반환 (limit 건)
        - query 있음: displayName, department, mail 포함 여부로 필터링
        - include_contacts=True: 조직 연락처(OrgContact)도 함께 반환
        - 특정 인원 상세 조회는 org_user_detail 도구를 사용하세요.

        Azure AD 앱에 다음 Delegated 권한이 필요합니다:
          - User.Read.All      (임직원 조회, 관리자 동의 필요)
          - OrgContact.Read.All (조직 연락처 조회, include_contacts=True 시)
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "ConsistencyLevel": "eventual",
        }

        select = "displayName,department,jobTitle,businessPhones,mobilePhone,mail,accountEnabled"
        users_out: list = []
        errors: list = []

        # ── 1) /v1.0/users 조회 (활성 사용자만) ──────────────────
        try:
            if query:
                # $search + $filter 동시 사용 불가 → over-fetch 후 코드 필터링
                uparams: dict = {
                    "$select": select,
                    "$top": min(limit * 3, 200),
                    "$search": f'"displayName:{query}" OR "mail:{query}" OR "department:{query}"',
                    "$count": "true",
                }
            else:
                uparams = {
                    "$select": select,
                    "$top": limit,
                    "$filter": "accountEnabled eq true",
                    "$orderby": "displayName",
                    "$count": "true",
                }

            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(
                    "https://graph.microsoft.com/v1.0/users",
                    params=uparams,
                    headers=headers,
                )
                resp.raise_for_status()
                data = resp.json()

            for u in data.get("value", []):
                if not u.get("accountEnabled", True):
                    continue
                users_out.append({
                    "type": "user",
                    "display_name": u.get("displayName"),
                    "job_title": u.get("jobTitle"),
                    "department": u.get("department"),
                    "mail": u.get("mail"),
                    "business_phones": u.get("businessPhones", []),
                    "mobile_phone": u.get("mobilePhone"),
                })
                if len(users_out) >= limit:
                    break
            log.info("임직원 조회 (활성): %d건 반환", len(users_out))
        except Exception as e:
            log.warning("임직원 조회 실패: %s", e)
            errors.append(f"users: {e}")

        # ── 2) /v1.0/contacts — 선택 옵션 ─────────────────────
        contacts_out: list = []
        if include_contacts:
            try:
                cparams: dict = {
                    "$select": "displayName,department,jobTitle,businessPhones,mobilePhone,emailAddresses",
                    "$top": limit,
                }
                if query:
                    cparams["$search"] = f'"displayName:{query}"'

                async with httpx.AsyncClient(timeout=30) as client:
                    cresp = await client.get(
                        "https://graph.microsoft.com/v1.0/contacts",
                        params=cparams,
                        headers=headers,
                    )
                    cresp.raise_for_status()
                    cdata = cresp.json()

                for c in cdata.get("value", []):
                    emails = [e.get("address") for e in c.get("emailAddresses", []) if e.get("address")]
                    contacts_out.append({
                        "type": "org_contact",
                        "display_name": c.get("displayName"),
                        "job_title": c.get("jobTitle"),
                        "department": c.get("department"),
                        "mail": emails[0] if emails else None,
                        "business_phones": c.get("businessPhones", []),
                        "mobile_phone": c.get("mobilePhone"),
                    })
                log.info("조직 연락처 조회: %d건 반환", len(contacts_out))
            except Exception as e:
                log.warning("조직 연락처 조회 실패: %s", e)
                errors.append(f"contacts: {e}")

        return {
            "success": True,
            "query": query or "(전체 활성 사용자)",
            "user_count": len(users_out),
            "contact_count": len(contacts_out),
            "users": users_out,
            **({"contacts": contacts_out} if include_contacts else {}),
            **({"warnings": errors} if errors else {}),
        }

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R9: 특정 임직원 상세 정보 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})

    async def org_user_detail(
        identifier: Annotated[str, Field(min_length=1, description="조회할 사용자 식별자. 이메일(mail), UPN(userPrincipalName), 이름(displayName) 모두 허용됩니다. 예: 'kjr@iroyal.kr', '김정렴'")],
    ) -> dict:
        """특정 임직원 한 명의 상세 정보를 조회합니다.

        - 이메일/UPN 입력 시: /v1.0/users/{id} 직접 조회 (정확·빠름)
        - 이름 입력 시: displayName으로 검색 후 첫 번째 활성 사용자 반환

        반환 정보: 이름, 직함, 부서, 이메일, 사내전화, 휴대전화

        Azure AD 앱에 User.Read.All Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "ConsistencyLevel": "eventual",
        }

        select = "displayName,department,jobTitle,businessPhones,mobilePhone,mail,accountEnabled"

        user_data = None

        # ── 이메일/UPN이면 직접 조회, 이름이면 $search ─────────
        is_email_or_upn = "@" in identifier
        async with httpx.AsyncClient(timeout=30) as client:
            if is_email_or_upn:
                resp = await client.get(
                    f"https://graph.microsoft.com/v1.0/users/{identifier}",
                    params={"$select": select},
                    headers=headers,
                )
                if resp.status_code == 200:
                    user_data = resp.json()
                else:
                    return {
                        "success": False,
                        "message": f"사용자를 찾을 수 없습니다: {identifier} (HTTP {resp.status_code})",
                    }
            else:
                resp = await client.get(
                    "https://graph.microsoft.com/v1.0/users",
                    params={
                        "$select": select,
                        "$search": f'"displayName:{identifier}"',
                        "$count": "true",
                        "$top": 10,
                    },
                    headers=headers,
                )
                if resp.status_code != 200:
                    return {
                        "success": False,
                        "message": f"사용자 검색 실패: HTTP {resp.status_code} — {resp.text[:200]}",
                    }
                results = resp.json().get("value", [])
                active = [u for u in results if u.get("accountEnabled", True)]
                if not active:
                    return {
                        "success": False,
                        "message": f"'{identifier}' 이름과 일치하는 활성 사용자가 없습니다.",
                        "candidates": len(results),
                    }
                user_data = active[0]
                if len(active) > 1:
                    log.info("이름 '%s' 검색 결과 %d명 — 첫 번째 선택", identifier, len(active))

        if not user_data:
            return {"success": False, "message": "사용자 데이터를 가져오지 못했습니다."}

        return {
            "success": True,
            "identifier": identifier,
            "user": {
                "display_name": user_data.get("displayName"),
                "job_title": user_data.get("jobTitle"),
                "department": user_data.get("department"),
                "mail": user_data.get("mail"),
                "business_phones": user_data.get("businessPhones", []),
                "mobile_phone": user_data.get("mobilePhone"),
            },
        }

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R10: MS365 팀 및 채널 검색
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_channels_search(
        team_name: Annotated[str, Field(description="검색할 팀 이름 (부분 일치). 비워두면 가입된 전체 팀 목록 반환")] = "",
        channel_name: Annotated[str, Field(description="검색할 채널 이름 (부분 일치). 비워두면 검색된 팀의 전체 채널 반환")] = "",
    ) -> dict:
        """사용자가 가입한 Teams의 팀과 하위 채널을 이름으로 검색하여 ID를 조회합니다.
        team_name이 비어있으면 사용자가 가입한 전체 팀 목록만 반환합니다.
        team_name을 지정하면 해당 키워드가 포함된 팀들과 그 안의 채널 목록을 조회하여 반환합니다.
        Azure AD 앱에 Team.ReadBasic.All, Channel.ReadBasic.All Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        teams_out = []
        errors = []

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                # 1. 가입된 전체 팀 목록 조회
                # GET /me/joinedTeams
                teams_resp = await client.get(
                    "https://graph.microsoft.com/v1.0/me/joinedTeams",
                    headers=headers,
                )
                teams_resp.raise_for_status()
                all_teams = teams_resp.json().get("value", [])

                # 2. team_name 키워드로 필터링
                if team_name:
                    filtered_teams = [
                        t for t in all_teams 
                        if team_name.lower() in (t.get("displayName") or "").lower()
                    ]
                else:
                    filtered_teams = all_teams

                # 3. 채널 정보 조회 (team_name이 있을 때만 상세 조회, 없으면 팀 목록만 반환)
                for t in filtered_teams:
                    team_id = t.get("id")
                    team_info = {
                        "team_id": team_id,
                        "display_name": t.get("displayName"),
                        "description": t.get("description"),
                        "channels": []
                    }

                    if team_name:
                        try:
                            channels_resp = await client.get(
                                f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels",
                                headers=headers,
                            )
                            if channels_resp.status_code == 200:
                                channels_data = channels_resp.json().get("value", [])
                                
                                # channel_name 필터링
                                if channel_name:
                                    filtered_channels = [
                                        c for c in channels_data
                                        if channel_name.lower() in (c.get("displayName") or "").lower()
                                    ]
                                else:
                                    filtered_channels = channels_data

                                for c in filtered_channels:
                                    team_info["channels"].append({
                                        "channel_id": c.get("id"),
                                        "display_name": c.get("displayName"),
                                        "description": c.get("description"),
                                    })
                        except Exception as e:
                            log.warning("팀(%s) 채널 조회 실패: %s", team_id, e)
                            errors.append(f"channels for {team_id}: {e}")

                    teams_out.append(team_info)

            log.info("teams_channels_search: %d 팀 반환", len(teams_out))
            return {
                "success": True,
                "team_name_query": team_name,
                "channel_name_query": channel_name,
                "team_count": len(teams_out),
                "teams": teams_out,
                **({"warnings": errors} if errors else {}),
            }

        except httpx.HTTPStatusError as e:
            log.error("teams_channels_search HTTP 오류: %s", e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": f"Graph API 오류: {e.response.text}"}
        except Exception as e:
            log.error("teams_channels_search 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R11: Teams 그룹 채팅방 메시지 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_group_chat_messages(
        chat_id: Annotated[str, Field(description="그룹 채팅방 ID. 비워두면 내 그룹 채팅 목록을 반환합니다.")] = "",
        date: Annotated[str, Field(
            pattern=r"^(\d{4}-\d{2}-\d{2})?$",
            description="조회 날짜 YYYY-MM-DD. 비우면 오늘",
        )] = "",
        limit: Annotated[int, Field(ge=1, le=100, description="최대 반환 건수 (기본 50)")] = 50,
    ) -> dict:
        """Teams 그룹 채팅방의 메시지를 날짜 기준으로 조회합니다.
        - chat_id 없음: 내가 속한 그룹 채팅 목록을 반환합니다. (chat_id를 확인한 뒤 다시 호출하세요)
        - chat_id 있음: 해당 채팅방의 지정 날짜 메시지를 조회합니다.
        Azure AD 앱에 Chat.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx
        from datetime import datetime
        from zoneinfo import ZoneInfo

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        KST = ZoneInfo("Asia/Seoul")

        # ── 채팅방 목록 반환 ────────────────────────────────────
        if not chat_id:
            try:
                async with httpx.AsyncClient(timeout=30) as client:
                    # $expand=members 는 /me/chats 에서 미지원 → 목록만 가져온 뒤 멤버 별도 조회
                    resp = await client.get(
                        "https://graph.microsoft.com/v1.0/me/chats",
                        params={
                            "$select": "id,topic,chatType,createdDateTime,lastUpdatedDateTime",
                            "$top": 50,
                        },
                        headers=headers,
                    )
                    resp.raise_for_status()
                    data = resp.json()

                    # 그룹 채팅만 필터링 (chatType 값: oneOnOne | group | meeting)
                    group_chats = [c for c in data.get("value", []) if c.get("chatType") == "group"]

                    # 각 채팅방 멤버를 별도 조회
                    chats_out = []
                    for c in group_chats:
                        members = []
                        try:
                            m_resp = await client.get(
                                f"https://graph.microsoft.com/v1.0/chats/{c['id']}/members",
                                params={"$select": "displayName,email"},
                                headers=headers,
                            )
                            if m_resp.status_code == 200:
                                members = [
                                    m.get("displayName") or m.get("email", "")
                                    for m in m_resp.json().get("value", [])
                                ]
                        except Exception:
                            pass

                        chats_out.append({
                            "chat_id": c.get("id"),
                            "topic": c.get("topic") or "(제목 없음)",
                            "members": members,
                            "last_updated": c.get("lastUpdatedDateTime"),
                        })

                return {
                    "success": True,
                    "message": "아래 chat_id 중 하나를 선택해 다시 호출하면 메시지를 조회할 수 있습니다.",
                    "count": len(chats_out),
                    "chats": chats_out,
                }
            except httpx.HTTPStatusError as e:
                log.error("teams_group_chat_messages(list) HTTP 오류: %s", e.response.text)
                return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
            except Exception as e:
                log.error("teams_group_chat_messages(list) 실패: %s", e)
                return {"success": False, "message": str(e)}

        # ── 메시지 조회 ─────────────────────────────────────────
        # /chats/{id}/messages 는 $filter=createdDateTime 미지원
        # → $filter 없이 최신순으로 가져오며 페이지네이션, Python에서 날짜 필터링
        target_date = date or datetime.now(KST).strftime("%Y-%m-%d")

        try:
            messages_out = []
            next_url: str = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
            next_params: dict | None = {"$top": 50}

            async with httpx.AsyncClient(timeout=60) as client:
                while next_url and len(messages_out) < limit:
                    resp = await client.get(next_url, params=next_params, headers=headers)
                    resp.raise_for_status()
                    data = resp.json()
                    next_params = None  # 이후 페이지는 nextLink 자체에 파라미터 포함

                    stop = False
                    for msg in data.get("value", []):
                        if msg.get("messageType") != "message":
                            continue
                        created_at = msg.get("createdDateTime") or ""
                        msg_date = _to_kst_date(created_at)
                        if msg_date > target_date:
                            continue          # 대상 날짜보다 미래 메시지 → 스킵
                        if not msg_date:
                            continue          # 파싱 실패 → 스킵 (stop 하지 않음)
                        if msg_date < target_date:
                            stop = True       # 대상 날짜보다 과거 → 더 이상 페이지 불필요
                            break
                        sender_info = msg.get("from") or {}
                        sender = (
                            (sender_info.get("user") or {}).get("displayName")
                            or (sender_info.get("application") or {}).get("displayName")
                            or "알 수 없음"
                        )
                        body_content = (msg.get("body") or {}).get("content") or ""
                        messages_out.append({
                            "id": msg.get("id"),
                            "sender": sender,
                            "created_at": created_at,
                            "body": body_content,
                            "attachments": len(msg.get("attachments") or []),
                        })
                        if len(messages_out) >= limit:
                            stop = True
                            break

                    if stop:
                        break
                    next_url = data.get("@odata.nextLink", "")

            messages_out.sort(key=lambda m: m["created_at"] or "")
            return {
                "success": True,
                "chat_id": chat_id,
                "date": target_date,
                "count": len(messages_out),
                "messages": messages_out,
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_group_chat_messages(messages) HTTP 오류: %s", e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_group_chat_messages(messages) 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R12: Teams 팀 채널 메시지 날짜 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_channel_messages(
        team_id: Annotated[str, Field(min_length=1, description="Teams 팀 ID (teams_channels_search 도구로 확인)")],
        channel_id: Annotated[str, Field(min_length=1, description="Teams 채널 ID (teams_channels_search 도구로 확인)")],
        date: Annotated[str, Field(
            pattern=r"^(\d{4}-\d{2}-\d{2})?$",
            description="조회 날짜 YYYY-MM-DD. 비우면 오늘",
        )] = "",
        limit: Annotated[int, Field(ge=1, le=100, description="최대 반환 건수 (기본 50)")] = 50,
    ) -> dict:
        """Teams 팀 채널의 메시지를 날짜 기준으로 조회합니다.
        team_id와 channel_id는 teams_channels_search 도구로 먼저 확인하세요.
        Azure AD 앱에 ChannelMessage.Read.All Delegated 권한 및 관리자 동의가 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx
        from datetime import datetime
        from zoneinfo import ZoneInfo

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        KST = ZoneInfo("Asia/Seoul")
        target_date = date or datetime.now(KST).strftime("%Y-%m-%d")

        # /teams/{id}/channels/{id}/messages 도 $filter=createdDateTime 미지원
        # → 페이지네이션으로 최신순 수신 후 Python에서 날짜 필터링
        try:
            messages_out = []
            next_url: str = (
                f"https://graph.microsoft.com/v1.0/teams/{team_id}"
                f"/channels/{channel_id}/messages"
            )
            next_params: dict | None = {"$top": 50}

            async with httpx.AsyncClient(timeout=60) as client:
                while next_url and len(messages_out) < limit:
                    resp = await client.get(next_url, params=next_params, headers=headers)
                    resp.raise_for_status()
                    data = resp.json()
                    next_params = None

                    stop = False
                    for msg in data.get("value", []):
                        if msg.get("messageType") != "message":
                            continue
                        created_at = msg.get("createdDateTime") or ""
                        msg_date = _to_kst_date(created_at)
                        if msg_date > target_date:
                            continue
                        if not msg_date:
                            continue          # 파싱 실패 → 스킵 (stop 하지 않음)
                        if msg_date < target_date:
                            stop = True
                            break
                        sender_info = msg.get("from") or {}
                        sender = (sender_info.get("user") or {}).get("displayName") or "알 수 없음"
                        body_content = (msg.get("body") or {}).get("content") or ""
                        messages_out.append({
                            "id": msg.get("id"),
                            "sender": sender,
                            "created_at": created_at,
                            "body": body_content,
                            "reactions": [r.get("reactionType") for r in (msg.get("reactions") or [])],
                            "web_url": msg.get("webUrl"),
                        })
                        if len(messages_out) >= limit:
                            stop = True
                            break

                    if stop:
                        break
                    next_url = data.get("@odata.nextLink", "")

            messages_out.sort(key=lambda m: m["created_at"] or "")
            return {
                "success": True,
                "team_id": team_id,
                "channel_id": channel_id,
                "date": target_date,
                "count": len(messages_out),
                "messages": messages_out,
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_channel_messages HTTP 오류: %s", e.response.text)
            if e.response.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": (
                        "권한 오류(403): Azure AD 앱에 'ChannelMessage.Read.All' Delegated 권한과 "
                        "관리자 동의가 필요합니다. Azure Portal에서 권한을 추가한 뒤 ms365_logout → ms365_login으로 재인증하세요."
                    ),
                }
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_channel_messages 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R13: 이메일 첨부파일 읽기
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def read_email_attachment(
        message_id: Annotated[str, Field(min_length=1, description="이메일 ID (outlook_email_search 반환 'id' 값)")],
        attachment_id: Annotated[str, Field(description="첨부파일 ID. 비우면 첨부파일 목록을 반환합니다.")] = "",
        owner: Annotated[str, Field(description="이메일 소유자 이메일. 비우면 로그인 계정.")] = "",
        max_chars: Annotated[int, Field(ge=1000, le=200000, description="반환할 최대 문자 수 (기본 20000)")] = 20000,
    ) -> dict:
        """Outlook 이메일의 첨부파일을 읽습니다.

        - attachment_id 없음: 이메일의 첨부파일 목록(id, name, size, content_type 등)을 반환합니다.
          Graph API 제약상 목록 조회 시에는 attachment_type을 정확히 구분하지 못하고 기본값 'file'로 표시됩니다
          (참조/아이템 첨부파일 여부는 attachment_id로 개별 조회해야 정확히 확인 가능).
        - attachment_id 있음: 해당 첨부파일 내용을 추출하여 반환합니다.
          xlsx/docx/pptx/pdf는 텍스트로 추출하고, 참조 첨부파일(OneDrive/SharePoint 링크)도 지원합니다.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        import httpx as _httpx

        auth_err = require_auth()
        if auth_err:
            return auth_err
        await fetch_user_profile()
        resource_owner = owner or get_user_email()
        token = get_access_token()
        if not token:
            return {"success": False, "message": "액세스 토큰을 가져올 수 없습니다. ms365_login으로 다시 로그인해주세요."}

        base_url = f"https://graph.microsoft.com/v1.0/users/{resource_owner}/messages/{message_id}/attachments"
        headers = {"Authorization": f"Bearer {token}"}

        try:
            if not attachment_id:
                # 첨부파일 목록 반환
                async with _httpx.AsyncClient(timeout=30) as client:
                    resp = await client.get(
                        base_url,
                        params={"$select": "id,name,size,contentType,isInline"},
                        headers=headers,
                    )
                if resp.status_code != 200:
                    return {"success": False, "message": f"첨부파일 목록 조회 실패: HTTP {resp.status_code} — {resp.text[:300]}"}
                items = resp.json().get("value", [])
                attachments_out = []
                for att in items:
                    odata_type = att.get("@odata.type", "")
                    att_type = "file"
                    if "referenceAttachment" in odata_type:
                        att_type = "reference"
                    elif "itemAttachment" in odata_type:
                        att_type = "item"
                    attachments_out.append({
                        "id": att.get("id"),
                        "name": att.get("name"),
                        "size": att.get("size"),
                        "content_type": att.get("contentType"),
                        "is_inline": att.get("isInline", False),
                        "attachment_type": att_type,
                        "source_url": att.get("sourceUrl"),  # referenceAttachment 전용
                    })
                return {
                    "success": True,
                    "message_id": message_id,
                    "owner": resource_owner,
                    "count": len(attachments_out),
                    "attachments": attachments_out,
                }

            # 특정 첨부파일 내용 읽기
            async with _httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(
                    f"{base_url}/{attachment_id}",
                    headers=headers,
                )
            if resp.status_code != 200:
                return {"success": False, "message": f"첨부파일 조회 실패: HTTP {resp.status_code}"}

            att_data = resp.json()
            odata_type = att_data.get("@odata.type", "")

            if "itemAttachment" in odata_type:
                return {
                    "success": False,
                    "message": "itemAttachment(이메일/일정 항목 첨부)는 현재 지원하지 않습니다. fileAttachment 또는 referenceAttachment만 읽을 수 있습니다.",
                    "attachment_type": "item",
                }

            if "referenceAttachment" in odata_type:
                source_url = att_data.get("sourceUrl", "")
                if not source_url:
                    return {"success": False, "message": "참조 첨부파일의 sourceUrl을 찾을 수 없습니다."}
                meta = await _resolve_file_meta(source_url, token)
                if "error" in meta:
                    return {"success": False, "message": meta["error"], "source_url": source_url}

                file_size = meta.get("size", 0)
                MAX_BYTES = 20 * 1024 * 1024
                if file_size > MAX_BYTES:
                    return {
                        "success": False,
                        "message": f"파일 크기({file_size / 1024 / 1024:.1f}MB)가 20MB 제한을 초과합니다.",
                        "filename": meta.get("name"),
                        "size": file_size,
                    }
                drive_id = meta.get("drive_id")
                item_id = meta.get("item_id")
                async with _httpx.AsyncClient(timeout=60) as client:
                    dl_resp = await client.get(
                        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
                        headers={"Authorization": f"Bearer {token}"},
                        follow_redirects=True,
                    )
                if dl_resp.status_code != 200:
                    return {"success": False, "message": f"참조 파일 다운로드 실패: HTTP {dl_resp.status_code}"}
                content_bytes = dl_resp.content

            else:
                # fileAttachment — /$value 로 바이너리 취득
                async with _httpx.AsyncClient(timeout=60) as client:
                    dl_resp = await client.get(
                        f"{base_url}/{attachment_id}/$value",
                        headers=headers,
                        follow_redirects=True,
                    )
                if dl_resp.status_code != 200:
                    return {"success": False, "message": f"첨부파일 다운로드 실패: HTTP {dl_resp.status_code}"}
                content_bytes = dl_resp.content
                meta = {
                    "name": att_data.get("name", ""),
                    "size": att_data.get("size", len(content_bytes)),
                    "mime_type": att_data.get("contentType", ""),
                }

            if not content_bytes:
                return {"success": True, "attachment_id": attachment_id, "content": ""}

            extracted = _extract_file_content(
                content_bytes,
                meta.get("mime_type", att_data.get("contentType", "")),
                meta.get("name", att_data.get("name", "")),
                max_chars,
            )
            return {
                "success": True,
                "message_id": message_id,
                "attachment_id": attachment_id,
                "filename": meta.get("name", att_data.get("name")),
                "size": meta.get("size", len(content_bytes)),
                "file_type": extracted["file_type"],
                "content": extracted["content"],
                "truncated": extracted["truncated"],
                "original_length": extracted["original_length"],
            }

        except Exception as e:
            log.error("read_email_attachment 실패 (message_id=%s): %s", message_id, e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R14: 이메일 첨부파일 전체 내용 일괄 읽기
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def read_email_attachments_all(
        message_id: Annotated[str, Field(min_length=1, description="이메일 ID (outlook_email_search 반환 'id' 값)")],
        owner: Annotated[str, Field(description="이메일 소유자 이메일. 비우면 로그인 계정.")] = "",
        max_chars_per_file: Annotated[int, Field(ge=500, le=50000, description="첨부파일당 최대 반환 문자 수 (기본 5000)")] = 5000,
        include_inline: Annotated[bool, Field(description="본문 삽입용 인라인 첨부파일(서명 이미지 등)도 포함할지 여부")] = False,
    ) -> dict:
        """이메일 1건의 첨부파일 전체 내용을 한 번의 호출로 읽습니다.

        read_email_attachment처럼 목록 조회 후 attachment_id로 한 개씩 다시 호출할 필요 없이,
        message_id만으로 모든 fileAttachment/referenceAttachment의 내용을 읽어 반환합니다.
        xlsx/docx/pptx/pdf는 텍스트로 추출합니다.
        itemAttachment(메일/일정 첨부), 20MB 초과 파일, 다운로드 실패 항목은 'skipped'에 사유와 함께 표시됩니다.
        Graph API 제약상 목록 단계에서는 첨부파일 유형을 구분하지 못해 itemAttachment/referenceAttachment도
        일반 파일 다운로드를 먼저 시도하며, 실패 시 'skipped'에 HTTP 오류 사유로 표시됩니다.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        import httpx as _httpx

        auth_err = require_auth()
        if auth_err:
            return auth_err
        await fetch_user_profile()
        resource_owner = owner or get_user_email()
        token = get_access_token()
        if not token:
            return {"success": False, "message": "액세스 토큰을 가져올 수 없습니다. ms365_login으로 다시 로그인해주세요."}

        base_url = f"https://graph.microsoft.com/v1.0/users/{resource_owner}/messages/{message_id}/attachments"
        headers = {"Authorization": f"Bearer {token}"}
        MAX_BYTES = 20 * 1024 * 1024

        try:
            async with _httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(
                    base_url,
                    params={"$select": "id,name,size,contentType,isInline"},
                    headers=headers,
                )
            if resp.status_code != 200:
                return {"success": False, "message": f"첨부파일 목록 조회 실패: HTTP {resp.status_code} — {resp.text[:300]}"}
            items = resp.json().get("value", [])

            results = []
            skipped = []

            for att in items:
                att_id = att.get("id")
                name = att.get("name", "")
                odata_type = att.get("@odata.type", "")

                if att.get("isInline", False) and not include_inline:
                    skipped.append({"id": att_id, "name": name, "reason": "인라인 첨부파일 (include_inline=true로 포함 가능)"})
                    continue
                if "itemAttachment" in odata_type:
                    skipped.append({"id": att_id, "name": name, "reason": "itemAttachment(메일/일정 첨부)는 지원하지 않음"})
                    continue

                try:
                    if "referenceAttachment" in odata_type:
                        source_url = att.get("sourceUrl", "")
                        if not source_url:
                            skipped.append({"id": att_id, "name": name, "reason": "sourceUrl 없음"})
                            continue
                        meta = await _resolve_file_meta(source_url, token)
                        if "error" in meta:
                            skipped.append({"id": att_id, "name": name, "reason": meta["error"]})
                            continue
                        file_size = meta.get("size") or 0
                        if file_size > MAX_BYTES:
                            skipped.append({"id": att_id, "name": name, "reason": f"크기 초과 ({file_size / 1024 / 1024:.1f}MB > 20MB)"})
                            continue
                        async with _httpx.AsyncClient(timeout=60) as client:
                            dl_resp = await client.get(
                                f"https://graph.microsoft.com/v1.0/drives/{meta.get('drive_id')}/items/{meta.get('item_id')}/content",
                                headers=headers,
                                follow_redirects=True,
                            )
                        if dl_resp.status_code != 200:
                            skipped.append({"id": att_id, "name": name, "reason": f"참조 파일 다운로드 실패: HTTP {dl_resp.status_code}"})
                            continue
                        content_bytes = dl_resp.content
                        mime_type = meta.get("mime_type", "")
                        filename = meta.get("name") or name
                        size = meta.get("size", len(content_bytes))
                    else:
                        size = att.get("size") or 0
                        if size > MAX_BYTES:
                            skipped.append({"id": att_id, "name": name, "reason": f"크기 초과 ({size / 1024 / 1024:.1f}MB > 20MB)"})
                            continue
                        async with _httpx.AsyncClient(timeout=60) as client:
                            dl_resp = await client.get(
                                f"{base_url}/{att_id}/$value",
                                headers=headers,
                                follow_redirects=True,
                            )
                        if dl_resp.status_code != 200:
                            skipped.append({"id": att_id, "name": name, "reason": f"다운로드 실패: HTTP {dl_resp.status_code}"})
                            continue
                        content_bytes = dl_resp.content
                        mime_type = att.get("contentType", "")
                        filename = name
                        size = att.get("size", len(content_bytes))

                    if not content_bytes:
                        results.append({
                            "id": att_id, "name": filename, "size": size,
                            "content_type": mime_type, "file_type": "empty", "content": "", "truncated": False,
                        })
                        continue

                    extracted = _extract_file_content(content_bytes, mime_type, filename, max_chars_per_file)
                    results.append({
                        "id": att_id,
                        "name": filename,
                        "size": size,
                        "content_type": mime_type,
                        "file_type": extracted["file_type"],
                        "content": extracted["content"],
                        "truncated": extracted["truncated"],
                        "original_length": extracted["original_length"],
                    })
                except Exception as e:
                    skipped.append({"id": att_id, "name": name, "reason": str(e)})

            log.info(
                "read_email_attachments_all: message=%s owner=%s read=%d skipped=%d",
                message_id, resource_owner, len(results), len(skipped),
            )
            return {
                "success": True,
                "message_id": message_id,
                "owner": resource_owner,
                "count": len(results),
                "attachments": results,
                "skipped": skipped,
            }
        except Exception as e:
            log.error("read_email_attachments_all 실패 (message_id=%s): %s", message_id, e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R15: Excel 워크시트 목록 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True})
    async def excel_list_worksheets(
        drive_id: Annotated[str, Field(description="OneDrive/SharePoint 드라이브 ID")],
        item_id: Annotated[str, Field(description="Excel 파일 item ID")],
    ) -> dict:
        """Excel 파일의 워크시트 목록을 조회합니다. excel_read_range/excel_write_range 호출 전 정확한 워크시트 이름을 확인할 때 사용하세요."""
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/workbook/worksheets",
                    headers={"Authorization": f"Bearer {access_token}"},
                )
                resp.raise_for_status()
                data = resp.json()

            worksheets = [
                {"id": ws.get("id"), "name": ws.get("name"), "position": ws.get("position")}
                for ws in data.get("value", [])
            ]
            log.info("excel_list_worksheets: drive=%s item=%s count=%d", drive_id, item_id, len(worksheets))
            return {"success": True, "worksheets": worksheets}
        except httpx.HTTPStatusError as e:
            log.error("excel_list_worksheets HTTP 오류: %s", e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("excel_list_worksheets 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R16: Excel 워크시트 범위 읽기
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True})
    async def excel_read_range(
        drive_id: Annotated[str, Field(description="OneDrive/SharePoint 드라이브 ID")],
        item_id: Annotated[str, Field(description="Excel 파일 item ID")],
        worksheet: Annotated[str, Field(description="워크시트 이름 또는 ID (excel_list_worksheets로 확인)")],
        address: Annotated[str, Field(description="읽을 범위 주소 (예: 'A1', 'A1:C10'). 비우면 사용 중인 전체 범위(usedRange)를 읽습니다.")] = "",
    ) -> dict:
        """Excel 워크시트의 범위 값을 조회합니다.
        address를 지정하지 않으면 워크시트의 사용 중인 전체 범위(usedRange)를 반환합니다.
        worksheet는 excel_list_worksheets 도구로 정확한 이름을 먼저 확인하세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx
        from urllib.parse import quote

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        encoded_worksheet = quote(worksheet, safe="")
        if address:
            url = (
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
                f"/workbook/worksheets/{encoded_worksheet}/range(address='{address}')"
            )
        else:
            url = (
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
                f"/workbook/worksheets/{encoded_worksheet}/usedRange"
            )

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(url, headers={"Authorization": f"Bearer {access_token}"})
                resp.raise_for_status()
                data = resp.json()

            log.info(
                "excel_read_range: drive=%s item=%s worksheet=%s address=%s",
                drive_id, item_id, worksheet, address or "(usedRange)",
            )
            return {
                "success": True,
                "worksheet": worksheet,
                "address": data.get("address"),
                "values": data.get("values"),
                "row_count": data.get("rowCount"),
                "column_count": data.get("columnCount"),
            }
        except httpx.HTTPStatusError as e:
            log.error("excel_read_range HTTP 오류: %s", e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("excel_read_range 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R17: 사용량 보고서 조회 (Reports API)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def org_usage_report(
        report_type: Annotated[
            Literal[
                "office365_active_users", "mailbox_usage", "onedrive_usage",
                "sharepoint_usage", "teams_user_activity", "teams_device_usage",
                "email_activity",
            ],
            Field(description=(
                "조회할 보고서 종류: "
                "office365_active_users(서비스별 최근 활동일) · mailbox_usage(메일박스 용량) · "
                "onedrive_usage(OneDrive 저장 용량) · sharepoint_usage(SharePoint 사이트 저장 용량) · "
                "teams_user_activity(Teams 채팅/통화/회의 활동) · teams_device_usage(Teams 접속 기기 유형) · "
                "email_activity(이메일 송수신 활동)"
            )),
        ],
        period: Annotated[
            Literal["D7", "D30", "D90", "D180"],
            Field(description="집계 기간: D7/D30/D90/D180일 (기본 D30)"),
        ] = "D30",
        limit: Annotated[int, Field(ge=1, le=500, description="최대 반환 행 수 (기본 50, 사용자 단위 상세 보고서라 전 직원 수만큼 행이 나올 수 있음)")] = 50,
    ) -> dict:
        """조직의 Microsoft 365 사용량 보고서를 조회합니다 (사용자/사이트 단위 상세 보고서, CSV 원본을 파싱해 반환).

        report_type으로 원하는 보고서를 선택하고, period로 집계 기간(D7/D30/D90/D180)을 지정하세요.
        결과 컬럼명은 Graph API가 반환하는 원본 그대로입니다 (예: "User Principal Name", "Last Activity Date").

        Azure AD 앱에 Reports.Read.All Delegated 권한이 필요합니다.
        또한 로그인 계정이 Reports Reader 이상의 Azure AD 관리자 역할을 보유해야 합니다.
        일반 직원 계정으로 호출하면 403 오류가 반환됩니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        endpoint = _USAGE_REPORT_ENDPOINTS[report_type]
        url = f"https://graph.microsoft.com/v1.0/reports/{endpoint}(period='{period}')"
        headers = {"Authorization": f"Bearer {access_token}"}

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(url, headers=headers, follow_redirects=True)
            if resp.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": (
                        "권한 오류(403): Azure AD 앱에 'Reports.Read.All' Delegated 권한이 필요하거나, "
                        "로그인 계정에 Reports Reader 이상의 관리자 역할이 없습니다. "
                        "권한을 새로 추가했다면 ms365_logout → ms365_login으로 재인증하세요."
                    ),
                }
            resp.raise_for_status()

            rows = _parse_usage_report_csv(resp.content.decode("utf-8-sig"))
            total_rows = len(rows)
            truncated = total_rows > limit
            log.info(
                "org_usage_report: type=%s period=%s rows=%d (by %s)",
                report_type, period, total_rows, get_user_email(),
            )
            return {
                "success": True,
                "report_type": report_type,
                "period": period,
                "total_rows": total_rows,
                "truncated": truncated,
                "rows": rows[:limit],
            }
        except httpx.HTTPStatusError as e:
            log.error("org_usage_report HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("org_usage_report 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R18: Teams 채널 메시지 댓글(답글) 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_channel_message_replies(
        team_id: Annotated[str, Field(min_length=1, description="Teams 팀 ID (teams_channels_search 도구로 확인)")],
        channel_id: Annotated[str, Field(min_length=1, description="Teams 채널 ID (teams_channels_search 도구로 확인)")],
        message_id: Annotated[str, Field(min_length=1, description="댓글을 조회할 원본 채널 메시지 ID (chat_message_search 또는 teams_channel_messages 반환 'id' 값)")],
        limit: Annotated[int, Field(ge=1, le=50, description="최대 반환 건수 (기본 50). Graph API 특성상 한 번에 최대 50건까지만 조회됩니다")] = 50,
    ) -> dict:
        """Teams 채널 메시지의 댓글(스레드 답글) 목록을 조회합니다.
        message_id는 chat_message_search 또는 teams_channel_messages 도구로 먼저 확인하세요.
        각 답글의 attachment_count로 첨부 여부를 먼저 확인한 뒤, 첨부파일 내용은
        teams_channel_message_attachments, 본문에 붙여넣은 이미지는 teams_channel_message_images
        도구에 이 도구가 반환한 답글 id를 reply_id로 넘겨 조회하세요.
        시스템 이벤트(참가/퇴장 등)와 messageType이 'message'가 아닌 항목은 replies 목록에서 제외되며,
        raw_count(Graph API가 반환한 전체 항목 수)와 count(실제 반환된 대화형 답글 수)가 다르면
        제외된 항목이 있다는 뜻입니다(삭제된 답글도 messageType은 유지되므로 deletedDateTime으로 확인 가능).
        Azure AD 앱에 ChannelMessage.Read.All Delegated 권한 및 관리자 동의가 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(url, params={"$top": min(limit, 50)}, headers=headers)
                resp.raise_for_status()
                data = resp.json()

            raw_items = data.get("value", [])
            replies_out = []
            for msg in raw_items:
                if msg.get("messageType") != "message":
                    continue
                sender_info = msg.get("from") or {}
                sender = (sender_info.get("user") or {}).get("displayName") or "알 수 없음"
                body_content = (msg.get("body") or {}).get("content") or ""
                replies_out.append({
                    "id": msg.get("id"),
                    "reply_to_id": msg.get("replyToId"),
                    "sender": sender,
                    "created_at": msg.get("createdDateTime"),
                    "deleted_at": msg.get("deletedDateTime"),
                    "body": body_content,
                    "reactions": [r.get("reactionType") for r in (msg.get("reactions") or [])],
                    "attachment_count": len(msg.get("attachments") or []),
                    "web_url": msg.get("webUrl"),
                })
                if len(replies_out) >= limit:
                    break

            replies_out.sort(key=lambda m: m["created_at"] or "")
            log.info("teams_channel_message_replies: team=%s channel=%s message=%s count=%d",
                     team_id, channel_id, message_id, len(replies_out))
            return {
                "success": True,
                "team_id": team_id,
                "channel_id": channel_id,
                "message_id": message_id,
                "raw_count": len(raw_items),
                "count": len(replies_out),
                "replies": replies_out,
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_channel_message_replies HTTP 오류: %s", e.response.text)
            if e.response.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": (
                        "권한 오류(403): Azure AD 앱에 'ChannelMessage.Read.All' Delegated 권한과 "
                        "관리자 동의가 필요합니다. Azure Portal에서 권한을 추가한 뒤 ms365_logout → ms365_login으로 재인증하세요."
                    ),
                }
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"메시지를 찾을 수 없습니다: {message_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_channel_message_replies 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R19: Teams 채널 메시지 첨부파일 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_channel_message_attachments(
        team_id: Annotated[str, Field(min_length=1, description="Teams 팀 ID (teams_channels_search 도구로 확인)")],
        channel_id: Annotated[str, Field(min_length=1, description="Teams 채널 ID (teams_channels_search 도구로 확인)")],
        message_id: Annotated[str, Field(min_length=1, description="첨부파일을 조회할 메시지 ID (chat_message_search, teams_channel_messages, teams_channel_message_replies 반환 'id' 값)")],
        reply_id: Annotated[str, Field(description="답글에 붙은 첨부파일을 조회할 경우 답글 메시지 ID. 원본 채널 메시지를 조회할 때는 비워둡니다.")] = "",
    ) -> dict:
        """Teams 채널 메시지(또는 답글)에 첨부된 파일·카드 목록을 조회합니다.
        Teams에서 공유된 파일은 contentType='reference'로 표시되며, content_url에 SharePoint/OneDrive
        공유 링크가 담겨 있습니다. 실제 파일 내용을 읽으려면 content_url을 sharepoint_search로 다시 검색해
        drive_id/item_id를 확인하거나, read_resource 도구에 URL을 그대로 전달하세요.
        contentType이 'reference'가 아닌 항목(공지·코드스니펫 등 카드류, application/vnd.microsoft.card.*)은
        사용자가 말하는 "첨부파일"이 아닌 경우가 많아 file_count와 card_count를 구분해 반환합니다.
        본문에 붙여넣은 인라인 이미지는 여기 나타나지 않습니다. 이미지는 teams_channel_message_images 도구를 사용하세요.
        Azure AD 앱에 ChannelMessage.Read.All Delegated 권한 및 관리자 동의가 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {"Authorization": f"Bearer {access_token}"}
        base = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}"
        url = f"{base}/replies/{reply_id}" if reply_id else base

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(url, headers=headers)
                resp.raise_for_status()
                data = resp.json()

            attachments_out = [
                {
                    "id": att.get("id"),
                    "name": att.get("name"),
                    "content_type": att.get("contentType"),
                    "content_url": att.get("contentUrl"),
                    "thumbnail_url": att.get("thumbnailUrl"),
                }
                for att in (data.get("attachments") or [])
            ]
            file_count = sum(1 for a in attachments_out if a["content_type"] == "reference")
            card_count = len(attachments_out) - file_count

            log.info("teams_channel_message_attachments: team=%s channel=%s message=%s reply=%s files=%d cards=%d",
                     team_id, channel_id, message_id, reply_id or "-", file_count, card_count)
            return {
                "success": True,
                "team_id": team_id,
                "channel_id": channel_id,
                "message_id": reply_id or message_id,
                "count": len(attachments_out),
                "file_count": file_count,
                "card_count": card_count,
                "attachments": attachments_out,
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_channel_message_attachments HTTP 오류: %s", e.response.text)
            if e.response.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": (
                        "권한 오류(403): Azure AD 앱에 'ChannelMessage.Read.All' Delegated 권한과 "
                        "관리자 동의가 필요합니다. Azure Portal에서 권한을 추가한 뒤 ms365_logout → ms365_login으로 재인증하세요."
                    ),
                }
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"메시지를 찾을 수 없습니다: {reply_id or message_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_channel_message_attachments 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R20: Teams 채널 메시지 인라인 이미지 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_channel_message_images(
        team_id: Annotated[str, Field(min_length=1, description="Teams 팀 ID (teams_channels_search 도구로 확인)")],
        channel_id: Annotated[str, Field(min_length=1, description="Teams 채널 ID (teams_channels_search 도구로 확인)")],
        message_id: Annotated[str, Field(min_length=1, description="이미지를 조회할 메시지 ID (chat_message_search, teams_channel_messages, teams_channel_message_replies 반환 'id' 값)")],
        reply_id: Annotated[str, Field(description="답글에 포함된 이미지를 조회할 경우 답글 메시지 ID. 원본 채널 메시지를 조회할 때는 비워둡니다.")] = "",
        hosted_content_id: Annotated[str, Field(description="내용을 읽을 이미지 1개의 ID (목록 조회로 먼저 확인). 비우면 이미지 ID 목록만 반환합니다.")] = "",
    ) -> dict:
        """Teams 채널 메시지(또는 답글) 본문에 붙여넣은 인라인 이미지를 조회합니다.
        - hosted_content_id 미지정: 메시지에 포함된 이미지(hostedContent) ID 목록만 반환합니다(내용 미포함).
          Graph API 제약상 목록 조회 단계에서는 이미지 크기·형식을 알 수 없습니다.
        - hosted_content_id 지정: 해당 이미지 1개의 실제 바이너리를 base64로 인코딩해 content_type과 함께 반환합니다.
          이미지 크기가 3MB를 초과하면 내용을 반환하지 않고 오류를 안내합니다.
        파일 형태로 공유된 첨부(카드/링크)는 여기 나타나지 않습니다. teams_channel_message_attachments 도구를 사용하세요.
        Azure AD 앱에 ChannelMessage.Read.All Delegated 권한 및 관리자 동의가 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import base64
        import httpx
        from urllib.parse import quote

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {"Authorization": f"Bearer {access_token}"}
        base = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}"
        msg_base = f"{base}/replies/{reply_id}" if reply_id else base

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                if not hosted_content_id:
                    resp = await client.get(f"{msg_base}/hostedContents", headers=headers)
                    resp.raise_for_status()
                    data = resp.json()
                    images_out = [{"hosted_content_id": item.get("id")} for item in data.get("value", [])]

                    log.info("teams_channel_message_images(목록): team=%s channel=%s message=%s reply=%s count=%d",
                             team_id, channel_id, message_id, reply_id or "-", len(images_out))
                    return {
                        "success": True,
                        "team_id": team_id,
                        "channel_id": channel_id,
                        "message_id": reply_id or message_id,
                        "count": len(images_out),
                        "images": images_out,
                    }

                content_url = f"{msg_base}/hostedContents/{quote(hosted_content_id, safe='')}/$value"
                async with client.stream("GET", content_url, headers=headers) as stream_resp:
                    if stream_resp.status_code >= 400:
                        # 스트리밍 응답은 본문을 먼저 읽어둬야 아래 except에서 e.response.text 접근이 가능하다.
                        await stream_resp.aread()
                    stream_resp.raise_for_status()

                    declared_length = stream_resp.headers.get("content-length")
                    if declared_length and int(declared_length) > _MAX_HOSTED_CONTENT_BYTES:
                        return {
                            "success": False,
                            "message": (
                                f"이미지 크기({int(declared_length) / 1024 / 1024:.1f}MB)가 반환 한도(3MB)를 초과합니다. "
                                "Teams 클라이언트에서 직접 확인하거나 web_url로 메시지를 열어보세요."
                            ),
                        }

                    # content-length가 없거나(청크 전송) 부정확할 수 있으므로, 수신하면서도 한도를 재확인한다.
                    chunks = bytearray()
                    async for chunk in stream_resp.aiter_bytes():
                        chunks.extend(chunk)
                        if len(chunks) > _MAX_HOSTED_CONTENT_BYTES:
                            return {
                                "success": False,
                                "message": (
                                    "이미지 크기가 반환 한도(3MB)를 초과해 다운로드를 중단했습니다. "
                                    "Teams 클라이언트에서 직접 확인하거나 web_url로 메시지를 열어보세요."
                                ),
                            }
                    raw = bytes(chunks)
                    content_type = stream_resp.headers.get("content-type", "application/octet-stream")

            log.info("teams_channel_message_images(내용): team=%s channel=%s message=%s reply=%s hosted_content=%s size=%d",
                     team_id, channel_id, message_id, reply_id or "-", hosted_content_id, len(raw))
            return {
                "success": True,
                "team_id": team_id,
                "channel_id": channel_id,
                "message_id": reply_id or message_id,
                "hosted_content_id": hosted_content_id,
                "content_type": content_type,
                "size": len(raw),
                "content_base64": base64.b64encode(raw).decode("ascii"),
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_channel_message_images HTTP 오류: %s", e.response.text[:300])
            if e.response.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": (
                        "권한 오류(403): Azure AD 앱에 'ChannelMessage.Read.All' Delegated 권한과 "
                        "관리자 동의가 필요합니다. Azure Portal에서 권한을 추가한 뒤 ms365_logout → ms365_login으로 재인증하세요."
                    ),
                }
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": "메시지 또는 이미지를 찾을 수 없습니다."}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_channel_message_images 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R21: Teams 읽지 않은 채팅 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 판별 기준: /me/chats 응답의 viewpoint.lastMessageReadDateTime(내가 마지막으로 읽은 시각)보다
    # lastMessagePreview.createdDateTime(마지막 메시지 시각)이 늦고, 마지막 메시지 발신자가 본인이 아니면 미확인.
    # viewpoint는 채팅 목록 조회(List chats)에서만 제공된다.
    _MAX_CHAT_PAGES = 4  # /me/chats 최대 4페이지(약 200개 채팅)까지만 확인

    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_unread_chats(
        chat_type: Annotated[
            Literal["all", "oneOnOne", "group", "meeting"],
            Field(description="채팅 유형 필터. all(기본)=전체, oneOnOne=1:1, group=그룹, meeting=회의 채팅"),
        ] = "all",
        include_messages: Annotated[bool, Field(
            description="true(기본)면 채팅별 읽지 않은 메시지 본문까지 반환, false면 채팅 목록과 마지막 메시지 미리보기만 반환",
        )] = True,
        max_chats: Annotated[int, Field(ge=1, le=50, description="반환할 최대 미확인 채팅 수 (기본 20)")] = 20,
        max_messages_per_chat: Annotated[int, Field(ge=1, le=50, description="채팅당 최대 반환 메시지 수 (기본 20)")] = 20,
    ) -> dict:
        """Teams 채팅(1:1·그룹·회의 채팅) 중 읽지 않은 메시지가 있는 채팅과 그 메시지를 조회합니다.
        Teams 클라이언트의 읽음 상태(내가 마지막으로 읽은 시각)를 기준으로 판별하며, 본인이 보낸 메시지는 제외합니다.
        이 도구는 읽음 처리를 하지 않습니다(조회만 수행).
        팀 채널 메시지는 Graph API가 읽음 상태를 제공하지 않으므로 teams_channel_new_messages 도구를 사용하세요.
        Azure AD 앱에 Chat.ReadWrite(또는 Chat.Read) Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {"Authorization": f"Bearer {access_token}"}
        warnings: list[str] = []

        try:
            async with httpx.AsyncClient(timeout=60) as client:
                me_resp = await client.get(
                    "https://graph.microsoft.com/v1.0/me", params={"$select": "id"}, headers=headers,
                )
                me_resp.raise_for_status()
                my_id = me_resp.json().get("id") or ""

                # 1. 채팅 목록 + 마지막 메시지 미리보기 + 읽음 상태(viewpoint)
                chats: list[dict] = []
                next_url = "https://graph.microsoft.com/v1.0/me/chats"
                next_params: dict | None = {
                    "$expand": "lastMessagePreview",
                    "$orderby": "lastMessagePreview/createdDateTime desc",
                    "$top": 50,
                }
                pages = 0
                while next_url and pages < _MAX_CHAT_PAGES:
                    resp = await client.get(next_url, params=next_params, headers=headers)
                    resp.raise_for_status()
                    data = resp.json()
                    chats.extend(data.get("value", []))
                    next_params = None
                    next_url = data.get("@odata.nextLink", "")
                    pages += 1
                if next_url:
                    warnings.append(f"채팅이 많아 최근 활동 순 상위 {len(chats)}개 채팅만 확인했습니다.")

                # 2. 미확인 채팅 판별
                unread_chats: list[dict] = []
                no_viewpoint = 0
                for c in chats:
                    if chat_type != "all" and c.get("chatType") != chat_type:
                        continue
                    viewpoint = c.get("viewpoint")
                    if viewpoint is None:
                        no_viewpoint += 1
                        continue
                    if viewpoint.get("isHidden"):
                        continue
                    preview = c.get("lastMessagePreview") or {}
                    if not preview or preview.get("isDeleted"):
                        continue
                    _, preview_sender_id = _message_sender(preview)
                    if my_id and preview_sender_id == my_id:
                        continue  # 마지막 메시지를 본인이 보냄 → 미확인 없음
                    last_msg_at = _parse_graph_dt(preview.get("createdDateTime") or "")
                    last_read_at = _parse_graph_dt(viewpoint.get("lastMessageReadDateTime") or "")
                    if not last_msg_at:
                        continue
                    if last_read_at and last_msg_at <= last_read_at:
                        continue
                    unread_chats.append({"chat": c, "last_read_at": last_read_at})

                if no_viewpoint:
                    warnings.append(
                        f"{no_viewpoint}개 채팅은 Graph API가 읽음 상태(viewpoint)를 반환하지 않아 판별에서 제외했습니다."
                    )

                total_unread_chats = len(unread_chats)
                unread_chats = unread_chats[:max_chats]

                # 3. 채팅별 이름(멤버)과 미확인 메시지 조회
                chats_out = []
                for item in unread_chats:
                    c = item["chat"]
                    last_read_at = item["last_read_at"]
                    chat_id = c.get("id")
                    preview = c.get("lastMessagePreview") or {}
                    preview_sender, _ = _message_sender(preview)

                    members: list[str] = []
                    if not c.get("topic"):
                        try:
                            m_resp = await client.get(
                                f"https://graph.microsoft.com/v1.0/chats/{chat_id}/members",
                                headers=headers,
                            )
                            if m_resp.status_code == 200:
                                members = [
                                    m.get("displayName") or m.get("email") or ""
                                    for m in m_resp.json().get("value", [])
                                    if m.get("userId") != my_id
                                ]
                        except Exception as e:
                            log.debug("채팅 멤버 조회 실패(%s): %s", chat_id, e)

                    chat_out = {
                        "chat_id": chat_id,
                        "chat_type": c.get("chatType"),
                        "title": c.get("topic") or ", ".join(m for m in members if m) or "(제목 없음)",
                        "last_read_at": _to_kst_str(last_read_at.isoformat()) if last_read_at else None,
                        "last_message": {
                            "sender": preview_sender,
                            "created_at": _to_kst_str(preview.get("createdDateTime") or ""),
                            "preview": _html_to_text((preview.get("body") or {}).get("content") or "")[:200],
                        },
                        "web_url": c.get("webUrl"),
                    }

                    if include_messages:
                        messages: list[dict] = []
                        truncated = False
                        try:
                            msg_url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
                            msg_params: dict | None = {"$top": 50, "$orderby": "createdDateTime desc"}
                            reached_read = False
                            while msg_url and not reached_read:
                                m_resp = await client.get(msg_url, params=msg_params, headers=headers)
                                m_resp.raise_for_status()
                                m_data = m_resp.json()
                                msg_params = None
                                for msg in m_data.get("value", []):
                                    created = _parse_graph_dt(msg.get("createdDateTime") or "")
                                    if last_read_at and created and created <= last_read_at:
                                        reached_read = True
                                        break
                                    if msg.get("messageType") != "message" or msg.get("deletedDateTime"):
                                        continue
                                    sender, sender_id = _message_sender(msg)
                                    if my_id and sender_id == my_id:
                                        continue
                                    if len(messages) >= max_messages_per_chat:
                                        truncated = True
                                        reached_read = True
                                        break
                                    body = msg.get("body") or {}
                                    messages.append({
                                        "id": msg.get("id"),
                                        "sender": sender,
                                        "created_at": _to_kst_str(msg.get("createdDateTime") or ""),
                                        "body": _html_to_text(body.get("content") or "")
                                                if body.get("contentType") == "html" else (body.get("content") or ""),
                                        "attachments": len(msg.get("attachments") or []),
                                    })
                                msg_url = m_data.get("@odata.nextLink", "")
                        except httpx.HTTPStatusError as e:
                            warnings.append(f"채팅({chat_id}) 메시지 조회 실패: HTTP {e.response.status_code}")
                        messages.sort(key=lambda m: m["created_at"])
                        chat_out["unread_count"] = len(messages)
                        chat_out["unread_count_truncated"] = truncated
                        chat_out["messages"] = messages

                    chats_out.append(chat_out)

            log.info("teams_unread_chats: 확인 %d개, 미확인 %d개", len(chats), total_unread_chats)
            return {
                "success": True,
                "chat_type": chat_type,
                "checked_chats": len(chats),
                "unread_chat_count": total_unread_chats,
                "returned_chat_count": len(chats_out),
                "chats": chats_out,
                "note": "조회만 했으며 읽음 처리는 하지 않았습니다. 시각은 KST 기준입니다.",
                **({"warnings": warnings} if warnings else {}),
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_unread_chats HTTP 오류: %s", e.response.text[:300])
            if e.response.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": "권한 오류(403): Azure AD 앱에 Chat.ReadWrite(또는 Chat.Read) Delegated 권한이 필요합니다.",
                }
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_unread_chats 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool R22: Teams 팀 채널 새 메시지 조회 (기준 시각 이후)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Graph API는 채널 메시지의 사용자별 읽음 상태를 제공하지 않는다.
    # 따라서 기준 시각(since) 이후 작성된 새 글·새 댓글을 "새 메시지"로 간주한다.
    # 채널 메시지 delta 조회의 $filter=lastModifiedDateTime gt {since}로 변경분만 받은 뒤
    # 루트 메시지(새 글)와 replies(새 댓글)를 createdDateTime 기준으로 다시 거른다.
    _CHANNEL_CONCURRENCY = 5

    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def teams_channel_new_messages(
        since: Annotated[str, Field(
            pattern=r"^(\d{4}-\d{2}-\d{2}( \d{2}:\d{2})?)?$",
            description="기준 시각(KST) 'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM'. 이 시각 이후 작성된 글·댓글을 반환. 비우면 최근 24시간",
        )] = "",
        team_name: Annotated[str, Field(description="팀 이름 필터(부분 일치). 비우면 가입한 전체 팀")] = "",
        channel_name: Annotated[str, Field(description="채널 이름 필터(부분 일치). 비우면 전체 채널")] = "",
        include_replies: Annotated[bool, Field(description="true(기본)면 기준 시각 이후 달린 댓글도 포함")] = True,
        max_channels: Annotated[int, Field(ge=1, le=50, description="확인할 최대 채널 수 (기본 30)")] = 30,
        limit_per_channel: Annotated[int, Field(ge=1, le=50, description="채널당 최대 반환 메시지 수 (기본 20)")] = 20,
    ) -> dict:
        """가입한 Teams 팀 채널에서 기준 시각 이후 새로 올라온 글과 댓글을 조회합니다.
        Graph API는 채널 메시지의 읽음/안읽음 상태를 제공하지 않으므로, Teams 화면의 '읽지 않음' 표시와 정확히 일치하지 않습니다.
        '마지막으로 확인한 시각'을 since로 넘겨 그 이후의 새 메시지를 확인하는 용도로 사용하세요. 본인이 작성한 글·댓글은 제외합니다.
        Azure AD 앱에 ChannelMessage.Read.All Delegated 권한 및 관리자 동의가 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        import asyncio
        import httpx
        from datetime import datetime, timedelta, timezone
        from zoneinfo import ZoneInfo

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        KST = ZoneInfo("Asia/Seoul")
        if since:
            fmt = "%Y-%m-%d %H:%M" if " " in since else "%Y-%m-%d"
            try:
                since_dt = datetime.strptime(since, fmt).replace(tzinfo=KST)
            except ValueError:
                return {"success": False, "message": f"since 형식이 올바르지 않습니다: {since}"}
        else:
            since_dt = datetime.now(KST) - timedelta(hours=24)
        since_utc = since_dt.astimezone(timezone.utc)
        if since_utc > datetime.now(timezone.utc):
            return {"success": False, "message": "since가 현재 시각보다 미래입니다."}
        since_filter = since_utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

        headers = {"Authorization": f"Bearer {access_token}"}
        warnings: list[str] = []

        try:
            async with httpx.AsyncClient(timeout=60) as client:
                me_resp = await client.get(
                    "https://graph.microsoft.com/v1.0/me", params={"$select": "id"}, headers=headers,
                )
                me_resp.raise_for_status()
                my_id = me_resp.json().get("id") or ""

                # 1. 대상 팀·채널 수집
                teams_resp = await client.get("https://graph.microsoft.com/v1.0/me/joinedTeams", headers=headers)
                teams_resp.raise_for_status()
                teams = [
                    t for t in teams_resp.json().get("value", [])
                    if not team_name or team_name.lower() in (t.get("displayName") or "").lower()
                ]
                if not teams:
                    return {"success": True, "since": since_dt.strftime("%Y-%m-%d %H:%M"), "channel_count": 0,
                            "channels": [], "message": "조건에 맞는 가입 팀이 없습니다."}

                targets: list[dict] = []
                for t in teams:
                    try:
                        ch_resp = await client.get(
                            f"https://graph.microsoft.com/v1.0/teams/{t['id']}/channels",
                            params={"$select": "id,displayName"},
                            headers=headers,
                        )
                        ch_resp.raise_for_status()
                        for ch in ch_resp.json().get("value", []):
                            if channel_name and channel_name.lower() not in (ch.get("displayName") or "").lower():
                                continue
                            targets.append({
                                "team_id": t["id"], "team_name": t.get("displayName"),
                                "channel_id": ch.get("id"), "channel_name": ch.get("displayName"),
                            })
                    except httpx.HTTPStatusError as e:
                        warnings.append(f"팀 '{t.get('displayName')}' 채널 목록 조회 실패: HTTP {e.response.status_code}")

                total_channels = len(targets)
                if total_channels > max_channels:
                    warnings.append(
                        f"대상 채널이 {total_channels}개라 앞의 {max_channels}개만 확인했습니다. "
                        "team_name/channel_name으로 범위를 좁히거나 max_channels를 늘리세요."
                    )
                    targets = targets[:max_channels]

                # 2. 채널별 변경분 조회 (delta + lastModifiedDateTime 필터)
                sem = asyncio.Semaphore(_CHANNEL_CONCURRENCY)

                async def _fetch_channel(tg: dict) -> dict:
                    async with sem:
                        url = (f"https://graph.microsoft.com/v1.0/teams/{tg['team_id']}"
                               f"/channels/{tg['channel_id']}/messages/delta")
                        params: dict | None = {
                            "$filter": f"lastModifiedDateTime gt {since_filter}",
                            "$top": 50,
                        }
                        if include_replies:
                            params["$expand"] = "replies"
                        items: list[dict] = []
                        try:
                            while url:
                                resp = await client.get(url, params=params, headers=headers)
                                resp.raise_for_status()
                                data = resp.json()
                                params = None
                                for msg in data.get("value", []):
                                    if msg.get("messageType") != "message":
                                        continue
                                    root_created = _parse_graph_dt(msg.get("createdDateTime") or "")
                                    root_sender, root_sender_id = _message_sender(msg)
                                    root_text = _html_to_text((msg.get("body") or {}).get("content") or "")
                                    if (root_created and root_created > since_utc and not msg.get("deletedDateTime")
                                            and not (my_id and root_sender_id == my_id)):
                                        items.append({
                                            "kind": "post",
                                            "message_id": msg.get("id"),
                                            "subject": msg.get("subject"),
                                            "sender": root_sender,
                                            "created_at_utc": msg.get("createdDateTime") or "",
                                            "created_at": _to_kst_str(msg.get("createdDateTime") or ""),
                                            "body": root_text,
                                            "attachments": len(msg.get("attachments") or []),
                                            "web_url": msg.get("webUrl"),
                                        })
                                    for rep in (msg.get("replies") or []) if include_replies else []:
                                        if rep.get("messageType") != "message" or rep.get("deletedDateTime"):
                                            continue
                                        rep_created = _parse_graph_dt(rep.get("createdDateTime") or "")
                                        rep_sender, rep_sender_id = _message_sender(rep)
                                        if not rep_created or rep_created <= since_utc:
                                            continue
                                        if my_id and rep_sender_id == my_id:
                                            continue
                                        items.append({
                                            "kind": "reply",
                                            "message_id": msg.get("id"),
                                            "reply_id": rep.get("id"),
                                            "parent_preview": (msg.get("subject") or root_text)[:80],
                                            "sender": rep_sender,
                                            "created_at_utc": rep.get("createdDateTime") or "",
                                            "created_at": _to_kst_str(rep.get("createdDateTime") or ""),
                                            "body": _html_to_text((rep.get("body") or {}).get("content") or ""),
                                            "attachments": len(rep.get("attachments") or []),
                                            "web_url": rep.get("webUrl"),
                                        })
                                url = data.get("@odata.nextLink", "")  # 마지막 페이지는 deltaLink만 반환
                        except httpx.HTTPStatusError as e:
                            return {**tg, "error_status": e.response.status_code}
                        except Exception as e:
                            return {**tg, "error": str(e)}
                        items.sort(key=lambda m: m["created_at_utc"])
                        for it in items:
                            it.pop("created_at_utc", None)
                        return {**tg, "new_count": len(items), "truncated": len(items) > limit_per_channel,
                                "messages": items[-limit_per_channel:]}

                results = await asyncio.gather(*(_fetch_channel(tg) for tg in targets))

            forbidden = [r for r in results if r.get("error_status") == 403]
            if results and len(forbidden) == len(results):
                return {
                    "success": False,
                    "status_code": 403,
                    "message": (
                        "권한 오류(403): Azure AD 앱에 'ChannelMessage.Read.All' Delegated 권한과 "
                        "관리자 동의가 필요합니다. Azure Portal에서 권한을 추가한 뒤 ms365_logout → ms365_login으로 재인증하세요."
                    ),
                }
            channels_out = []
            for r in results:
                if r.get("error_status") or r.get("error"):
                    warnings.append(
                        f"'{r['team_name']} > {r['channel_name']}' 조회 실패: "
                        f"{('HTTP ' + str(r['error_status'])) if r.get('error_status') else r.get('error')}"
                    )
                    continue
                if r["new_count"]:
                    channels_out.append(r)
            channels_out.sort(key=lambda r: r["new_count"], reverse=True)

            log.info("teams_channel_new_messages: 채널 %d개 확인, 새 메시지 채널 %d개", len(results), len(channels_out))
            return {
                "success": True,
                "since": since_dt.strftime("%Y-%m-%d %H:%M"),
                "checked_channel_count": len(results),
                "channel_with_new_count": len(channels_out),
                "total_new_messages": sum(r["new_count"] for r in channels_out),
                "channels": channels_out,
                "note": ("채널 메시지는 Graph API가 읽음 상태를 제공하지 않아 기준 시각 이후 작성된 글·댓글을 반환합니다. "
                         "시각은 KST 기준입니다."),
                **({"warnings": warnings} if warnings else {}),
            }
        except httpx.HTTPStatusError as e:
            log.error("teams_channel_new_messages HTTP 오류: %s", e.response.text[:300])
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("teams_channel_new_messages 실패: %s", e)
            return {"success": False, "message": str(e)}

    log.info(
        "MS365 Graph API 읽기 도구 22개 등록 완료 "
        "(채팅검색/회의가용/캘린더/이메일/리소스/SP파일/SP폴더/임직원조회/임직원상세/Teams채널조회/그룹채팅/채널메시지/이메일첨부/이메일첨부일괄/"
        "Excel워크시트목록/Excel범위읽기/사용량보고서/채널댓글조회/채널첨부파일조회/채널이미지조회/미확인채팅/채널새메시지)"
    )
