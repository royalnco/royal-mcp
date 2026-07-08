# ms365/graph.py — MS365 Graph API 쓰기 도구 모듈
# ──────────────────────────────────────────────────────────────
# 필요 패키지: pip install msgraph-sdk azure-identity msal-extensions
#
# 인증 방식: DeviceCodeCredential (Delegated 권한, 사용자 본인 로그인)
# 인증 모듈: ms365/auth.py (서버 시작 시 초기화)
#
# [Delegated 권한 필요 목록]
#   ChannelMessage.Send   — Teams 채널 메시지 발송
#   Chat.ReadWrite        — Teams 1:1 채팅 발송/읽기
#   Sites.ReadWrite.All   — SharePoint 파일 생성/수정
#   Files.ReadWrite.All   — OneDrive 파일 쓰기
#   Mail.Send             — 이메일 발송 (로그인한 본인 계정으로 발송)
#   Mail.ReadWrite        — 이메일 범주(Category) 설정
#   Calendars.ReadWrite   — 캘린더 일정 생성/수정
#   MailboxSettings.ReadWrite — Outlook 범주(masterCategories) 관리
#
# ──────────────────────────────────────────────────────────────

import logging
from typing import Annotated, List, Optional

from pydantic import Field

log = logging.getLogger("royal-mcp.ms365")

_VALID_CATEGORY_COLORS = {
    "none", "preset0", "preset1", "preset2", "preset3", "preset4",
    "preset5", "preset6", "preset7", "preset8", "preset9", "preset10",
    "preset11", "preset12", "preset13", "preset14", "preset15", "preset16",
    "preset17", "preset18", "preset19", "preset20", "preset21", "preset22",
    "preset23", "preset24",
}


def register_ms365_tools(mcp):
    """MS365 Graph API 쓰기 도구를 MCP 서버에 등록합니다.

    도구는 항상 등록되지만, 각 도구 호출 시 로그인 여부를 확인합니다.
    미로그인 시 ms365_login 도구를 먼저 호출하도록 안내합니다.
    """
    from ms365.auth import get_graph_client, get_user_email, fetch_user_profile, require_auth

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-1: Teams 채널 메시지 발송
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def send_teams_channel_message(
        team_id: Annotated[str, Field(description="Teams 팀 ID (chat_message_search로 확인)")],
        channel_id: Annotated[str, Field(description="Teams 채널 ID")],
        message: Annotated[str, Field(min_length=1, description="전송할 메시지 본문")],
    ) -> dict:
        """Teams 채널에 메시지를 전송합니다.
        team_id와 channel_id는 MS365 read MCP의 chat_message_search 도구로 미리 확인하세요.
        Azure AD 앱에 ChannelMessage.Send Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        try:
            from msgraph.generated.models.chat_message import ChatMessage
            from msgraph.generated.models.item_body import ItemBody

            body = ChatMessage()
            body.body = ItemBody()
            body.body.content = message

            result = await (
                graph.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.post(body)
            )
            log.info("send_teams_channel_message: team=%s channel=%s id=%s (by %s)",
                     team_id, channel_id, result.id, get_user_email())
            return {
                "success": True,
                "message_id": result.id,
                "team_id": team_id,
                "channel_id": channel_id,
                "sent_by": get_user_email(),
            }
        except Exception as e:
            log.error("send_teams_channel_message 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-2: SharePoint 텍스트 파일 생성
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def create_sharepoint_file(
        site_id: Annotated[str, Field(description="SharePoint 사이트 ID (sharepoint_search로 확인)")],
        folder_path: Annotated[str, Field(description="업로드 폴더 경로 (예: Documents/Reports)")],
        filename: Annotated[str, Field(description="파일명 (예: report.txt)")],
        content: Annotated[str, Field(description="파일 내용 (UTF-8 텍스트)")],
    ) -> dict:
        """SharePoint에 텍스트 파일을 생성합니다. 같은 경로에 파일이 존재하면 덮어씁니다.
        site_id는 MS365 read MCP의 sharepoint_search 도구로 미리 확인하세요.
        Azure AD 앱에 Sites.ReadWrite.All Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        try:
            content_bytes = content.encode("utf-8")
            path = f"{folder_path}/{filename}"

            drive = await graph.sites.by_site_id(site_id).drive.get()
            item_path_id = f"root:/{folder_path}/{filename}:"
            result = await (
                graph.drives.by_drive_id(drive.id)
                .items.by_drive_item_id(item_path_id)
                .content.put(content_bytes)
            )
            log.info("create_sharepoint_file: site=%s path=%s (by %s)", site_id, path, get_user_email())
            return {
                "success": True,
                "filename": filename,
                "path": path,
                "site_id": site_id,
                "item_id": result.id if result else None,
                "created_by": get_user_email(),
            }
        except Exception as e:
            log.error("create_sharepoint_file 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-3: Outlook 이메일 발송
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def send_email(
        to: Annotated[str, Field(description="수신자 이메일 주소")],
        subject: Annotated[str, Field(min_length=1, description="이메일 제목")],
        body_text: Annotated[str, Field(description="이메일 본문 (텍스트)")],
    ) -> dict:
        """Outlook으로 이메일을 발송합니다.
        발신자는 현재 로그인한 본인 계정입니다.
        Azure AD 앱에 Mail.Send Delegated 권한이 필요합니다.
        발송 전 반드시 사용자에게 내용(수신자/제목/본문)을 확인받으세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        try:
            from msgraph.generated.models.message import Message
            from msgraph.generated.models.recipient import Recipient
            from msgraph.generated.models.email_address import EmailAddress
            from msgraph.generated.models.item_body import ItemBody
            from msgraph.generated.models.body_type import BodyType
            from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
                SendMailPostRequestBody,
            )

            email_addr = EmailAddress()
            email_addr.address = to

            recip = Recipient()
            recip.email_address = email_addr

            item_body = ItemBody()
            item_body.content_type = BodyType.Text
            item_body.content = body_text

            msg = Message()
            msg.subject = subject
            msg.body = item_body
            msg.to_recipients = [recip]

            req_body = SendMailPostRequestBody()
            req_body.message = msg

            # graph.me — 로그인한 본인 계정으로 발송
            await graph.me.send_mail.post(req_body)
            sender = get_user_email()
            log.info("send_email: from=%s to=%s subject=%s", sender, to, subject)
            return {"success": True, "from": sender, "to": to, "subject": subject}
        except Exception as e:
            log.error("send_email 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-4: 캘린더 일정 생성
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def create_calendar_event(
        title: Annotated[str, Field(min_length=1, description="일정 제목")],
        start: Annotated[str, Field(pattern=r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$", description="시작 시간 ISO8601 (예: 2026-04-23T09:00:00)")],
        end: Annotated[str, Field(pattern=r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$", description="종료 시간 ISO8601 (예: 2026-04-23T10:00:00)")],
        attendees: Annotated[List[str], Field(description="참석자 이메일 목록 (예: [\"user@royal.co.kr\"])")],
        is_online_meeting: Annotated[bool, Field(description="true: Teams 온라인 회의 링크 자동 생성")] = False,
    ) -> dict:
        """Outlook 캘린더에 일정을 생성합니다. 시간대는 KST(Asia/Seoul) 기준입니다.
        로그인한 본인 캘린더에 일정이 생성됩니다.
        is_online_meeting=true 시 Teams 회의 링크(join_url)가 자동 생성됩니다.
        Azure AD 앱에 Calendars.ReadWrite Delegated 권한이 필요합니다.
        생성 전 반드시 사용자에게 일정 내용을 확인받으세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        try:
            from msgraph.generated.models.event import Event
            from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
            from msgraph.generated.models.attendee import Attendee
            from msgraph.generated.models.email_address import EmailAddress

            event_start = DateTimeTimeZone()
            event_start.date_time = start
            event_start.time_zone = "Asia/Seoul"

            event_end = DateTimeTimeZone()
            event_end.date_time = end
            event_end.time_zone = "Asia/Seoul"

            attendee_list = []
            for addr in attendees:
                ea = EmailAddress()
                ea.address = addr
                att = Attendee()
                att.email_address = ea
                attendee_list.append(att)

            event = Event()
            event.subject = title
            event.start = event_start
            event.end = event_end
            event.attendees = attendee_list

            if is_online_meeting:
                from msgraph.generated.models.online_meeting_provider_type import OnlineMeetingProviderType
                event.is_online_meeting = True
                event.online_meeting_provider = OnlineMeetingProviderType.teams_for_business

            # graph.me — 로그인한 본인 캘린더에 생성
            result = await graph.me.events.post(event)
            log.info("create_calendar_event: id=%s subject=%s online=%s (by %s)",
                     result.id, result.subject, is_online_meeting, get_user_email())

            join_url = None
            if is_online_meeting and result.online_meeting:
                join_url = result.online_meeting.join_url

            return {
                "success": True,
                "event_id": result.id,
                "subject": result.subject,
                "start": start,
                "end": end,
                "attendees": attendees,
                "is_online_meeting": is_online_meeting,
                **({"join_url": join_url} if join_url else {}),
                "created_by": get_user_email(),
            }
        except Exception as e:
            log.error("create_calendar_event 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-5: Teams 1:1 채팅 메시지 발송
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def send_teams_chat_message(
        recipient_email: Annotated[str, Field(min_length=1, description="수신자 이메일 주소 (예: 'user@iroyal.kr'). org_users_search로 미리 확인할 수 있습니다.")],
        message: Annotated[str, Field(min_length=1, description="전송할 메시지 본문 (텍스트)")],
    ) -> dict:
        """Teams 1:1 채팅으로 메시지를 전송합니다.

        내부적으로 수신자와의 1:1 채팅방을 자동으로 생성(또는 기존 채팅방 재사용)한 뒤
        메시지를 발송합니다. 수신자는 Teams 채팅 탭에서 메시지를 확인할 수 있습니다.

        사용 예시:
          - recipient_email: "hong@iroyal.kr"
          - message: "회의 자료 검토 부탁드립니다."

        Azure AD 앱에 Chat.ReadWrite Delegated 권한이 필요합니다.
        발송 전 반드시 사용자에게 수신자와 메시지 내용을 확인받으세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                # ── Step 1: 로그인 사용자 ID 가져오기 ──────────────
                me_resp = await client.get(
                    "https://graph.microsoft.com/v1.0/me",
                    params={"$select": "id,displayName,mail,userPrincipalName"},
                    headers=headers,
                )
                me_resp.raise_for_status()
                me_data = me_resp.json()
                my_id = me_data["id"]
                my_email = me_data.get("mail") or me_data.get("userPrincipalName")

                # ── Step 2: 수신자 사용자 ID 조회 ─────────────────
                user_resp = await client.get(
                    f"https://graph.microsoft.com/v1.0/users/{recipient_email}",
                    params={"$select": "id,displayName,mail,userPrincipalName"},
                    headers=headers,
                )
                user_resp.raise_for_status()
                recipient_data = user_resp.json()
                recipient_id = recipient_data["id"]
                recipient_name = recipient_data.get("displayName", recipient_email)

                # ── Step 3: 1:1 채팅방 생성 (이미 존재하면 기존 채팅방 반환) ──
                create_chat_body = {
                    "chatType": "oneOnOne",
                    "members": [
                        {
                            "@odata.type": "#microsoft.graph.aadUserConversationMember",
                            "roles": ["owner"],
                            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
                        },
                        {
                            "@odata.type": "#microsoft.graph.aadUserConversationMember",
                            "roles": ["owner"],
                            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{recipient_id}')",
                        },
                    ],
                }
                chat_resp = await client.post(
                    "https://graph.microsoft.com/v1.0/chats",
                    json=create_chat_body,
                    headers=headers,
                )
                chat_resp.raise_for_status()
                chat_data = chat_resp.json()
                chat_id = chat_data["id"]

                # ── Step 4: 메시지 발송 ───────────────────────────
                msg_body = {
                    "body": {
                        "content": message,
                    }
                }
                msg_resp = await client.post(
                    f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages",
                    json=msg_body,
                    headers=headers,
                )
                msg_resp.raise_for_status()
                msg_data = msg_resp.json()

            log.info(
                "send_teams_chat_message: from=%s to=%s (%s) chat=%s msg=%s",
                my_email, recipient_email, recipient_name, chat_id, msg_data.get("id"),
            )
            return {
                "success": True,
                "chat_id": chat_id,
                "message_id": msg_data.get("id"),
                "from": my_email,
                "to": recipient_email,
                "to_name": recipient_name,
                "message_preview": message[:100],
            }
        except httpx.HTTPStatusError as e:
            error_body = e.response.text
            log.error("send_teams_chat_message HTTP 오류: %s — %s", e.response.status_code, error_body)
            return {
                "success": False,
                "status_code": e.response.status_code,
                "message": f"Graph API 오류: {error_body}",
            }
        except Exception as e:
            log.error("send_teams_chat_message 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-6: Outlook 이메일 범주(Category) 설정
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def outlook_email_update_category(
        message_id: Annotated[str, Field(min_length=1, description="범주를 설정할 이메일 ID (outlook_email_search 반환 'id' 값)")],
        categories: Annotated[List[str], Field(description="설정할 범주 이름 배열 (예: [\"Red Category\", \"Important\"]). 빈 배열 전달 시 범주 전체 제거")],
        mailbox_owner: Annotated[str, Field(description="메일박스 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 이메일에 범주(Category)를 설정합니다.
        outlook_email_search 도구로 message_id를 먼저 확인하세요.
        빈 배열([])을 전달하면 기존 범주가 모두 제거됩니다.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
        )

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(
                    url,
                    json={"categories": categories},
                    headers=headers,
                )
                resp.raise_for_status()
                data = resp.json()

            log.info(
                "outlook_email_update_category: message=%s categories=%s (by %s)",
                message_id, categories, get_user_email(),
            )
            return {
                "success": True,
                "message_id": message_id,
                "categories": data.get("categories", categories),
                "subject": data.get("subject"),
                "updated_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            log.error("outlook_email_update_category HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            if e.response.status_code == 403:
                return {
                    "success": False,
                    "status_code": 403,
                    "message": "권한 오류(403): Azure AD 앱에 'Mail.ReadWrite' Delegated 권한이 필요합니다.",
                }
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("outlook_email_update_category 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-7: 캘린더 일정 수정
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def update_calendar_event(
        event_id: Annotated[str, Field(min_length=1, description="수정할 일정 ID (outlook_calendar_search 반환 'id' 값)")],
        title: Annotated[str, Field(description="변경할 일정 제목. 비우면 현재 제목 유지")] = "",
        start: Annotated[str, Field(
            pattern=r"^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})?$",
            description="변경할 시작 시간 ISO8601 KST (예: 2026-05-20T09:00:00). 비우면 현재 시간 유지",
        )] = "",
        end: Annotated[str, Field(
            pattern=r"^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})?$",
            description="변경할 종료 시간 ISO8601 KST (예: 2026-05-20T10:00:00). 비우면 현재 시간 유지",
        )] = "",
        attendees: Annotated[Optional[List[str]], Field(description="변경할 참석자 이메일 목록. 지정 시 기존 참석자를 전체 교체. null이면 참석자 변경 없음")] = None,
        body_text: Annotated[str, Field(description="변경할 일정 본문(설명). 비우면 현재 본문 유지")] = "",
        is_online_meeting: Annotated[Optional[bool], Field(description="Teams 온라인 회의 링크 추가 여부. true: 링크 생성. ※ 이미 생성된 온라인 회의는 false로 해제 불가(Graph API 제약). null: 변경 없음")] = None,
        calendar_owner: Annotated[str, Field(description="캘린더 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 캘린더 일정을 수정합니다. 지정한 필드만 변경되며 나머지는 유지됩니다.
        event_id는 outlook_calendar_search 도구로 먼저 확인하세요.
        attendees를 지정하면 기존 참석자 목록이 전체 교체됩니다.
        is_online_meeting=true 시 Teams 회의 링크(join_url)가 생성됩니다.
        Azure AD 앱에 Calendars.ReadWrite Delegated 권한이 필요합니다.
        수정 전 반드시 사용자에게 변경 내용을 확인받으세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = calendar_owner or get_user_email()

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/events/{event_id}"
            if calendar_owner
            else f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
        )

        # 지정된 필드만 PATCH body에 포함
        patch_body: dict = {}
        if title:
            patch_body["subject"] = title
        if start:
            patch_body["start"] = {"dateTime": start, "timeZone": "Asia/Seoul"}
        if end:
            patch_body["end"] = {"dateTime": end, "timeZone": "Asia/Seoul"}
        if attendees is not None:
            patch_body["attendees"] = [
                {"emailAddress": {"address": addr}, "type": "required"}
                for addr in attendees
            ]
        if body_text:
            patch_body["body"] = {"contentType": "text", "content": body_text}
        if is_online_meeting is True:
            patch_body["isOnlineMeeting"] = True
            patch_body["onlineMeetingProvider"] = "teamsForBusiness"
        elif is_online_meeting is False:
            return {"success": False, "message": "온라인 회의는 한번 생성되면 해제할 수 없습니다 (Graph API 제약)."}


        if not patch_body:
            return {"success": False, "message": "변경할 항목이 없습니다. title·start·end·attendees·body_text·is_online_meeting 중 하나 이상을 지정하세요."}

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(url, json=patch_body, headers=headers)
                resp.raise_for_status()
                data = resp.json()

            log.info(
                "update_calendar_event: event=%s fields=%s (by %s)",
                event_id, list(patch_body.keys()), get_user_email(),
            )
            join_url = (data.get("onlineMeeting") or {}).get("joinUrl")
            return {
                "success": True,
                "event_id": event_id,
                "subject": data.get("subject"),
                "start": (data.get("start") or {}).get("dateTime"),
                "end": (data.get("end") or {}).get("dateTime"),
                "updated_fields": list(patch_body.keys()),
                **({"join_url": join_url} if join_url else {}),
                "updated_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            log.error("update_calendar_event HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("update_calendar_event 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-8: 캘린더 일정 삭제
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "idempotentHint": True})
    async def delete_calendar_event(
        event_id: Annotated[str, Field(min_length=1, description="삭제할 일정 ID (outlook_calendar_search 반환 'id' 값)")],
        calendar_owner: Annotated[str, Field(description="캘린더 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 캘린더 일정을 삭제합니다. 삭제된 일정은 복구할 수 없습니다.
        event_id는 outlook_calendar_search 도구로 먼저 확인하세요.
        Azure AD 앱에 Calendars.ReadWrite Delegated 권한이 필요합니다.
        삭제 전 반드시 사용자에게 대상 일정을 확인받으세요.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = calendar_owner or get_user_email()

        headers = {"Authorization": f"Bearer {access_token}"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/events/{event_id}"
            if calendar_owner
            else f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
        )

        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.delete(url, headers=headers)
                resp.raise_for_status()  # 204 No Content

            log.info("delete_calendar_event: event=%s (by %s)", event_id, get_user_email())
            return {
                "success": True,
                "event_id": event_id,
                "deleted_by": get_user_email(),
                "message": "일정이 삭제되었습니다.",
            }
        except httpx.HTTPStatusError as e:
            log.error("delete_calendar_event HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"일정을 찾을 수 없습니다: {event_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("delete_calendar_event 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-9: 이메일 답장 (즉시 발송)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def outlook_email_reply(
        message_id: Annotated[str, Field(min_length=1, description="답장할 메일 ID (outlook_email_search 반환 'id' 값)")],
        comment: Annotated[str, Field(min_length=1, description="답장 본문")],
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일에 즉시 답장(Reply)합니다. 발신자에게만 전송됩니다.
        message_id는 outlook_email_search로 먼저 확인하세요.
        발송 전 반드시 사용자에게 내용을 확인받으세요.
        Azure AD 앱에 Mail.Send Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}/reply"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/reply"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(url, json={"comment": comment}, headers=headers)
                resp.raise_for_status()
            log.info("outlook_email_reply: msg=%s (by %s)", message_id, get_user_email())
            return {"success": True, "message_id": message_id, "sent_by": get_user_email()}
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-10: 이메일 전체 답장 (즉시 발송)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def outlook_email_reply_all(
        message_id: Annotated[str, Field(min_length=1, description="전체 답장할 메일 ID (outlook_email_search 반환 'id' 값)")],
        comment: Annotated[str, Field(min_length=1, description="답장 본문")],
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일에 전체 답장(Reply All)합니다. 발신자 및 모든 수신자에게 전송됩니다.
        message_id는 outlook_email_search로 먼저 확인하세요.
        발송 전 반드시 사용자에게 수신자 목록과 내용을 확인받으세요.
        Azure AD 앱에 Mail.Send Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}/replyAll"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/replyAll"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(url, json={"comment": comment}, headers=headers)
                resp.raise_for_status()
            log.info("outlook_email_reply_all: msg=%s (by %s)", message_id, get_user_email())
            return {"success": True, "message_id": message_id, "sent_by": get_user_email()}
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-11: 이메일 전달
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def outlook_email_forward(
        message_id: Annotated[str, Field(min_length=1, description="전달할 메일 ID (outlook_email_search 반환 'id' 값)")],
        to_addresses: Annotated[List[str], Field(description="전달 대상 이메일 목록 (예: [\"user@iroyal.kr\"])")],
        comment: Annotated[str, Field(description="전달 시 추가할 메모 (생략 가능)")] = "",
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일을 다른 수신자에게 전달(Forward)합니다.
        message_id는 outlook_email_search로 먼저 확인하세요.
        발송 전 반드시 사용자에게 수신자와 내용을 확인받으세요.
        Azure AD 앱에 Mail.Send Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}/forward"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/forward"
        )
        body: dict = {
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_addresses],
        }
        if comment:
            body["comment"] = comment
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(url, json=body, headers=headers)
                resp.raise_for_status()
            log.info("outlook_email_forward: msg=%s to=%s (by %s)", message_id, to_addresses, get_user_email())
            return {"success": True, "message_id": message_id, "forwarded_to": to_addresses, "sent_by": get_user_email()}
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-12: 이메일 폴더 이동 / 아카이브
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def outlook_email_move(
        message_id: Annotated[str, Field(min_length=1, description="이동할 메일 ID (outlook_email_search 반환 'id' 값)")],
        destination_folder: Annotated[str, Field(
            min_length=1,
            description="이동 대상 폴더. 잘 알려진 이름: inbox, archive, drafts, sentitems, deleteditems, junkemail. 또는 폴더 ID 직접 입력.",
        )],
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일을 지정한 폴더로 이동합니다. archive로 이동하면 보관처리됩니다.
        message_id는 outlook_email_search로 먼저 확인하세요.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}/move"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/move"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(url, json={"destinationId": destination_folder}, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            log.info("outlook_email_move: msg=%s → %s (by %s)", message_id, destination_folder, get_user_email())
            return {
                "success": True,
                "new_message_id": data.get("id"),
                "destination_folder": destination_folder,
                "moved_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-13: 이메일 삭제 (휴지통 이동)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "idempotentHint": True})
    async def outlook_email_delete(
        message_id: Annotated[str, Field(min_length=1, description="삭제할 메일 ID (outlook_email_search 반환 'id' 값)")],
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일을 삭제합니다 (휴지통으로 이동). Outlook 휴지통에서 복구할 수 있습니다.
        message_id는 outlook_email_search로 먼저 확인하세요.
        삭제 전 반드시 사용자에게 확인받으세요.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.delete(url, headers=headers)
                resp.raise_for_status()
            log.info("outlook_email_delete: msg=%s (by %s)", message_id, get_user_email())
            return {"success": True, "message_id": message_id, "deleted_by": get_user_email(), "message": "이메일이 휴지통으로 이동되었습니다."}
        except httpx.HTTPStatusError as e:
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"메일을 찾을 수 없습니다: {message_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-14: 이메일 읽음/안읽음 처리
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def outlook_email_mark_read(
        message_id: Annotated[str, Field(min_length=1, description="처리할 메일 ID (outlook_email_search 반환 'id' 값)")],
        is_read: Annotated[bool, Field(description="true: 읽음 처리, false: 안읽음으로 변경")] = True,
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일의 읽음/안읽음 상태를 변경합니다.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(url, json={"isRead": is_read}, headers=headers)
                resp.raise_for_status()
            log.info("outlook_email_mark_read: msg=%s isRead=%s (by %s)", message_id, is_read, get_user_email())
            return {"success": True, "message_id": message_id, "is_read": is_read}
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-15: 이메일 플래그(팔로우업) 설정
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def outlook_email_set_flag(
        message_id: Annotated[str, Field(min_length=1, description="처리할 메일 ID (outlook_email_search 반환 'id' 값)")],
        flag_status: Annotated[str, Field(description="플래그 상태: notFlagged(플래그 해제), flagged(팔로우업 필요), complete(완료)")],
        mailbox_owner: Annotated[str, Field(description="메일함 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """이메일에 팔로우업 플래그를 설정합니다.
        flag_status: notFlagged(해제) / flagged(팔로우업 필요) / complete(처리 완료)
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        valid = {"notFlagged", "flagged", "complete"}
        if flag_status not in valid:
            return {"success": False, "message": f"flag_status는 {valid} 중 하나여야 합니다."}

        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/messages/{message_id}"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(url, json={"flag": {"flagStatus": flag_status}}, headers=headers)
                resp.raise_for_status()
            log.info("outlook_email_set_flag: msg=%s flag=%s (by %s)", message_id, flag_status, get_user_email())
            return {"success": True, "message_id": message_id, "flag_status": flag_status}
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-16: OneDrive/SharePoint 파일 업로드 (바이너리 + 대용량)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def onedrive_upload_file(
        filename: Annotated[str, Field(min_length=1, description="업로드할 파일명 (예: report.xlsx)")],
        content: Annotated[str, Field(description="파일 내용. encoding=utf-8이면 텍스트, encoding=base64이면 Base64 인코딩 바이트")],
        folder_path: Annotated[str, Field(description="업로드 폴더 경로 (예: Documents/Reports). 루트이면 빈 문자열")] = "",
        encoding: Annotated[str, Field(description="콘텐츠 인코딩: utf-8(텍스트 파일) 또는 base64(바이너리 파일)")] = "utf-8",
        site_id: Annotated[str, Field(description="SharePoint 사이트 ID (sharepoint_search 반환값). site_id/drive_id 중 하나 필수")] = "",
        drive_id: Annotated[str, Field(description="OneDrive/SharePoint 드라이브 ID (sharepoint_search 반환 drive_id). site_id/drive_id 중 하나 필수")] = "",
    ) -> dict:
        """OneDrive 또는 SharePoint에 파일을 업로드합니다.
        텍스트·바이너리 모두 지원하며, 4MB 초과 파일은 업로드 세션(청크 전송)을 자동 사용합니다.
        같은 경로에 파일이 있으면 덮어씁니다.
        site_id 또는 drive_id 중 하나를 반드시 지정하세요.
        Azure AD 앱에 Files.ReadWrite.All 또는 Sites.ReadWrite.All 권한이 필요합니다.
        """
        if not site_id and not drive_id:
            return {"success": False, "message": "site_id 또는 drive_id 중 하나를 지정하세요."}

        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx, base64

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

        try:
            if encoding == "base64":
                content_bytes = base64.b64decode(content)
            else:
                content_bytes = content.encode("utf-8")
        except Exception as e:
            return {"success": False, "message": f"콘텐츠 인코딩 오류: {e}"}

        try:
            async with httpx.AsyncClient(timeout=60) as client:
                # drive_id 확보
                if not drive_id:
                    r = await client.get(
                        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive",
                        params={"$select": "id"},
                        headers=headers,
                    )
                    r.raise_for_status()
                    drive_id = r.json()["id"]

                upload_path = f"{folder_path}/{filename}".lstrip("/")
                total_size = len(content_bytes)
                CHUNK_LIMIT = 4 * 1024 * 1024  # 4MB

                if total_size <= CHUNK_LIMIT:
                    # 단순 PUT 업로드
                    put_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/root:/{upload_path}:/content"
                    put_headers = {
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/octet-stream",
                    }
                    resp = await client.put(put_url, content=content_bytes, headers=put_headers)
                    resp.raise_for_status()
                    data = resp.json()
                else:
                    # 업로드 세션 (대용량)
                    session_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/root:/{upload_path}:/createUploadSession"
                    sess_resp = await client.post(
                        session_url,
                        json={"item": {"@microsoft.graph.conflictBehavior": "replace", "name": filename}},
                        headers=headers,
                    )
                    sess_resp.raise_for_status()
                    upload_url = sess_resp.json()["uploadUrl"]

                    CHUNK_SIZE = 10 * 1024 * 1024  # 10MB 청크
                    data = {}
                    for start in range(0, total_size, CHUNK_SIZE):
                        end = min(start + CHUNK_SIZE - 1, total_size - 1)
                        chunk = content_bytes[start:end + 1]
                        chunk_headers = {
                            "Content-Range": f"bytes {start}-{end}/{total_size}",
                            "Content-Length": str(len(chunk)),
                        }
                        chunk_resp = await client.put(upload_url, content=chunk, headers=chunk_headers)
                        if chunk_resp.status_code in (200, 201):
                            data = chunk_resp.json()
                        elif chunk_resp.status_code != 202:
                            return {"success": False, "status_code": chunk_resp.status_code, "message": chunk_resp.text[:300]}

            log.info("onedrive_upload_file: drive=%s path=%s size=%d (by %s)", drive_id, upload_path, total_size, get_user_email())
            return {
                "success": True,
                "filename": filename,
                "path": upload_path,
                "drive_id": drive_id,
                "item_id": data.get("id"),
                "size": total_size,
                "web_url": data.get("webUrl"),
                "uploaded_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("onedrive_upload_file 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-17: 파일/폴더 이름 변경
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def onedrive_rename_item(
        drive_id: Annotated[str, Field(min_length=1, description="드라이브 ID (sharepoint_search 반환 drive_id)")],
        item_id: Annotated[str, Field(min_length=1, description="파일/폴더 ID (sharepoint_search 반환 item_id)")],
        new_name: Annotated[str, Field(min_length=1, description="변경할 이름 (확장자 포함, 예: report_v2.xlsx)")],
    ) -> dict:
        """OneDrive/SharePoint 파일 또는 폴더의 이름을 변경합니다.
        drive_id와 item_id는 sharepoint_search 결과에서 확인하세요.
        Azure AD 앱에 Files.ReadWrite.All 또는 Sites.ReadWrite.All 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(url, json={"name": new_name}, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            log.info("onedrive_rename_item: drive=%s item=%s → %s (by %s)", drive_id, item_id, new_name, get_user_email())
            return {
                "success": True,
                "item_id": data.get("id"),
                "new_name": data.get("name"),
                "web_url": data.get("webUrl"),
            }
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-18: 파일/폴더 이동
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def onedrive_move_item(
        drive_id: Annotated[str, Field(min_length=1, description="드라이브 ID (sharepoint_search 반환 drive_id)")],
        item_id: Annotated[str, Field(min_length=1, description="이동할 파일/폴더 ID (sharepoint_search 반환 item_id)")],
        new_parent_item_id: Annotated[str, Field(min_length=1, description="대상 폴더 ID (sharepoint_folder_search 반환 item_id)")],
    ) -> dict:
        """OneDrive/SharePoint 파일 또는 폴더를 다른 폴더로 이동합니다.
        drive_id/item_id는 sharepoint_search, 대상 폴더 ID는 sharepoint_folder_search로 확인하세요.
        Azure AD 앱에 Files.ReadWrite.All 또는 Sites.ReadWrite.All 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        body = {"parentReference": {"id": new_parent_item_id}}
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(url, json=body, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            log.info("onedrive_move_item: drive=%s item=%s → parent=%s (by %s)", drive_id, item_id, new_parent_item_id, get_user_email())
            return {
                "success": True,
                "item_id": data.get("id"),
                "name": data.get("name"),
                "web_url": data.get("webUrl"),
                "new_parent_id": (data.get("parentReference") or {}).get("id"),
            }
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-19: 파일/폴더 삭제
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "idempotentHint": True})
    async def onedrive_delete_item(
        drive_id: Annotated[str, Field(min_length=1, description="드라이브 ID (sharepoint_search 반환 drive_id)")],
        item_id: Annotated[str, Field(min_length=1, description="삭제할 파일/폴더 ID (sharepoint_search 반환 item_id)")],
    ) -> dict:
        """OneDrive/SharePoint 파일 또는 폴더를 삭제합니다 (휴지통으로 이동).
        drive_id와 item_id는 sharepoint_search 결과에서 확인하세요.
        삭제 전 반드시 사용자에게 확인받으세요.
        Azure AD 앱에 Files.ReadWrite.All 또는 Sites.ReadWrite.All 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.delete(url, headers=headers)
                resp.raise_for_status()
            log.info("onedrive_delete_item: drive=%s item=%s (by %s)", drive_id, item_id, get_user_email())
            return {"success": True, "drive_id": drive_id, "item_id": item_id, "deleted_by": get_user_email(), "message": "항목이 휴지통으로 이동되었습니다."}
        except httpx.HTTPStatusError as e:
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"항목을 찾을 수 없습니다: {item_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-20: 공유 링크 생성
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def onedrive_create_share_link(
        drive_id: Annotated[str, Field(min_length=1, description="드라이브 ID (sharepoint_search 반환 drive_id)")],
        item_id: Annotated[str, Field(min_length=1, description="파일/폴더 ID (sharepoint_search 반환 item_id)")],
        link_type: Annotated[str, Field(description="링크 유형: view(읽기 전용), edit(편집 가능)")] = "view",
        scope: Annotated[str, Field(description="공유 범위: organization(사내 전용, 권장), anonymous(외부 공유 — 회사 정책 확인 필요)")] = "organization",
    ) -> dict:
        """OneDrive/SharePoint 파일 또는 폴더의 공유 링크를 생성합니다.
        anonymous(외부 공유)는 테넌트 정책에서 허용한 경우에만 동작합니다.
        Azure AD 앱에 Files.ReadWrite.All 또는 Sites.ReadWrite.All 권한이 필요합니다.
        """
        if link_type not in {"view", "edit"}:
            return {"success": False, "message": "link_type은 'view' 또는 'edit'이어야 합니다."}
        if scope not in {"organization", "anonymous"}:
            return {"success": False, "message": "scope는 'organization' 또는 'anonymous'이어야 합니다."}

        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/createLink"
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(url, json={"type": link_type, "scope": scope}, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            link = data.get("link", {})
            log.info("onedrive_create_share_link: drive=%s item=%s type=%s scope=%s (by %s)", drive_id, item_id, link_type, scope, get_user_email())
            return {
                "success": True,
                "link_type": link.get("type"),
                "scope": link.get("scope"),
                "web_url": link.get("webUrl"),
                "created_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-21: 파일 버전 목록 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def onedrive_list_versions(
        drive_id: Annotated[str, Field(min_length=1, description="드라이브 ID (sharepoint_search 반환 drive_id)")],
        item_id: Annotated[str, Field(min_length=1, description="파일 ID (sharepoint_search 반환 item_id)")],
    ) -> dict:
        """OneDrive/SharePoint 파일의 버전 목록을 조회합니다.
        drive_id와 item_id는 sharepoint_search 결과에서 확인하세요.
        Azure AD 앱에 Files.ReadWrite.All 또는 Sites.ReadWrite.All 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        headers = {"Authorization": f"Bearer {access_token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/versions"
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(url, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            versions = [
                {
                    "version_id": v.get("id"),
                    "last_modified": v.get("lastModifiedDateTime"),
                    "modified_by": (v.get("lastModifiedBy") or {}).get("user", {}).get("displayName"),
                    "size": (v.get("size") or 0),
                }
                for v in data.get("value", [])
            ]
            return {"success": True, "drive_id": drive_id, "item_id": item_id, "version_count": len(versions), "versions": versions}
        except httpx.HTTPStatusError as e:
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-22: Outlook 범주 목록 조회
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def outlook_category_list(
        mailbox_owner: Annotated[str, Field(description="메일박스 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 사용자 정의 범주(masterCategories) 전체 목록을 조회합니다.
        반환된 category_id를 outlook_category_update_color / outlook_category_delete 도구에 사용하세요.
        Azure AD 앱에 MailboxSettings.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/outlook/masterCategories"
            if mailbox_owner
            else "https://graph.microsoft.com/v1.0/me/outlook/masterCategories"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.get(url, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            categories = [
                {"category_id": c.get("id"), "display_name": c.get("displayName"), "color": c.get("color")}
                for c in data.get("value", [])
            ]
            log.info("outlook_category_list: %d개 반환 (owner=%s)", len(categories), owner)
            return {"success": True, "owner": owner, "count": len(categories), "categories": categories}
        except httpx.HTTPStatusError as e:
            log.error("outlook_category_list HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            if e.response.status_code == 403:
                return {"success": False, "status_code": 403, "message": "권한 오류(403): Azure AD 앱에 'MailboxSettings.ReadWrite' Delegated 권한이 필요합니다."}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("outlook_category_list 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-23: Outlook 범주 생성
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def outlook_category_create(
        display_name: Annotated[str, Field(min_length=1, description="생성할 범주 이름. 한 번 생성하면 이름을 변경할 수 없습니다 (Graph API 제약).")],
        color: Annotated[str, Field(description=(
            "색상 프리셋. none(없음), preset0(빨강), preset1(주황), preset2(연갈색), preset3(노랑), "
            "preset4(녹색), preset5(청록), preset6(올리브), preset7(파랑), preset8(보라), preset9(자주), "
            "preset10(강철), preset11(진한강철), preset12(회색), preset13(진한회색), preset14(검정), "
            "preset15(진한빨강), preset16(진한주황), preset17(진한갈색), preset18(진한노랑), "
            "preset19(진한녹색), preset20(진한청록), preset21(진한올리브), preset22(진한파랑), "
            "preset23(진한보라), preset24(진한자주)"
        ))] = "none",
        mailbox_owner: Annotated[str, Field(description="메일박스 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 사용자 정의 범주를 새로 생성합니다.
        ※ 범주 이름(display_name)은 생성 후 변경 불가 (Graph API 제약).
           이름이 틀렸을 경우 삭제 후 재생성이 필요하며, 기존에 태그된 이메일·일정의 범주 태그가 사라집니다.
        Azure AD 앱에 MailboxSettings.ReadWrite Delegated 권한이 필요합니다.
        """
        if color not in _VALID_CATEGORY_COLORS:
            return {"success": False, "message": f"유효하지 않은 색상: '{color}'. none 또는 preset0~preset24 중 하나를 입력하세요."}

        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/outlook/masterCategories"
            if mailbox_owner
            else "https://graph.microsoft.com/v1.0/me/outlook/masterCategories"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(url, json={"displayName": display_name, "color": color}, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            log.info("outlook_category_create: name=%s color=%s id=%s (by %s)", display_name, color, data.get("id"), get_user_email())
            return {
                "success": True,
                "category_id": data.get("id"),
                "display_name": data.get("displayName"),
                "color": data.get("color"),
                "created_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            log.error("outlook_category_create HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            if e.response.status_code == 403:
                return {"success": False, "status_code": 403, "message": "권한 오류(403): Azure AD 앱에 'MailboxSettings.ReadWrite' Delegated 권한이 필요합니다."}
            if e.response.status_code in (400, 409):
                return {"success": False, "status_code": e.response.status_code, "message": f"범주 생성 실패 (이미 같은 이름이 존재할 수 있음): {e.response.text[:300]}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("outlook_category_create 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-24: Outlook 범주 색상 변경
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": True})
    async def outlook_category_update_color(
        category_id: Annotated[str, Field(min_length=1, description="색상을 변경할 범주 ID (outlook_category_list 반환 'category_id' 값)")],
        color: Annotated[str, Field(description=(
            "새 색상 프리셋. none(없음), preset0(빨강), preset1(주황), preset2(연갈색), preset3(노랑), "
            "preset4(녹색), preset5(청록), preset6(올리브), preset7(파랑), preset8(보라), preset9(자주), "
            "preset10(강철), preset11(진한강철), preset12(회색), preset13(진한회색), preset14(검정), "
            "preset15(진한빨강), preset16(진한주황), preset17(진한갈색), preset18(진한노랑), "
            "preset19(진한녹색), preset20(진한청록), preset21(진한올리브), preset22(진한파랑), "
            "preset23(진한보라), preset24(진한자주)"
        ))],
        mailbox_owner: Annotated[str, Field(description="메일박스 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 범주의 색상을 변경합니다.
        ※ 범주 이름(displayName)은 Graph API 제약으로 변경 불가 — 색상(color)만 수정됩니다.
        category_id는 outlook_category_list 도구로 먼저 확인하세요.
        Azure AD 앱에 MailboxSettings.ReadWrite Delegated 권한이 필요합니다.
        """
        if color not in _VALID_CATEGORY_COLORS:
            return {"success": False, "message": f"유효하지 않은 색상: '{color}'. none 또는 preset0~preset24 중 하나를 입력하세요."}

        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/outlook/masterCategories/{category_id}"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/outlook/masterCategories/{category_id}"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.patch(url, json={"color": color}, headers=headers)
                resp.raise_for_status()
                data = resp.json()
            log.info("outlook_category_update_color: id=%s color=%s (by %s)", category_id, color, get_user_email())
            return {
                "success": True,
                "category_id": data.get("id"),
                "display_name": data.get("displayName"),
                "color": data.get("color"),
                "updated_by": get_user_email(),
            }
        except httpx.HTTPStatusError as e:
            log.error("outlook_category_update_color HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            if e.response.status_code == 403:
                return {"success": False, "status_code": 403, "message": "권한 오류(403): Azure AD 앱에 'MailboxSettings.ReadWrite' Delegated 권한이 필요합니다."}
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"범주를 찾을 수 없습니다: {category_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("outlook_category_update_color 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-25: Outlook 범주 삭제
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True, "idempotentHint": True})
    async def outlook_category_delete(
        category_id: Annotated[str, Field(min_length=1, description="삭제할 범주 ID (outlook_category_list 반환 'category_id' 값)")],
        mailbox_owner: Annotated[str, Field(description="메일박스 소유자 이메일 (비우면 로그인 계정)")] = "",
    ) -> dict:
        """Outlook 범주를 삭제합니다.
        ※ 삭제된 범주는 복구 불가하며, 해당 범주가 태그된 모든 이메일·일정에서 범주 태그가 함께 제거됩니다.
        category_id는 outlook_category_list 도구로 먼저 확인하세요.
        삭제 전 반드시 사용자에게 확인받으세요.
        Azure AD 앱에 MailboxSettings.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err

        from ms365.auth import get_access_token
        import httpx

        access_token = get_access_token()
        if not access_token:
            return {"success": False, "message": "MS365 토큰을 가져오는데 실패했습니다. ms365_login을 다시 호출해주세요."}

        await fetch_user_profile()
        owner = mailbox_owner or get_user_email()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = (
            f"https://graph.microsoft.com/v1.0/users/{owner}/outlook/masterCategories/{category_id}"
            if mailbox_owner
            else f"https://graph.microsoft.com/v1.0/me/outlook/masterCategories/{category_id}"
        )
        try:
            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.delete(url, headers=headers)
                resp.raise_for_status()
            log.info("outlook_category_delete: id=%s (by %s)", category_id, get_user_email())
            return {
                "success": True,
                "category_id": category_id,
                "deleted_by": get_user_email(),
                "message": "범주가 삭제되었습니다. 해당 범주가 태그된 이메일·일정의 범주 태그도 함께 제거됩니다.",
            }
        except httpx.HTTPStatusError as e:
            log.error("outlook_category_delete HTTP 오류: %s — %s", e.response.status_code, e.response.text)
            if e.response.status_code == 403:
                return {"success": False, "status_code": 403, "message": "권한 오류(403): Azure AD 앱에 'MailboxSettings.ReadWrite' Delegated 권한이 필요합니다."}
            if e.response.status_code == 404:
                return {"success": False, "status_code": 404, "message": f"범주를 찾을 수 없습니다: {category_id}"}
            return {"success": False, "status_code": e.response.status_code, "message": e.response.text[:300]}
        except Exception as e:
            log.error("outlook_category_delete 실패: %s", e)
            return {"success": False, "message": str(e)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # Tool MS365-26: Outlook 임시보관함(Drafts) 초안 작성
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    @mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False})
    async def create_draft_email(
        to: Annotated[str, Field(description="수신자 이메일 주소")],
        subject: Annotated[str, Field(min_length=1, description="이메일 제목")],
        body_text: Annotated[str, Field(description="이메일 본문 (텍스트)")],
    ) -> dict:
        """Outlook 임시보관함(Drafts)에 이메일 초안을 작성하여 저장합니다.
        발송되지는 않으며, 로그인한 본인 계정의 임시보관함에 저장됩니다.
        이후 Outlook 클라이언트에서 검토 후 직접 발송하거나, outlook_email_search로 찾아
        outlook_email_move/delete 등으로 관리할 수 있습니다.
        Azure AD 앱에 Mail.ReadWrite Delegated 권한이 필요합니다.
        """
        auth_err = require_auth()
        if auth_err:
            return auth_err
        graph = get_graph_client()
        await fetch_user_profile()
        try:
            from msgraph.generated.models.message import Message
            from msgraph.generated.models.recipient import Recipient
            from msgraph.generated.models.email_address import EmailAddress
            from msgraph.generated.models.item_body import ItemBody
            from msgraph.generated.models.body_type import BodyType

            email_addr = EmailAddress()
            email_addr.address = to

            recip = Recipient()
            recip.email_address = email_addr

            item_body = ItemBody()
            item_body.content_type = BodyType.Text
            item_body.content = body_text

            msg = Message()
            msg.subject = subject
            msg.body = item_body
            msg.to_recipients = [recip]

            # graph.me.messages.post — 발송하지 않고 임시보관함(Drafts)에 생성
            result = await graph.me.messages.post(msg)
            sender = get_user_email()
            log.info("create_draft_email: owner=%s to=%s subject=%s id=%s",
                     sender, to, subject, result.id)
            return {
                "success": True,
                "message_id": result.id,
                "owner": sender,
                "to": to,
                "subject": subject,
                "web_link": result.web_link,
                "folder": "Drafts",
            }
        except Exception as e:
            log.error("create_draft_email 실패: %s", e)
            return {"success": False, "message": str(e)}

    log.info("MS365 Graph API 쓰기 도구 26개 등록 완료")
