# royal-mcp 프로젝트 가이드

로얄앤컴퍼니 사내 MCP 서버. MS365 Graph API(이메일·캘린더·Teams·SharePoint·OneDrive)를 Claude Code에서 사용할 수 있게 연결한다.

---

## 프로젝트 구조

```
royal-mcp/
├── server.py               # MCP 서버 진입점 (STDIO / HTTP 모드)
├── ms365/
│   ├── __init__.py         # 공개 인터페이스 재수출
│   ├── auth.py             # 인증 (Device Code / Azure AD JWT)
│   ├── graph.py            # MS365 쓰기 도구 (send_*, create_*, update_*)
│   └── graph_read.py       # MS365 읽기 도구 + 파일 파싱
├── erp/
│   ├── __init__.py         # 공개 인터페이스 재수출
│   └── insa.py             # 기간계 DB API 도구 (인사 캘린더 등)
├── requirements.txt        # 직원 배포용 (STDIO 모드)
├── requirements-dev.txt    # 개발자용 (HTTP 모드 포함, -r requirements.txt)
├── .env.example            # 환경변수 템플릿
└── 배포_가이드.md            # 직원 PC 설치 가이드 (10분 소요)
```

---

## 실행 모드

### STDIO 모드 (직원 PC — 기본)
```
python server.py
```
- Claude Code 로컬 연결용 (`mcp.json` 설정 필요)
- 인증: Device Code Flow (`ms365 로그인해줘` 호출)
- 의존성: `pip install -r requirements.txt`

### HTTP 모드 (개발자 / 사내 서버)
```
python server.py --http
python server.py --http --port 8080 --host 0.0.0.0
```
- 다중 사용자 공용 서버용
- 인증: Azure AD OAuth (JWT 검증, ms365_login 불필요)
- 의존성: `pip install -r requirements-dev.txt` (uvicorn, starlette 포함)

---

## 환경변수 (`.env`)

| 변수 | 필수 | 설명 |
|------|------|------|
| `TENANT_ID` | ✅ | Azure AD 테넌트 ID |
| `CLIENT_ID` | ✅ | Azure AD 앱 등록 클라이언트 ID |
| `MCP_SERVER_NAME` | 선택 | 서버 표시 이름 (기본: `royal-mcp`) |
| `MS365_AUTO_AUTH` | 선택 | 서버 시작 시 토큰 캐시 자동 복원 여부 (기본: `true`) |
| `MS365_TOKEN_CACHE_PATH` | 선택 | 토큰 캐시 파일 경로 (기본: `~/.royal-mcp/token_cache.bin`) |

`.env.example`을 복사해 `.env`를 만든다. `.env`는 절대 커밋하지 않는다.

---

## 의존성 구조

### requirements.txt (배포용)
- **MCP 프레임워크**: `mcp`, `fastmcp>=2.0.0,<4`
- **MS365**: `msgraph-sdk`, `azure-identity`, `msal-extensions`
- **인증**: `PyJWT`, `httpx`
- **파일 파싱** (graph_read.py 내 지연 import): `python-docx`, `python-pptx`, `pypdf`, `openpyxl`
- **기타**: `python-dotenv`, `pydantic`, `tzdata` (Windows 타임존)

### requirements-dev.txt (개발자 추가)
- `uvicorn`, `starlette` — HTTP 모드 전용 (server.py에서 직접 import)

> **주의**: `uvicorn`과 `starlette`는 fastmcp의 전이 의존성으로 우연히 설치되지만, 명시적으로 `requirements-dev.txt`에만 포함한다. 배포용에 추가하지 않는다.

---

## 인증 흐름

### STDIO 모드 (Device Code)
1. `ms365_login()` 호출 → Device Code + URL 반환
2. 브라우저에서 `https://microsoft.com/devicelogin` 접속 → `@iroyal.kr` 계정 로그인

### HTTP 모드 (Azure AD JWT)
1. 클라이언트가 `/.well-known/oauth-protected-resource` 탐색
2. Azure AD OAuth로 사용자 직접 인증 → Bearer 토큰 취득
3. `_AzureADMiddleware`가 요청마다 JWT 검증 (`validate_azure_token`)
4. 검증된 사용자 정보를 `contextvars`에 저장 → 도구에서 참조

### Azure AD 앱 설정 요건
- 인증 → 퍼블릭 클라이언트 흐름 허용 = **예** (STDIO용)
- 인증 → 리디렉션 URI: Claude.ai OAuth 콜백 URL 추가 (HTTP 모드용)
- API 사용 권한: Delegated 권한만 사용 (Application 권한 없음)
- 필요 스코프:

| 스코프 | 용도 |
|--------|------|
| `User.Read` | 로그인 사용자 프로필 조회 |
| `User.Read.All` | 임직원 전체 목록 조회 (관리자 동의 필요) |
| `OrgContact.Read.All` | 조직 연락처 조회 (관리자 동의 필요) |
| `Mail.ReadWrite` | 이메일 읽기·범주·플래그·이동·삭제 |
| `Mail.Send` | 이메일 발송·답장·전달 |
| `Calendars.ReadWrite` | 캘린더 일정 조회·생성·수정·삭제 |
| `ChannelMessage.Send` | Teams 채널 메시지 발송 |
| `ChannelMessage.ReadWrite` | Teams 채널 메시지 수정·삭제 (`ChannelMessage.Send`와 별도로 관리자 동의 필요) |
| `Chat.ReadWrite` | Teams 1:1·그룹 채팅 발송·읽기 |
| `Team.ReadBasic.All` | 가입 팀 목록 조회 |
| `Sites.ReadWrite.All` | SharePoint 파일 읽기·쓰기 |
| `Files.ReadWrite.All` | OneDrive 파일 읽기·쓰기 |
| `MailboxSettings.ReadWrite` | Outlook 범주(masterCategories) 관리 |
| `Reports.Read.All` | 조직 사용량 보고서 조회 (로그인 계정이 Reports Reader 이상 관리자 역할 필요) |

> **참고**: `ChannelMessage.Read.All`(채널 메시지 읽기)은 관리자 동의가 필요하므로 기본 스코프에서 제외됨. 채널 메시지 조회 도구 사용 시 별도 추가 필요.

---

## 제공 도구 목록 (총 61개)

### 서버 도구 (server.py) — 3개

| 도구 | 설명 |
|------|------|
| `ms365_login` | MS365 로그인 시작 (STDIO: Device Code, HTTP: Azure AD OAuth 상태 반환) |
| `ms365_logout` | MS365 로그아웃 (STDIO: 캐시 삭제, HTTP: Azure 포털 안내) |
| `ms365_status` | 현재 로그인 상태 및 사용자 정보 확인 |

### 쓰기 도구 (graph.py) — 35개

| 번호 | 도구 | 설명 |
|------|------|------|
| 1 | `send_teams_channel_message` | Teams 채널 메시지 발송 (body_type='html' 지정 시 굵게·목록·줄바꿈 등 서식 적용, `mentions`로 사용자·채널·팀 @멘션 지원, `attachments`로 파일 카드(미리보기) 첨부 지원) |
| 2 | `create_sharepoint_file` | SharePoint 텍스트 파일 생성 |
| 3 | `send_email` | Outlook 이메일 발송 (첨부파일 지원: 로컬 파일 3MB 이하 또는 OneDrive/SharePoint 파일의 사내 전용 공유 링크) |
| 4 | `create_calendar_event` | 캘린더 일정 생성 (Teams 링크 포함 옵션) |
| 5 | `send_teams_chat_message` | Teams 채팅 메시지 발송 (`recipient_email`로 1:1, `chat_id`로 기존 그룹·회의 채팅 지정, `body_type` text/html, `attachments`로 파일 카드(미리보기) 첨부 지원) |
| 6 | `outlook_email_update_category` | 이메일 범주(Category) 설정 |
| 7 | `update_calendar_event` | 캘린더 일정 수정 |
| 8 | `delete_calendar_event` | 캘린더 일정 삭제 |
| 9 | `outlook_email_reply` | 이메일 답장 (발신자만, 첨부파일 지원) |
| 10 | `outlook_email_reply_all` | 이메일 전체 답장 (첨부파일 지원) |
| 11 | `outlook_email_forward` | 이메일 전달 (첨부파일 추가 지원, 원본 첨부는 자동 유지) |
| 12 | `outlook_email_move` | 이메일 폴더 이동 / 아카이브 |
| 13 | `outlook_email_delete` | 이메일 삭제 (휴지통 이동) |
| 14 | `outlook_email_mark_read` | 이메일 읽음/안읽음 처리 |
| 15 | `outlook_email_set_flag` | 이메일 팔로우업 플래그 설정 |
| 16 | `onedrive_upload_file` | OneDrive/SharePoint 파일 업로드 (4MB 초과 시 청크 자동 처리) |
| 17 | `onedrive_rename_item` | 파일/폴더 이름 변경 |
| 18 | `onedrive_move_item` | 파일/폴더 이동 |
| 19 | `onedrive_delete_item` | 파일/폴더 삭제 (휴지통 이동) |
| 20 | `onedrive_create_share_link` | 공유 링크 생성 (권한: view/edit/review/blocksDownload, 범위: organization/users/anonymous, 만료일 지정). review·blocksDownload·users 범위는 Graph beta API 사용 |
| 21 | `onedrive_list_versions` | 파일 버전 목록 조회 |
| 22 | `outlook_category_list` | Outlook 범주(masterCategories) 목록 조회 |
| 23 | `outlook_category_create` | Outlook 범주 생성 |
| 24 | `outlook_category_update_color` | Outlook 범주 색상 변경 |
| 25 | `outlook_category_delete` | Outlook 범주 삭제 |
| 26 | `create_draft_email` | Outlook 임시보관함(Drafts)에 이메일 초안 작성·저장 (발송 안 함, 첨부파일 지원) |
| 27 | `excel_write_range` | Excel 워크시트 범위에 값 쓰기 (세션 불필요, 즉시 파일에 반영·되돌리기 불가) |
| 28 | `word_find_replace` | Word 문서(.docx) 본문·표 텍스트 찾기/바꾸기. 변경건수 반환, 0건이면 재업로드 생략 |
| 29 | `pptx_find_replace` | PowerPoint(.pptx) 슬라이드·표·그룹도형 텍스트 찾기/바꾸기. 변경건수 반환, 0건이면 재업로드 생략 |
| 30 | `excel_add_worksheet` | Excel 파일에 새 워크시트 추가 (이름 생략 시 자동 명명). 값 입력은 `excel_write_range`를 이어서 호출 |
| 31 | `update_teams_channel_message` | Teams 채널 메시지 수정 (본문 전체 교체, 답글 수정 지원, `mentions`로 사용자·채널·팀 @멘션 지원, 본인이 보낸 메시지만 가능) |
| 32 | `delete_teams_channel_message` | Teams 채널 메시지 삭제 (Soft Delete, 답글 삭제 지원, 본인이 보낸 메시지만 가능) |
| 33 | `onedrive_grant_permission` | OneDrive/SharePoint 파일·폴더에 특정 사내 사용자의 read/write 권한 부여 (invite) |
| 34 | `onedrive_list_permissions` | 파일·폴더에 설정된 권한 목록 조회 (permission_id 확인, 상속·공유링크 여부 포함) |
| 35 | `onedrive_revoke_permission` | 특정 사용자의 권한 회수 (상속된 권한·다인 공유링크는 회수 거부하고 사유 안내) |

### 읽기 도구 (graph_read.py) — 22개

| 번호 | 도구 | 설명 |
|------|------|------|
| R1 | `chat_message_search` | Teams 채널·채팅 메시지 키워드 검색 (본문 전체 반환, 단 전체 검색 모드는 검색 엔진 제공 본문) |
| R2 | `find_meeting_availability` | 참석자 캘린더 가용 시간 조회 (getSchedule) |
| R3 | `outlook_calendar_search` | 캘린더 일정 검색 (날짜 범위·키워드) |
| R4 | `outlook_email_search` | 이메일 KQL 키워드 검색 |
| R5 | `read_resource` | 메일 본문·일정 상세·파일 내용 읽기 |
| R6 | `sharepoint_search` | SharePoint/OneDrive 파일·폴더 검색 |
| R7 | `sharepoint_folder_search` | SharePoint/OneDrive 폴더만 검색 |
| R8 | `org_users_search` | 임직원·조직 연락처 목록 조회 및 검색 |
| R9 | `org_user_detail` | 특정 임직원 상세 정보 조회 |
| R10 | `teams_channels_search` | Teams 팀·채널 이름 검색 및 ID 조회 |
| R11 | `teams_group_chat_messages` | Teams 그룹 채팅방 메시지 날짜 조회 (본문 전체 반환) |
| R12 | `teams_channel_messages` | Teams 팀 채널 메시지 날짜 조회 (본문 전체 반환) |
| R13 | `read_email_attachment` | 이메일 첨부파일 목록 조회 및 내용 읽기 |
| R14 | `read_email_attachments_all` | 이메일 1건의 첨부파일 전체 내용을 한 번의 호출로 일괄 읽기 |
| R15 | `excel_list_worksheets` | Excel 파일의 워크시트 목록 조회 (이름/ID/위치) |
| R16 | `excel_read_range` | Excel 워크시트 범위 값 조회. address 생략 시 사용 중인 전체 범위(usedRange) 반환 |
| R17 | `org_usage_report` | 조직 사용량 보고서 조회 (Office365 활성사용자/메일박스/OneDrive/SharePoint/Teams활동/Teams기기/이메일활동). Reports Reader 이상 관리자 계정 필요 |
| R18 | `teams_channel_message_replies` | Teams 채널 메시지의 댓글(스레드 답글) 목록 조회 (본문 전체 반환) |
| R19 | `teams_channel_message_attachments` | Teams 채널 메시지(또는 답글)의 첨부파일·카드 목록 조회 (파일 내용은 content_url을 sharepoint_search/read_resource로 재조회) |
| R20 | `teams_channel_message_images` | Teams 채널 메시지(또는 답글) 본문의 인라인 이미지 조회 (목록만/개별 이미지 base64, 3MB 초과 시 반환 안 함) |
| R21 | `teams_unread_chats` | Teams 채팅(1:1·그룹·회의) 중 읽지 않은 메시지가 있는 채팅과 미확인 메시지 조회. 채팅 목록의 `viewpoint.lastMessageReadDateTime`(마지막으로 읽은 시각) 기준, 본인 메시지 제외, 읽음 처리는 하지 않음 |
| R22 | `teams_channel_new_messages` | 가입 팀 채널에서 기준 시각(since, 기본 최근 24시간) 이후 작성된 새 글·댓글 조회. Graph API가 채널 읽음 상태를 제공하지 않아 Teams 화면의 '읽지 않음'과 다를 수 있음. 채널 메시지 delta(`lastModifiedDateTime gt` 필터, `replies` 확장) 사용 |

### 기간계 DB ERP 도구 (erp/insa.py) — 1개

| 번호 | 도구 | 설명 |
|------|------|------|
| E1 | `get_insa_calendar` | 특정 날짜의 인사 캘린더(근무일/휴일/요일) 조회. **MS365 로그인 필요** |

> `get_insa_calendar` 응답 필드: `TP_WEEK` (1=일~7=토), `FG1_HOLIDAY` (W=평일, H=공휴일)

---

## 이메일 첨부파일 (graph.py — EmailAttachment)

`send_email`, `create_draft_email`, `outlook_email_reply`, `outlook_email_reply_all`, `outlook_email_forward` 5개 도구가 공통으로 `attachments` 파라미터를 지원한다. 각 항목은 아래 둘 중 하나를 지정한다.

| 지정 방식 | 동작 | 제약 |
|-----------|------|------|
| `local_path` | 파일을 읽어 base64로 인코딩해 이메일에 직접 첨부(fileAttachment) | 파일당 3MB 이하, 로컬 첨부 합계도 3MB 이하. **STDIO 모드에서만 허용**(사용자 본인 PC 경로 기준). **HTTP 모드(사내 공용 서버)에서는 서버 파일시스템 임의 접근 위험 때문에 거부됨** |
| `drive_id` + `item_id` | 파일 내용을 내려받지 않고, 사내 전용(`organizationView`) 공유 링크를 생성해 링크 첨부(referenceAttachment) | drive_id/item_id는 `sharepoint_search`로 확인. 수신자가 사내 계정으로 로그인해야 열람 가능. 파일 크기 제약 없음(HTTP 모드에서도 사용 가능) |

3MB(또는 로컬 첨부 합계 3MB)를 초과하는 파일은 `onedrive_upload_file`로 먼저 업로드한 뒤 `drive_id`+`item_id`로 첨부하도록 안내한다. `outlook_email_reply`/`reply_all`/`forward`는 첨부파일이 있으면 `body_type` 설정과 무관하게 내부적으로 초안(Draft) 경유 방식으로 전환되며, 첨부 처리 중 오류가 발생하면 초안을 정리(삭제)하고 발송하지 않는다.

---

## Teams @멘션 (graph.py — _build_mention_body)

`send_teams_channel_message`, `update_teams_channel_message`는 `mentions` 목록을 지원한다. 항목은 예약어 `channel`(현재 채널 → `mentioned.conversation`, conversationIdentityType=channel, 표시 이름은 채널명), `team`(현재 팀 → conversationIdentityType=team, 표시 이름은 팀명) 또는 사용자 이메일(`/users/{email}` 조회 → `mentioned.user`)이며 예약어는 대소문자를 구분하지 않는다. 대상별 ID·표시 이름을 조회한 뒤 본문의 `{mention0}`, `{mention1}`… 자리표시자를 `<at id="N">표시이름</at>`으로 치환한 뒤 Graph `mentions` 배열을 함께 보낸다. 자리표시자가 없으면 본문 맨 앞에 멘션을 붙이며, 멘션이 있으면 body_type과 무관하게 HTML로 발송한다(`text`는 이스케이프 후 줄바꿈을 `<br>`로 변환). 같은 자리표시자 중복 사용, 범위를 벗어난 자리표시자, 존재하지 않는 사용자, 채널명·팀명 조회 실패는 발송 전에 오류로 반환한다. 반환값 `mentioned`는 `[{type: user|channel|team, name, target}]` 형식이다. 예: 주간보고 공지 `mentions=["channel"], message="<p>{mention0}</p><p>임원보고자료 ...</p>"`

---

## Teams 채널·채팅 메시지 첨부파일 (graph.py — TeamsAttachment)

`send_teams_channel_message`와 `send_teams_chat_message`는 `attachments` 목록을 지원한다. 아래 표와 규칙은 채널 기준이며, 채팅의 차이는 이 절 끝의 "채팅 첨부"에 정리한다. Teams 첨부는 파일을 메시지에 싣지 않고 SharePoint 파일을 참조(`contentType: "reference"`)하며, Teams가 SharePoint 썸네일로 파일 카드(미리보기)를 그린다.

| 지정 방식 | 동작 | 제약 |
|-----------|------|------|
| `local_path` | `GET /teams/{team}/channels/{channel}/filesFolder`로 채널 폴더(채널 '파일' 탭)를 찾아 그 루트에 업로드(4MB 초과 시 업로드 세션) 후 참조 첨부 | 파일당 250MB 이하. **STDIO 모드에서만 허용**(HTTP 모드는 서버 파일시스템 접근 위험으로 거부). 채널 폴더가 아직 없으면(채널 '파일' 탭을 한 번도 안 연 경우) 오류 |
| `drive_id` + `item_id` | 업로드 없이 기존 파일 정보(eTag, webUrl)를 조회해 참조 첨부 | 채널 구성원에게 해당 파일 열람 권한이 있어야 카드가 열림. 폴더는 거부 |

- 첨부 ID는 driveItem `eTag`(`"{GUID},N"`)의 GUID이며, 본문 끝에 `<attachment id="GUID"></attachment>`를 첨부 순서대로 붙인다. 첨부가 있으면 body_type과 무관하게 HTML로 발송한다(`text`는 이스케이프 후 줄바꿈을 `<br>`로 변환).
- `contentUrl`은 driveItem `webUrl`을 쓰되, Office 웹 뷰어 주소(`_layouts/15/Doc.aspx`)로 오면 채널 폴더 webUrl + 파일명으로 대체한다.
- `file_conflict`: `rename`(기본, 같은 이름이 있으면 새 이름으로 저장해 이전 메시지의 첨부 보존) 또는 `replace`(덮어쓰기).
- 입력 검증 오류나 업로드·조회 실패가 하나라도 있으면 메시지를 발송하지 않는다. 이미 업로드된 파일은 채널 폴더에 남고 반환값 `uploaded`에 표시된다.
- 반환값: `attachments`(`[{name, url}]`), `uploaded`(`[{source, name, size, drive_id, item_id, web_url}]`).
- 업로드 직후에는 SharePoint 썸네일이 생성되기 전이라 잠시 아이콘만 보일 수 있다.
- 필요 권한: `ChannelMessage.Send` + `Files.ReadWrite.All`(또는 `Sites.ReadWrite.All`).

**채팅 첨부 (`send_teams_chat_message`)**
- 발송 대상: `recipient_email`(1:1 채팅방 생성 또는 재사용)과 `chat_id`(기존 그룹·회의·1:1 채팅) 중 하나만 지정한다. 새 그룹 채팅은 만들지 않는다.
- `local_path` 업로드 위치: 본인 OneDrive 루트의 `Microsoft Teams Chat Files` 또는 `Microsoft Teams 채팅 파일` 폴더(Teams 앱이 채팅 첨부에 쓰는 폴더). 둘 다 없으면 `Microsoft Teams Chat Files`를 만든다.
- 권한: Graph API로 올린 파일은 자동 공유되지 않으므로, 발송 전에 본인을 제외한 채팅 참여자 전원(`/chats/{id}/members`의 userId)에게 `/invite`로 읽기 권한을 부여한다(`requireSignIn=true`, `sendInvitation=false`). `drive_id`+`item_id`로 지정한 기존 파일에도 같은 권한을 부여하며, 부여에 실패하면(본인에게 공유 권한이 없는 파일, 외부 게스트 참여자 등) 발송하지 않는다.
- `body_type` 기본값은 하위 호환을 위해 `text`이다. 첨부가 있으면 HTML로 변환해 발송한다.
- 반환값: `chat_id`, `chat_type`(oneOnOne/group/meeting), `to_name`(1:1은 수신자 이름, 그룹은 채팅 주제), `attachments`, `uploaded`, `shared_with`(권한을 부여한 참여자 수).
- 필요 권한: `Chat.ReadWrite` + `Files.ReadWrite.All`.

---

## 파일 파싱 (graph_read.py)

OneDrive/SharePoint 파일을 다운로드 후 메모리 내에서 텍스트 추출한다. 디스크에 저장하지 않는다.

| 확장자 | 라이브러리 | import 방식 |
|--------|------------|-------------|
| `.xlsx` | `openpyxl` | 지연 import |
| `.docx` | `python-docx` | 지연 import |
| `.pptx` | `python-pptx` | 지연 import |
| `.pdf` | `pypdf` | 지연 import |

---

## Office 문서 편집 도구 (Excel/Word/PowerPoint)

| 파일 유형 | 방식 | 도구 |
|-----------|------|------|
| Excel | Graph Workbook API로 셀/범위 직접 read·write (세션 불필요, 산발적 호출 기준) | `excel_list_worksheets`, `excel_read_range`, `excel_write_range`, `excel_add_worksheet` |
| Word/PowerPoint | Graph에 콘텐츠 편집 API가 없어 다운로드→python-docx/pptx로 치환→재업로드(덮어쓰기) 방식 | `word_find_replace`, `pptx_find_replace` |

**알려진 한계 (설계상 수용, 사용자 승인됨)**
- 텍스트가 여러 run(서식 조각)에 걸쳐 나뉘어 있으면 매칭되지 않아 `replaced_count: 0`을 반환합니다. 이는 오류가 아니라 python-docx/pptx의 구조적 한계이며, 재업로드 자체를 생략하므로 파일은 변경되지 않습니다.
- 저장 시 문서 전체를 재작성하므로 주석, 추적 변경, 일부 고급 서식이 유실될 수 있습니다. 문제가 있으면 `onedrive_list_versions`로 이전 버전을 확인해 복구하세요.
- Word는 본문 단락과 표만 탐색합니다 (머리글/바닥글, 텍스트 상자, 중첩 표 제외). PowerPoint는 그룹 도형 내부까지 재귀 탐색합니다.

**동시편집 보호**: `word_find_replace`/`pptx_find_replace`는 다운로드 시점의 eTag를 기억해 재업로드 시 `If-Match` 헤더로 검증합니다. 그 사이 다른 사용자가 파일을 수정했으면 덮어쓰지 않고 명확한 오류를 반환합니다. `excel_write_range`는 range 단위로만 PATCH하므로 이 문제 자체가 없습니다.

**실사용 확인된 제약 (2026-07-08)**: 셀 데이터가 전혀 없는 완전히 빈 xlsx(예: `openpyxl.Workbook()` 직후 바로 저장한 파일)를 OneDrive에 올리면 Graph Workbook API가 `FileCorruptTryRepair`/`unsupportedWorkbook` 오류로 거부합니다. 새 Excel 파일을 만들 때는 로컬에서 최소 1개 셀 이상 값을 채운 뒤 업로드하세요(완전한 빈 시트를 만들고 `excel_write_range`로 채우는 순서는 실패함).

---

## 배포 주의사항

### ZIP 배포 시 반드시 제외
- `.venv\` — 생성 PC의 절대 경로가 박혀 있어 타 PC에서 동작 불가
- `.env` — 테넌트 ID 등 민감 정보 포함
- `__pycache__\`, `*.pyc`

### 직원 PC 설치 경로 제한
- 경로에 **한글, 공백 불가** (Python venv 경로 파싱 오류)
- 권장: `C:\royal-mcp\`, `D:\tools\royal-mcp\`

### `.env` 파일 생성 (Windows)
메모장으로 생성 시 저장 형식을 **"모든 파일(\*.\*)"** 으로 변경해야 `.env.txt`로 저장되는 것을 방지할 수 있다.

---

## Git 커밋 규칙

**파일 수정 후 git commit / push를 자동으로 수행하지 않는다.**
작업이 완료되면 변경 내용을 요약하고, 커밋 및 푸시 여부를 반드시 사용자에게 먼저 확인한다.

---

## 자주 발생하는 문제

| 증상 | 원인 | 해결 |
|------|------|------|
| `ImportError: uvicorn` | `requirements.txt`로 설치 후 HTTP 모드 실행 | `pip install -r requirements-dev.txt` |
| MCP 도구가 Claude Code에 안 보임 | `mcp.json` 미반영 | Claude Code 완전 재시작 |
| `.env` 인식 안 됨 | `.env.txt`로 저장됨 | 파일 형식 확인, 확장자 제거 |
| `Activate.ps1` 실행 오류 | PowerShell 실행 정책 제한 | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |
| 로그인 후 권한 오류 | Azure AD 동의 미완료 | 관리자에게 Azure AD 권한 동의 요청 |
| `teams_channel_messages`/`teams_channel_message_replies`/`teams_channel_message_attachments`/`teams_channel_message_images`/`teams_channel_new_messages` 403 오류 | `ChannelMessage.Read.All` 관리자 동의 미완료 | Azure Portal에서 `ChannelMessage.Read.All` 권한 추가 후 재로그인 |
| `update_teams_channel_message`/`delete_teams_channel_message` 403 오류 | `ChannelMessage.ReadWrite` 미동의, 또는 본인이 보낸 메시지가 아님 | Azure Portal에서 `ChannelMessage.ReadWrite` 권한 추가 후 재로그인. 타인이 보낸 메시지는 Graph API 정책상 수정·삭제 불가 |
| 임직원 조회 결과 없음 | `User.Read.All` 관리자 동의 미완료 | Azure Portal에서 `User.Read.All` 권한 추가 |
| `org_usage_report` 403 오류 | 로그인 계정에 Reports Reader 이상 관리자 역할 없음, 또는 권한 추가 후 재로그인 안 함 | Azure AD에서 계정에 Reports Reader 역할 부여 + `ms365_logout` → `ms365_login` 재인증 |
