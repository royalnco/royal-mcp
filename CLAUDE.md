# royal-ms365-mcp 프로젝트 가이드

로얄앤컴퍼니 사내 MCP 서버. MS365 Graph API(이메일·캘린더·Teams·SharePoint·OneDrive)를 Claude Code에서 사용할 수 있게 연결한다.

---

## 프로젝트 구조

```
royal-ms365-mcp/
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
| `MCP_SERVER_NAME` | 선택 | 서버 표시 이름 (기본: `royal-ms365-mcp`) |
| `MS365_AUTO_AUTH` | 선택 | 서버 시작 시 토큰 캐시 자동 복원 여부 (기본: `true`) |
| `MS365_TOKEN_CACHE_PATH` | 선택 | 토큰 캐시 파일 경로 (기본: `~/.royal-ms365-mcp/token_cache.bin`) |

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
| `Chat.ReadWrite` | Teams 1:1·그룹 채팅 발송·읽기 |
| `Team.ReadBasic.All` | 가입 팀 목록 조회 |
| `Sites.ReadWrite.All` | SharePoint 파일 읽기·쓰기 |
| `Files.ReadWrite.All` | OneDrive 파일 읽기·쓰기 |
| `MailboxSettings.ReadWrite` | Outlook 범주(masterCategories) 관리 |

> **참고**: `ChannelMessage.Read.All`(채널 메시지 읽기)은 관리자 동의가 필요하므로 기본 스코프에서 제외됨. 채널 메시지 조회 도구 사용 시 별도 추가 필요.

---

## 제공 도구 목록 (총 44개)

### 서버 도구 (server.py) — 3개

| 도구 | 설명 |
|------|------|
| `ms365_login` | MS365 로그인 시작 (STDIO: Device Code, HTTP: Azure AD OAuth 상태 반환) |
| `ms365_logout` | MS365 로그아웃 (STDIO: 캐시 삭제, HTTP: Azure 포털 안내) |
| `ms365_status` | 현재 로그인 상태 및 사용자 정보 확인 |

### 쓰기 도구 (graph.py) — 26개

| 번호 | 도구 | 설명 |
|------|------|------|
| 1 | `send_teams_channel_message` | Teams 채널 메시지 발송 |
| 2 | `create_sharepoint_file` | SharePoint 텍스트 파일 생성 |
| 3 | `send_email` | Outlook 이메일 발송 |
| 4 | `create_calendar_event` | 캘린더 일정 생성 (Teams 링크 포함 옵션) |
| 5 | `send_teams_chat_message` | Teams 1:1 채팅 메시지 발송 |
| 6 | `outlook_email_update_category` | 이메일 범주(Category) 설정 |
| 7 | `update_calendar_event` | 캘린더 일정 수정 |
| 8 | `delete_calendar_event` | 캘린더 일정 삭제 |
| 9 | `outlook_email_reply` | 이메일 답장 (발신자만) |
| 10 | `outlook_email_reply_all` | 이메일 전체 답장 |
| 11 | `outlook_email_forward` | 이메일 전달 |
| 12 | `outlook_email_move` | 이메일 폴더 이동 / 아카이브 |
| 13 | `outlook_email_delete` | 이메일 삭제 (휴지통 이동) |
| 14 | `outlook_email_mark_read` | 이메일 읽음/안읽음 처리 |
| 15 | `outlook_email_set_flag` | 이메일 팔로우업 플래그 설정 |
| 16 | `onedrive_upload_file` | OneDrive/SharePoint 파일 업로드 (4MB 초과 시 청크 자동 처리) |
| 17 | `onedrive_rename_item` | 파일/폴더 이름 변경 |
| 18 | `onedrive_move_item` | 파일/폴더 이동 |
| 19 | `onedrive_delete_item` | 파일/폴더 삭제 (휴지통 이동) |
| 20 | `onedrive_create_share_link` | 공유 링크 생성 (view/edit, organization/anonymous) |
| 21 | `onedrive_list_versions` | 파일 버전 목록 조회 |
| 22 | `outlook_category_list` | Outlook 범주(masterCategories) 목록 조회 |
| 23 | `outlook_category_create` | Outlook 범주 생성 |
| 24 | `outlook_category_update_color` | Outlook 범주 색상 변경 |
| 25 | `outlook_category_delete` | Outlook 범주 삭제 |
| 26 | `create_draft_email` | Outlook 임시보관함(Drafts)에 이메일 초안 작성·저장 (발송 안 함) |

### 읽기 도구 (graph_read.py) — 14개

| 번호 | 도구 | 설명 |
|------|------|------|
| R1 | `chat_message_search` | Teams 채널·채팅 메시지 키워드 검색 |
| R2 | `find_meeting_availability` | 참석자 캘린더 가용 시간 조회 (getSchedule) |
| R3 | `outlook_calendar_search` | 캘린더 일정 검색 (날짜 범위·키워드) |
| R4 | `outlook_email_search` | 이메일 KQL 키워드 검색 |
| R5 | `read_resource` | 메일 본문·일정 상세·파일 내용 읽기 |
| R6 | `sharepoint_search` | SharePoint/OneDrive 파일·폴더 검색 |
| R7 | `sharepoint_folder_search` | SharePoint/OneDrive 폴더만 검색 |
| R8 | `org_users_search` | 임직원·조직 연락처 목록 조회 및 검색 |
| R9 | `org_user_detail` | 특정 임직원 상세 정보 조회 |
| R10 | `teams_channels_search` | Teams 팀·채널 이름 검색 및 ID 조회 |
| R11 | `teams_group_chat_messages` | Teams 그룹 채팅방 메시지 날짜 조회 |
| R12 | `teams_channel_messages` | Teams 팀 채널 메시지 날짜 조회 |
| R13 | `read_email_attachment` | 이메일 첨부파일 목록 조회 및 내용 읽기 |
| R14 | `read_email_attachments_all` | 이메일 1건의 첨부파일 전체 내용을 한 번의 호출로 일괄 읽기 |

### 기간계 DB ERP 도구 (erp/insa.py) — 1개

| 번호 | 도구 | 설명 |
|------|------|------|
| E1 | `get_insa_calendar` | 특정 날짜의 인사 캘린더(근무일/휴일/요일) 조회. **MS365 로그인 필요** |

> `get_insa_calendar` 응답 필드: `TP_WEEK` (1=일~7=토), `FG1_HOLIDAY` (W=평일, H=공휴일)

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

## 배포 주의사항

### ZIP 배포 시 반드시 제외
- `.venv\` — 생성 PC의 절대 경로가 박혀 있어 타 PC에서 동작 불가
- `.env` — 테넌트 ID 등 민감 정보 포함
- `__pycache__\`, `*.pyc`

### 직원 PC 설치 경로 제한
- 경로에 **한글, 공백 불가** (Python venv 경로 파싱 오류)
- 권장: `C:\royal-ms365-mcp\`, `D:\tools\royal-ms365-mcp\`

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
| `teams_channel_messages` 403 오류 | `ChannelMessage.Read.All` 관리자 동의 미완료 | Azure Portal에서 `ChannelMessage.Read.All` 권한 추가 후 재로그인 |
| 임직원 조회 결과 없음 | `User.Read.All` 관리자 동의 미완료 | Azure Portal에서 `User.Read.All` 권한 추가 |
