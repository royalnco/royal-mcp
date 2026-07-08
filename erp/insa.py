# erp/insa.py — 기간계 DB 인사 캘린더 API 연동
import re
import logging

import httpx

from ms365.auth import require_auth

log = logging.getLogger("royal-mcp.erp-insa")

_ERP_BASE_URL = "https://order.iroyal.kr/google"


def register_erp_tools(mcp):
    """기간계 DB ERP 도구를 MCP 서버에 등록합니다."""

    @mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
    async def get_insa_calendar(ymd: str) -> dict:
        """특정 날짜의 인사 캘린더(근무일/휴일) 정보를 조회합니다.

        기간계 DB API를 호출합니다. MS365 로그인이 되어 있어야 사용할 수 있습니다.

        Args:
            ymd: 조회 날짜 (YYYYMMDD 형식, 예: 20260630)

        Returns:
            {
                "result": true,
                "data": [
                    {
                        "DT_CAL": "20260630",
                        "TP_WEEK": "3",      # 1=일 2=월 3=화 4=수 5=목 6=금 7=토
                        "FG1_HOLIDAY": "W"   # W=평일, H=공휴일
                    }
                ]
            }
        """
        auth_error = require_auth()
        if auth_error:
            return auth_error

        if not re.fullmatch(r"\d{8}", ymd):
            return {
                "success": False,
                "message": f"날짜 형식 오류: YYYYMMDD 8자리로 입력하세요. (입력값: {ymd})",
            }

        url = f"{_ERP_BASE_URL}/get_insa_calendar"
        try:
            async with httpx.AsyncClient(timeout=10.0) as client:
                resp = await client.get(url, params={"ymd": ymd})
                resp.raise_for_status()
                return resp.json()
        except httpx.TimeoutException:
            log.error("ERP API 타임아웃: %s?ymd=%s", url, ymd)
            return {"success": False, "message": "기간계 API 응답 시간이 초과되었습니다 (10초)."}
        except httpx.HTTPStatusError as e:
            log.error("ERP API HTTP 오류: %s %s", e.response.status_code, url)
            return {"success": False, "message": f"기간계 API HTTP 오류: {e.response.status_code}"}
        except Exception as e:
            log.error("ERP API 호출 실패: %s", e)
            return {"success": False, "message": f"기간계 API 호출 실패: {e}"}
