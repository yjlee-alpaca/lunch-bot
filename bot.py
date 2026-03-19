"""
점심 입체금 자동 청구 슬랙봇
- 매일 12:30 KST에 슬랙 DM 알림
- 영수증 사진(PNG/JPG/PDF) 받으면 Claude Vision으로 금액/날짜 추출
- 2만원 이상이면 동행자 확인
- Playwright로 하이웍스 자동 로그인 → 폼 작성 → 영수증 첨부 → 제출
"""

import os
import re
import json
import base64
import asyncio
import logging
import tempfile
import httpx
from datetime import datetime, timezone, timedelta
from pathlib import Path

from slack_bolt.async_app import AsyncApp
from slack_bolt.adapter.socket_mode.async_handler import AsyncSocketModeHandler
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
from playwright.async_api import async_playwright

# ── 로깅 ──────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── 환경변수 ───────────────────────────────────────────
SLACK_BOT_TOKEN      = os.environ["SLACK_BOT_TOKEN"]       # xoxb-...
SLACK_APP_TOKEN      = os.environ["SLACK_APP_TOKEN"]       # xapp-... (Socket Mode)
ANTHROPIC_API_KEY    = os.environ["ANTHROPIC_API_KEY"]     # sk-ant-...
HIWORKS_EMAIL        = os.environ["HIWORKS_EMAIL"]         # 하이웍스 로그인 이메일
HIWORKS_PASSWORD     = os.environ["HIWORKS_PASSWORD"]      # 하이웍스 비밀번호
MY_SLACK_USER_ID     = os.environ["MY_SLACK_USER_ID"]      # 슬랙 유저 ID (U로 시작)
HIWORKS_COMPANY_ID   = os.environ.get("HIWORKS_COMPANY_ID", "")  # 회사 도메인 (선택)

KST = timezone(timedelta(hours=9))

# ── 임시 상태 저장 (메모리) ────────────────────────────
# { channel_id: { amount, store_name, date, file_path } }
pending: dict = {}

# ── 슬랙 앱 초기화 ────────────────────────────────────
app = AsyncApp(token=SLACK_BOT_TOKEN)


# ══════════════════════════════════════════════════════
# 1. 스케줄러: 매일 12:30 KST 알림
# ══════════════════════════════════════════════════════
async def send_lunch_reminder():
    now = datetime.now(KST)
    if now.weekday() >= 5:          # 토·일 제외
        return

    date_str = now.strftime("%Y-%m-%d")
    client = app.client
    try:
        await client.chat_postMessage(
            channel=MY_SLACK_USER_ID,
            text=(
                f"🍱 점심 드셨나요? ({date_str})\n"
                "영수증 사진을 올려주시면 입체금 청구를 자동으로 처리해드릴게요!\n"
                "PNG, JPG, PDF 모두 가능해요."
            ),
        )
        log.info("점심 알림 전송 완료")
    except Exception as e:
        log.error(f"알림 전송 실패: {e}")


# ══════════════════════════════════════════════════════
# 2. 슬랙 이벤트: 파일 업로드 처리
# ══════════════════════════════════════════════════════
@app.event("message.im")
async def handle_message(event, say, client):
    # 봇 메시지 무시
    if event.get("bot_id") or event.get("subtype"):
        return

    channel = event["channel"]
    user    = event.get("user", "")
    files   = event.get("files", [])
    text    = event.get("text", "").strip()

    # DM만 처리 (슬랙 DM 채널 ID는 D로 시작)
    if not channel.startswith("D"):
        return

    # ── 파일 첨부 ──
    if files:
        await say("📎 영수증 분석 중... 잠깐만요! 🔍")
        file = files[0]
        mime = file.get("mimetype", "")

        if not (mime.startswith("image/") or mime == "application/pdf"):
            await say("❌ PNG, JPG, PDF 파일만 지원해요. 다시 올려주세요!")
            return

        # 파일 다운로드
        try:
            tmp_path = await download_slack_file(client, file)
        except Exception as e:
            await say(f"❌ 파일 다운로드 실패: {e}")
            return

        # Claude Vision으로 분석
        today = datetime.now(KST).strftime("%Y-%m-%d")
        try:
            result = await analyze_receipt(tmp_path, mime, today)
        except Exception as e:
            await say(f"❌ 영수증 읽기 실패: {e}\n금액을 직접 입력해주세요 (예: `15000`)")
            return

        amount     = result["amount"]
        store_name = result.get("store_name", "")
        date       = result.get("date", today)

        log.info(f"영수증 분석 완료: {store_name} {amount}원 {date}")

        if amount >= 20000:
            # 동행자 확인 대기
            pending[channel] = {
                "amount": amount,
                "store_name": store_name,
                "date": date,
                "file_path": str(tmp_path),
                "mime": mime,
            }
            await say(
                f"💰 *{amount:,}원* 이네요! 여러 명이 드셨나요?\n\n"
                "함께 드신 분 이름을 입력해주세요 _(예: 김철수, 이영희)_\n"
                "혼자 드셨으면 `혼자` 라고 입력해주세요."
            )
        else:
            await process_expense(say, channel, amount, store_name, date, str(tmp_path), mime, [])

    # ── 텍스트 응답 (동행자 입력) ──
    elif text and channel in pending:
        data = pending.pop(channel)
        companions = []
        if text != "혼자":
            companions = [s.strip() for s in re.split(r"[,，、\s]+", text) if s.strip()]

        await process_expense(
            say,
            channel,
            data["amount"],
            data["store_name"],
            data["date"],
            data["file_path"],
            data["mime"],
            companions,
        )


# ══════════════════════════════════════════════════════
# 3. 하이웍스 자동 제출
# ══════════════════════════════════════════════════════
MEAL_ALLOWANCE = 11_000  # 1인당 식대 한도


def calc_claim_amount(amount: int, companions: list) -> tuple[int, int]:
    """
    실제 청구 금액 계산
    - 1인당 금액 = 총액 / 인원수
    - 청구액 = min(1인당 금액, 11,000)
    반환: (1인당 실제 금액, 최종 청구 금액)
    """
    headcount = len(companions) + 1  # 본인 포함
    per_person = round(amount / headcount)
    claim = min(per_person, MEAL_ALLOWANCE)
    return per_person, claim


async def process_expense(say, channel, amount, store_name, date, file_path, mime, companions):
    companion_text = ", ".join(companions) if companions else "없음"
    per_person, claim = calc_claim_amount(amount, companions)
    headcount = len(companions) + 1

    # 금액 안내 메시지 구성
    amount_detail = ""
    if headcount > 1:
        amount_detail = f"\n• 1인당 실제: *{per_person:,}원* ({amount:,}원 ÷ {headcount}명)"
    if claim < per_person:
        amount_detail += f"\n• 청구 금액: *{claim:,}원* (한도 적용)"
    else:
        amount_detail += f"\n• 청구 금액: *{claim:,}원*"

    await say(
        f"✅ 입력 정보 확인\n"
        f"• 날짜: *{date}*\n"
        f"• 사용처: *{store_name or '(영수증 참고)'}*\n"
        f"• 영수증 금액: *{amount:,}원*\n"
        f"• 동행자: *{companion_text}*"
        f"{amount_detail}\n\n"
        "⏳ 하이웍스에 자동 제출 중..."
    )

    try:
        await submit_to_hiworks(
            date=date,
            amount=claim,          # 한도 적용된 청구 금액
            store_name=store_name,
            companions=companions,
            file_path=file_path,
        )
        await say("🎉 하이웍스 입체금 청구 완료! 결재 올라갔어요.")
    except Exception as e:
        log.error(f"하이웍스 제출 실패: {e}")
        await say(
            f"❌ 자동 제출 실패: `{e}`\n\n"
            "아래 내용으로 직접 입력해주세요:\n"
            f"```\n날짜: {date}\n사용처: {store_name}\n금액: {claim:,}원\n동행자: {companion_text}\n```"
        )


async def submit_to_hiworks(date, amount, store_name, companions, file_path):
    """Playwright로 하이웍스 입체금 청구 자동 제출"""
    companion_text = ", ".join(companions) if companions else ""
    reason = f"점심식대 ({store_name})"
    if companion_text:
        reason += f" - {companion_text} 포함"

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,  # 로컬 테스트 시 False로 바꾸면 브라우저 보임
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )
        page = await browser.new_page()

        try:
            # ── 로그인 ──
            log.info("하이웍스 로그인 중...")
            await page.goto("https://office.hiworks.com/", wait_until="networkidle")

            # 회사 도메인 입력 (있는 경우)
            if HIWORKS_COMPANY_ID:
                domain_input = page.locator("input[name='company_id'], input[placeholder*='도메인'], input[placeholder*='회사']")
                if await domain_input.count() > 0:
                    await domain_input.first.fill(HIWORKS_COMPANY_ID)
                    await page.keyboard.press("Enter")
                    await page.wait_for_load_state("networkidle")

            # 이메일/아이디
            await page.locator("input[type='email'], input[name='user_id'], input[name='email']").first.fill(HIWORKS_EMAIL)
            # 비밀번호
            await page.locator("input[type='password']").first.fill(HIWORKS_PASSWORD)
            await page.keyboard.press("Enter")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(2)

            log.info(f"현재 URL: {page.url}")

            # ── 전자결재 > 입체금 청구 페이지 이동 ──
            log.info("입체금 청구 페이지 이동 중...")
            await page.goto(
                "https://approval.office.hiworks.com/approval/document/write",
                wait_until="networkidle",
            )
            await asyncio.sleep(2)

            # 문서 종류: 입체금 청구 선택
            # (이미 기본값이 '입체금 청구'로 설정된 경우 스킵)
            doc_type_select = page.locator("select").first
            if await doc_type_select.count() > 0:
                try:
                    await doc_type_select.select_option(label="입체금 청구")
                except Exception:
                    pass  # 이미 선택된 경우

            await asyncio.sleep(1)

            # ── 본문 테이블 채우기 ──
            log.info("폼 입력 중...")

            # 테이블 셀 찾기 - 레이블 옆 입력칸
            # 하이웍스 에디터는 contenteditable 영역 안의 테이블
            editor = page.locator(".fr-element, [contenteditable='true']").first
            await editor.click()

            # 테이블 셀을 직접 찾아서 채우기
            cells = page.locator("table td")
            cell_count = await cells.count()
            log.info(f"테이블 셀 수: {cell_count}")

            # 각 행의 두 번째 td (입력칸)에 값 입력
            # 순서: 사용방법(고정), 사용일시, 사용금액, 비목, 사유, 동행자, 관련사업
            field_values = {
                "사용일시": date,
                "사용금액": str(amount),
                "비   목": "점심식대",
                "비목":    "점심식대",
                "사   유": "복리후생",
                "사유":    "복리후생",
                "동 행 자": companion_text,
                "동행자":   companion_text,
            }

            # 테이블 rows 순회
            rows = page.locator("table tr")
            row_count = await rows.count()
            for i in range(row_count):
                row = rows.nth(i)
                tds = row.locator("td")
                td_count = await tds.count()
                if td_count >= 2:
                    label = (await tds.nth(0).inner_text()).strip()
                    for key, val in field_values.items():
                        if key in label and val:
                            input_cell = tds.nth(1)
                            await input_cell.click()
                            await asyncio.sleep(0.3)
                            await page.keyboard.press("Control+A")
                            await page.keyboard.type(val)
                            log.info(f"  {label} → {val}")
                            break

            # ── 파일 첨부 ──
            log.info("파일 첨부 중...")
            file_input = page.locator("input[type='file']")
            if await file_input.count() > 0:
                await file_input.first.set_input_files(file_path)
                await asyncio.sleep(2)
            else:
                # 파일 첨부 버튼 클릭 방식
                attach_btn = page.locator("text=파일 첨부, a:has-text('파일'), button:has-text('파일')").first
                if await attach_btn.count() > 0:
                    async with page.expect_file_chooser() as fc_info:
                        await attach_btn.click()
                    file_chooser = await fc_info.value
                    await file_chooser.set_files(file_path)
                    await asyncio.sleep(2)

            # ── 기안하기 (제출) ──
            log.info("기안하기 클릭...")
            submit_btn = page.locator("button:has-text('기안하기'), a:has-text('기안하기')").first
            await submit_btn.click()
            await asyncio.sleep(2)

            # 확인 팝업 처리
            confirm_btn = page.locator("button:has-text('확인'), button:has-text('기안'), button:has-text('제출')").first
            if await confirm_btn.count() > 0:
                await confirm_btn.click()
                await asyncio.sleep(2)

            log.info("✅ 하이웍스 제출 완료!")

        finally:
            await browser.close()


# ══════════════════════════════════════════════════════
# 4. Claude Vision으로 영수증 분석
# ══════════════════════════════════════════════════════
async def analyze_receipt(file_path: Path, mime: str, fallback_date: str) -> dict:
    """Claude API로 영수증에서 금액/날짜/가게명 추출"""
    with open(file_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode()

    if mime == "application/pdf":
        content_block = {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": data},
        }
    else:
        content_block = {
            "type": "image",
            "source": {"type": "base64", "media_type": mime, "data": data},
        }

    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-opus-4-5",
                "max_tokens": 256,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            content_block,
                            {
                                "type": "text",
                                "text": (
                                    f"이 영수증에서 정보를 추출해주세요. JSON만 응답하세요:\n"
                                    f'{{"amount": 숫자(최종결제금액 원단위), '
                                    f'"store_name": "가게명", '
                                    f'"date": "YYYY-MM-DD"}}\n'
                                    f"날짜 없으면 {fallback_date} 사용."
                                ),
                            },
                        ],
                    }
                ],
            },
        )

    resp.raise_for_status()
    text = resp.json()["content"][0]["text"]
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("영수증에서 정보를 읽지 못했어요")
    return json.loads(match.group())


# ══════════════════════════════════════════════════════
# 5. 슬랙 파일 다운로드
# ══════════════════════════════════════════════════════
async def download_slack_file(client, file: dict) -> Path:
    url = file.get("url_private_download") or file.get("url_private")
    ext = Path(file.get("name", "receipt.png")).suffix or ".png"

    async with httpx.AsyncClient(timeout=30) as http:
        resp = await http.get(
            url,
            headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
        )
        resp.raise_for_status()

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(resp.content)
    tmp.close()
    return Path(tmp.name)


# ══════════════════════════════════════════════════════
# 6. 진입점
# ══════════════════════════════════════════════════════
async def main():
    # 스케줄러 설정 (매일 12:30 KST = 03:30 UTC)
    scheduler = AsyncIOScheduler(timezone="Asia/Seoul")
    scheduler.add_job(
        send_lunch_reminder,
        CronTrigger(hour=12, minute=30, day_of_week="mon-fri"),
    )
    scheduler.start()
    log.info("스케줄러 시작 (매일 월~금 12:30 KST)")

    # Socket Mode로 슬랙봇 실행
    handler = AsyncSocketModeHandler(app, SLACK_APP_TOKEN)
    log.info("슬랙봇 시작!")
    await handler.start_async()


if __name__ == "__main__":
    asyncio.run(main())
