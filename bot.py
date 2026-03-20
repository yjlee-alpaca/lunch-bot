import os, re, json, base64, asyncio, logging, tempfile
import httpx
from datetime import datetime, timezone, timedelta
from pathlib import Path

from slack_bolt.async_app import AsyncApp
from slack_bolt.adapter.socket_mode.async_handler import AsyncSocketModeHandler
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

SLACK_BOT_TOKEN   = os.environ["SLACK_BOT_TOKEN"]
SLACK_APP_TOKEN   = os.environ["SLACK_APP_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
HIWORKS_EMAIL     = os.environ["HIWORKS_EMAIL"]
HIWORKS_PASSWORD  = os.environ["HIWORKS_PASSWORD"]
MY_SLACK_USER_ID  = os.environ["MY_SLACK_USER_ID"]
HIWORKS_COMPANY_ID = os.environ.get("HIWORKS_COMPANY_ID", "")

KST = timezone(timedelta(hours=9))
MEAL_ALLOWANCE = 11000
pending = {}

app = AsyncApp(token=SLACK_BOT_TOKEN)

# ── 유틸 ──────────────────────────────────────────────
def format_date(dt):
    return dt.strftime("%Y-%m-%d")

def calc_claim(amount, companions):
    headcount = len(companions) + 1
    per_person = round(amount / headcount)
    claim = min(per_person, MEAL_ALLOWANCE)
    return per_person, claim

async def download_slack_file(file):
    url = file.get("url_private_download") or file.get("url_private")
    ext = Path(file.get("name", "receipt.png")).suffix or ".png"
    async with httpx.AsyncClient(timeout=60) as http:
        resp = await http.get(url, headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"}, follow_redirects=True)
        resp.raise_for_status()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(resp.content)
    tmp.close()
    log.info(f"파일 다운로드 완료: {len(resp.content)} bytes, mime: {file.get('mimetype')}")
    return Path(tmp.name)

async def resize_image_if_needed(file_path, mime):
    """이미지가 2MB 이상이면 리사이즈"""
    if mime == "application/pdf":
        return file_path
    try:
        from PIL import Image
        import io
        size = os.path.getsize(file_path)
        if size <= 2 * 1024 * 1024:
            return file_path
        log.info(f"이미지 리사이즈 필요: {size} bytes")
        img = Image.open(file_path)
        # 최대 1600px로 줄이기
        img.thumbnail((1600, 1600), Image.LANCZOS)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        img.convert("RGB").save(tmp.name, "JPEG", quality=85)
        tmp.close()
        new_size = os.path.getsize(tmp.name)
        log.info(f"리사이즈 완료: {new_size} bytes")
        return Path(tmp.name)
    except Exception as e:
        log.warning(f"리사이즈 실패, 원본 사용: {e}")
        return file_path

async def analyze_receipt(file_path, mime, fallback_date):
    # 항상 JPEG로 변환해서 API에 전송
    try:
        from PIL import Image as PILImage
        img = PILImage.open(file_path)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        img.convert("RGB").save(tmp.name, "JPEG", quality=85)
        tmp.close()
        final_path = Path(tmp.name)
        api_mime = "image/jpeg"
        log.info(f"JPEG 변환 완료: {os.path.getsize(final_path)} bytes")
    except Exception as e:
        log.warning(f"JPEG 변환 실패: {e}, 원본 사용")
        final_path = file_path
        api_mime = "image/jpeg" if not mime == "application/pdf" else "application/pdf"

    with open(final_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode()

    if api_mime == "application/pdf":
        content_block = {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": data}}
    else:
        content_block = {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": data}}

    prompt = '이 영수증에서 정보를 추출해주세요. JSON만 응답하세요: {"amount": 숫자, "store_name": "가게명", "date": "YYYY-MM-DD"} 날짜 없으면 ' + fallback_date

    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": "claude-opus-4-5",
                "max_tokens": 256,
                "messages": [{"role": "user", "content": [content_block, {"type": "text", "text": prompt}]}]
            }
        )
    log.info(f"Claude API 응답: {resp.status_code}")
    resp.raise_for_status()
    text = resp.json()["content"][0]["text"]
    log.info(f"Claude 응답 내용: {text}")
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        raise ValueError("영수증 정보를 읽지 못했어요")
    return json.loads(match.group())

# ── 스케줄러: 12:30 알림 ──────────────────────────────
async def send_lunch_reminder():
    now = datetime.now(KST)
    if now.weekday() >= 5:
        return
    date_str = format_date(now)
    await app.client.chat_postMessage(
        channel=MY_SLACK_USER_ID,
        text=f"점심 드셨나요? ({date_str})\n영수증 사진을 올려주시면 입체금 청구를 자동으로 처리해드릴게요!\nPNG, JPG, PDF 모두 가능해요."
    )
    log.info("점심 알림 전송 완료")

# ── 파일 처리 공통 함수 ───────────────────────────────
async def process_file(file, channel, client_obj):
    mime = file.get("mimetype", "")
    if not (mime.startswith("image/") or mime == "application/pdf"):
        await client_obj.chat_postMessage(channel=channel, text="PNG, JPG, PDF 파일만 지원해요.")
        return

    await client_obj.chat_postMessage(channel=channel, text="영수증 분석 중... 잠깐만요!")
    today = format_date(datetime.now(KST))

    try:
        tmp_path = await download_slack_file(file)
        result = await analyze_receipt(tmp_path, mime, today)
    except Exception as e:
        log.error(f"영수증 분석 실패: {e}")
        await client_obj.chat_postMessage(channel=channel, text=f"영수증 읽기 실패: {e}\n금액을 직접 입력해주세요 (예: 15000)")
        return

    amount = result["amount"]
    store_name = result.get("store_name", "")
    date = result.get("date", today)
    log.info(f"분석 완료: {store_name} {amount}원 {date}")

    if amount >= 20000:
        pending[channel] = {"amount": amount, "store_name": store_name, "date": date, "file_path": str(tmp_path), "mime": mime}
        await client_obj.chat_postMessage(
            channel=channel,
            text=f"금액이 {amount:,}원이네요! 여러 명이 드셨나요?\n함께 드신 분 이름을 입력해주세요 (예: 김철수, 이영희)\n혼자 드셨으면 '혼자' 라고 입력해주세요."
        )
    else:
        await do_process_expense(client_obj, channel, amount, store_name, date, str(tmp_path), mime, [])

async def do_process_expense(client_obj, channel, amount, store_name, date, file_path, mime, companions):
    companion_text = ", ".join(companions) if companions else "없음"
    per_person, claim = calc_claim(amount, companions)
    headcount = len(companions) + 1

    detail = ""
    if headcount > 1:
        detail = f"\n- 1인당 실제: {per_person:,}원 ({amount:,}원 / {headcount}명)"
    if claim < per_person:
        detail += f"\n- 청구 금액: {claim:,}원 (한도 적용)"
    else:
        detail += f"\n- 청구 금액: {claim:,}원"

    await client_obj.chat_postMessage(
        channel=channel,
        text=f"입력 정보 확인\n- 날짜: {date}\n- 사용처: {store_name or '(영수증 참고)'}\n- 영수증 금액: {amount:,}원\n- 동행자: {companion_text}{detail}\n\n하이웍스에 자동 제출 중..."
    )

    try:
        await submit_to_hiworks(date=date, amount=claim, store_name=store_name, companions=companions, file_path=file_path)
        await client_obj.chat_postMessage(channel=channel, text="하이웍스 입체금 청구 완료! 결재 올라갔어요.")
    except Exception as e:
        log.error(f"하이웍스 제출 실패: {e}")
        await client_obj.chat_postMessage(
            channel=channel,
            text=f"자동 제출 실패: {e}\n\n아래 내용으로 직접 입력해주세요:\n날짜: {date}\n사용처: {store_name}\n금액: {claim:,}원\n동행자: {companion_text}"
        )

# ── 이벤트 핸들러 ─────────────────────────────────────
@app.event("message")
async def handle_message(event, client):
    log.info(f"message 이벤트: channel_type={event.get('channel_type')}, subtype={event.get('subtype')}, files={bool(event.get('files'))}")
    if event.get("bot_id"):
        return

    channel = event["channel"]
    channel_type = event.get("channel_type", "")
    subtype = event.get("subtype", "")
    files = event.get("files", [])
    text = event.get("text", "").strip()

    if channel_type != "im":
        return

    if subtype and subtype != "file_share":
        return

    if files:
        await process_file(files[0], channel, client)
    elif text and channel in pending:
        data = pending.pop(channel)
        companions = [] if text == "혼자" else [s.strip() for s in re.split(r"[,，、\s]+", text) if s.strip()]
        await do_process_expense(client, channel, data["amount"], data["store_name"], data["date"], data["file_path"], data["mime"], companions)

@app.event("file_shared")
async def handle_file_shared(event, client):
    log.info(f"file_shared 이벤트: {event}")
    file_id = event.get("file_id") or event.get("file", {}).get("id")
    channel = event.get("channel_id")
    if not file_id or not channel:
        return
    try:
        file_info = await client.files_info(file=file_id)
        await process_file(file_info["file"], channel, client)
    except Exception as e:
        log.error(f"file_shared 처리 실패: {e}")
        await client.chat_postMessage(channel=channel, text=f"파일 처리 실패: {e}")

# ── 하이웍스 자동 제출 ────────────────────────────────
async def submit_to_hiworks(date, amount, store_name, companions, file_path):
    from playwright.async_api import async_playwright
    companion_text = ", ".join(companions) if companions else ""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        page = await browser.new_page()
        try:
            log.info("하이웍스 로그인 중...")
            await page.goto("https://office.hiworks.com/", wait_until="networkidle")
            if HIWORKS_COMPANY_ID:
                domain_input = page.locator("input[name='company_id'], input[placeholder*='도메인']")
                if await domain_input.count() > 0:
                    await domain_input.first.fill(HIWORKS_COMPANY_ID)
                    await page.keyboard.press("Enter")
                    await page.wait_for_load_state("networkidle")
            await page.locator("input[type='email'], input[name='user_id'], input[name='email']").first.fill(HIWORKS_EMAIL)
            await page.locator("input[type='password']").first.fill(HIWORKS_PASSWORD)
            await page.keyboard.press("Enter")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(2)

            await page.goto("https://approval.office.hiworks.com/approval/document/write", wait_until="networkidle")
            await asyncio.sleep(2)

            field_values = {
                "사용일시": date, "사용금액": str(amount),
                "비   목": "점심식대", "비목": "점심식대",
                "사   유": "복리후생", "사유": "복리후생",
                "동 행 자": companion_text, "동행자": companion_text,
            }
            rows = page.locator("table tr")
            for i in range(await rows.count()):
                tds = rows.nth(i).locator("td")
                if await tds.count() >= 2:
                    label = (await tds.nth(0).inner_text()).strip()
                    for key, val in field_values.items():
                        if key in label and val:
                            await tds.nth(1).click()
                            await asyncio.sleep(0.2)
                            await page.keyboard.press("Control+A")
                            await page.keyboard.type(val)
                            break

            file_input = page.locator("input[type='file']")
            if await file_input.count() > 0:
                await file_input.first.set_input_files(file_path)
                await asyncio.sleep(2)

            submit_btn = page.locator("button:has-text('기안하기'), a:has-text('기안하기')").first
            await submit_btn.click()
            await asyncio.sleep(2)
            confirm_btn = page.locator("button:has-text('확인'), button:has-text('기안')").first
            if await confirm_btn.count() > 0:
                await confirm_btn.click()
                await asyncio.sleep(2)
            log.info("하이웍스 제출 완료!")
        finally:
            await browser.close()

# ── 메인 ──────────────────────────────────────────────
async def main():
    scheduler = AsyncIOScheduler(timezone="Asia/Seoul")
    scheduler.add_job(send_lunch_reminder, CronTrigger(hour=12, minute=30, day_of_week="mon-fri", timezone="Asia/Seoul"))
    scheduler.start()
    log.info("스케줄러 시작 (매일 월~금 12:30 KST)")
    handler = AsyncSocketModeHandler(app, SLACK_APP_TOKEN)
    log.info("슬랙봇 시작!")
    await handler.start_async()

if __name__ == "__main__":
    asyncio.run(main())
