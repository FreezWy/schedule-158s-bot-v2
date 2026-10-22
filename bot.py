import logging
import os
import re
from datetime import datetime
from io import BytesIO
from urllib.parse import unquote, urljoin
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

SITE = "https://medcol-ptz.ru/student/studies/"
TZ = ZoneInfo("Europe/Moscow")
USER_AGENT = "Mozilla/5.0 (compatible; Schedule158sBot/1.0)"
TIMEOUT = (15, 60)

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger(__name__)

# Each day occupies four columns in the college's current workbook.
# Column B contains lesson time; C:F Monday, G:J Tuesday, etc.
DAY_BLOCKS = [
    ("Понедельник", 3, 6),
    ("Вторник", 7, 10),
    ("Среда", 11, 14),
    ("Четверг", 15, 18),
    ("Пятница", 19, 22),
    ("Суббота", 23, 26),
]
GROUP_RE = re.compile(r"^\s*158\s*с\s*$", re.IGNORECASE)
DATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")


def clean(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y %H:%M")
    return " ".join(str(value).replace("\xa0", " ").split())


def get_latest_schedule(session):
    response = session.get(SITE, timeout=TIMEOUT)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")

    candidates = []
    for anchor in soup.find_all("a", href=True):
        href = urljoin(SITE, anchor["href"])
        label = clean(anchor.get_text(" ", strip=True))
        combined = unquote(f"{label} {href}")
        path = unquote(href.split("?", 1)[0]).lower()
        if not path.endswith(".xlsx"):
            continue
        if not re.search(r"156\s*с\s*[-–—]\s*158\s*с", combined, re.IGNORECASE):
            continue

        dates = DATE_RE.findall(combined)
        if len(dates) < 2:
            log.info("Пропускаю файл без диапазона дат в ссылке: %s", label or href)
            continue
        try:
            start = datetime.strptime(dates[-2], "%d.%m.%Y").date()
            end = datetime.strptime(dates[-1], "%d.%m.%Y").date()
        except ValueError:
            continue
        candidates.append((start, end, label or href, href))

    if not candidates:
        raise RuntimeError(
            "Не нашёл на странице Excel-файл .xlsx для диапазона групп 156с–158с "
            "с двумя датами в названии/ссылке. Возможно, сайт изменил оформление."
        )

    start, end, label, url = max(candidates, key=lambda item: (item[0], item[1]))
    log.info("Выбран файл расписания: %s (%s–%s)", label, start, end)

    file_response = session.get(url, timeout=TIMEOUT)
    file_response.raise_for_status()
    content_type = file_response.headers.get("Content-Type", "").lower()
    if "html" in content_type and b"<html" in file_response.content[:500].lower():
        raise RuntimeError("Вместо Excel сайт вернул HTML-страницу.")

    try:
        workbook = load_workbook(BytesIO(file_response.content), data_only=True, read_only=False)
    except Exception as exc:
        raise RuntimeError(f"Не удалось открыть скачанный Excel-файл: {exc}") from exc

    return label, start, end, workbook


def find_group_block(ws):
    """Return (first_row, end_row) for the 158с block in column A."""
    first_row = None
    for row in range(1, ws.max_row + 1):
        if GROUP_RE.match(clean(ws.cell(row, 1).value)):
            first_row = row
            break
    if first_row is None:
        return None

    end_row = ws.max_row + 1
    for row in range(first_row + 1, ws.max_row + 1):
        value = clean(ws.cell(row, 1).value)
        if value:
            end_row = row
            break
    return first_row, end_row


def date_from_header(value):
    match = re.search(r"(\d{2}\.\d{2})(?:\.\d{4})?", clean(value))
    return match.group(1) if match else ""


def extract_group(workbook):
    for ws in workbook.worksheets:
        block = find_group_block(ws)
        if not block:
            continue

        first_row, end_row = block
        lines = []
        for day, start_col, end_col in DAY_BLOCKS:
            # Header cells are B, G, K, O, S and W in this workbook layout.
            header_col = start_col - 1 if start_col == 3 else start_col
            date_label = date_from_header(ws.cell(2, header_col).value)
            day_lines = []

            for row in range(first_row, end_row):
                time_label = clean(ws.cell(row, 2).value)
                subjects = []
                for col in range(start_col, end_col + 1):
                    value = clean(ws.cell(row, col).value)
                    if value and value not in subjects:
                        subjects.append(value)
                if subjects:
                    # Different cells can represent separate subgroups; keep each entry.
                    for subject in subjects:
                        if time_label:
                            day_lines.append(f"• {time_label}\n  {subject}")
                        else:
                            day_lines.append(f"• {subject}")

            if day_lines:
                heading = f"📅 {day}"
                if date_label:
                    heading += f", {date_label}"
                lines.append(heading + "\n" + "\n".join(day_lines))

        if not lines:
            raise RuntimeError("Блок 158с найден, но занятия в его колонках не обнаружены.")
        return "\n\n".join(lines)

    raise RuntimeError("Группа 158с не найдена в скачанном Excel-файле.")


def split_message(text, limit=3800):
    chunks, current = [], ""
    for line in text.splitlines():
        if len(line) > limit:
            if current:
                chunks.append(current)
                current = ""
            for i in range(0, len(line), limit):
                chunks.append(line[i:i + limit])
            continue
        candidate = f"{current}\n{line}" if current else line
        if len(candidate) > limit:
            if current:
                chunks.append(current)
            current = line
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks


def send_telegram(session, message):
    token = os.environ["TELEGRAM_BOT_TOKEN"].strip()
    chat_id = os.environ["TELEGRAM_CHAT_ID"].strip()
    if not token or not chat_id:
        raise RuntimeError("Не заданы TELEGRAM_BOT_TOKEN и/или TELEGRAM_CHAT_ID.")

    api_url = f"https://api.telegram.org/bot{token}/sendMessage"
    for part in split_message(message):
        response = session.post(
            api_url,
            json={"chat_id": chat_id, "text": part, "disable_web_page_preview": True},
            timeout=TIMEOUT,
        )
        response.raise_for_status()
        payload = response.json()
        if not payload.get("ok"):
            raise RuntimeError(f"Telegram API вернул ошибку: {payload}")


def main():
    # Validate secrets before doing web requests.
    if not os.getenv("TELEGRAM_BOT_TOKEN") or not os.getenv("TELEGRAM_CHAT_ID"):
        raise RuntimeError("Добавь TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID в GitHub Secrets.")

    with requests.Session() as session:
        session.headers.update({"User-Agent": USER_AGENT})
        filename, start, end, workbook = get_latest_schedule(session)
        today_date = datetime.now(TZ).date()
        if end < today_date:
            workbook.close()
            message = (
                "⚠️ НОВОЕ РАСПИСАНИЕ ПОКА НЕ НАЙДЕНО\\n"
                f"На сайте всё ещё указан файл: {filename}\\n"
                f"Период файла: {start.strftime('%d.%m.%Y')}–{end.strftime('%d.%m.%Y')}\\n"
                f"Сегодня: {today_date.strftime('%d.%m.%Y')}\\n\\n"
                "Старое расписание не отправляю, потому что его период уже закончился. "
                "Проверь страницу колледжа: " + SITE
            )
            send_telegram(session, message)
            log.warning("Не отправляю устаревшее расписание: %s (%s–%s)", filename, start, end)
            return

        try:
            schedule = extract_group(workbook)
        finally:
            workbook.close()

        today = today_date.strftime("%d.%m.%Y")
        message = (
            "🩺 РАСПИСАНИЕ ГРУППЫ 158с\n"
            f"📆 Период: {start.strftime('%d.%m.%Y')}–{end.strftime('%d.%m.%Y')}\n"
            f"🔎 Проверено: {today}\n\n"
            f"{schedule}\n\n"
            f"🌐 Официальный сайт: {SITE}"
        )
        send_telegram(session, message)
        log.info("Расписание группы 158с успешно отправлено в Telegram.")


if __name__ == "__main__":
    main()
