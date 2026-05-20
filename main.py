import os
import logging
from typing import Any

import requests
import pandas as pd
from dotenv import load_dotenv


REPORT_URL = "https://middleware-01.fromtech.kz/test_report_dev/test_report"
OUTPUT_FILE = "unique_phones.xlsx"
LOG_FILE = "app.log"


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE),
            logging.StreamHandler()
        ]
    )


def load_config() -> tuple[str, str, str]:
    load_dotenv()

    api_key = os.getenv("API_KEY")
    telegram_token = os.getenv("TELEGRAM_BOT_TOKEN")
    telegram_chat_id = os.getenv("TELEGRAM_CHAT_ID")

    if not api_key:
        raise ValueError("API_KEY is missing in .env")
    if not telegram_token:
        raise ValueError("TELEGRAM_BOT_TOKEN is missing in .env")
    if not telegram_chat_id:
        raise ValueError("TELEGRAM_CHAT_ID is missing in .env")

    if telegram_token.startswith("bot"):
        telegram_token = telegram_token[3:]

    return api_key, telegram_token, telegram_chat_id


def send_telegram_message(token: str, chat_id: str, text: str) -> None:
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text
    }

    try:
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()
        logging.info("Telegram notification sent")
    except requests.RequestException as error:
        logging.warning("Failed to send Telegram notification: %s", error)


def fetch_report(api_key: str) -> dict[str, Any]:
    headers = {
        "X-API-Key": api_key
    }

    response = requests.get(REPORT_URL, headers=headers, timeout=30)

    if response.status_code == 401:
        raise PermissionError("Unauthorized: invalid API key")

    response.raise_for_status()
    return response.json()


def extract_unique_phones(report: dict[str, Any]) -> list[str]:
    result = report.get("data", {}).get("result", [])

    if not isinstance(result, list):
        raise ValueError("Invalid report format: data.result is not a list")

    phones = []

    for item in result:
        if not isinstance(item, dict):
            continue

        phone = item.get("msisdn")

        if phone:
            phones.append(str(phone).strip())

    unique_phones = sorted(set(phones))

    return unique_phones


def save_to_excel(phones: list[str], filename: str) -> None:
    df = pd.DataFrame({
        "phone_number": phones
    })

    df.to_excel(filename, index=False)


def main() -> None:
    setup_logging()

    api_key = ""
    telegram_token = ""
    telegram_chat_id = ""

    try:
        api_key, telegram_token, telegram_chat_id = load_config()

        logging.info("Script started")
        send_telegram_message(
            telegram_token,
            telegram_chat_id,
            "Скрипт запущен"
        )

        report = fetch_report(api_key)
        logging.info("Report fetched successfully")

        phones = extract_unique_phones(report)
        logging.info("Unique phone numbers found: %s", len(phones))

        save_to_excel(phones, OUTPUT_FILE)
        logging.info("Excel file created: %s", OUTPUT_FILE)

        send_telegram_message(
            telegram_token,
            telegram_chat_id,
            f"Скрипт завершён успешно. Уникальных номеров: {len(phones)}. Файл: {OUTPUT_FILE}"
        )

    except Exception as error:
        logging.exception("Script failed")

        if telegram_token and telegram_chat_id:
            send_telegram_message(
                telegram_token,
                telegram_chat_id,
                f"Ошибка при выполнении скрипта: {error}"
            )

        raise


if __name__ == "__main__":
    main()
