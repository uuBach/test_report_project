import logging
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from dotenv import load_dotenv
import os


OUTPUT_FILE = "unique_phones.xlsx"
LOG_FILE = "app.log"


@dataclass
class Config:
    report_url: str
    api_key: str

    telegram_token: str
    telegram_chat_id: str

    smtp_host: str
    smtp_port: int
    sender_email: str
    sender_password: str
    receiver_email: str


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )


def get_required_env(name: str) -> str:
    value = os.getenv(name)

    if value is None or not value.strip():
        raise ValueError(f"{name} is missing in .env")

    return value.strip()


def load_config() -> Config:
    load_dotenv()

    telegram_token = get_required_env("TELEGRAM_BOT_TOKEN")

    if telegram_token.startswith("bot"):
        telegram_token = telegram_token[3:]

    smtp_port_raw = get_required_env("SMTP_PORT")

    try:
        smtp_port = int(smtp_port_raw)
    except ValueError as error:
        raise ValueError("SMTP_PORT must be an integer") from error

    return Config(
        report_url=get_required_env("REPORT_URL"),
        api_key=get_required_env("API_KEY"),

        telegram_token=telegram_token,
        telegram_chat_id=get_required_env("TELEGRAM_CHAT_ID"),

        smtp_host=get_required_env("SMTP_HOST"),
        smtp_port=smtp_port,
        sender_email=get_required_env("SENDER_EMAIL"),
        sender_password=get_required_env("SENDER_PASSWORD").replace(" ", ""),
        receiver_email=get_required_env("RECEIVER_EMAIL"),
    )


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


def fetch_report(report_url: str, api_key: str) -> dict[str, Any]:
    headers = {
        "X-API-Key": api_key
    }

    try:
        response = requests.get(report_url, headers=headers, timeout=30)
    except requests.Timeout as error:
        raise ConnectionError("Report API request timed out") from error
    except requests.ConnectionError as error:
        raise ConnectionError("Failed to connect to Report API") from error
    except requests.RequestException as error:
        raise RuntimeError(f"Report API request failed: {error}") from error

    if response.status_code == 401:
        raise PermissionError("Unauthorized: invalid API key")

    try:
        response.raise_for_status()
    except requests.HTTPError as error:
        response_text = response.text[:500]
        raise RuntimeError(
            f"Report API returned HTTP {response.status_code}. Response body: {response_text}"
        ) from error

    try:
        report = response.json()
    except requests.exceptions.JSONDecodeError as error:
        response_text = response.text[:500]
        raise ValueError(
            f"Report API returned invalid JSON. Response body: {response_text}"
        ) from error

    if not isinstance(report, dict):
        raise ValueError("Invalid report format: root JSON value must be an object")

    return report


def extract_unique_phones(report: dict[str, Any]) -> list[str]:
    if "data" not in report:
        raise ValueError("Invalid report format: missing 'data' field")

    data = report["data"]

    if not isinstance(data, dict):
        raise ValueError("Invalid report format: 'data' must be an object")

    if "result" not in data:
        raise ValueError("Invalid report format: missing 'data.result' field")

    result = data["result"]

    if not isinstance(result, list):
        raise ValueError("Invalid report format: 'data.result' must be a list")

    if not result:
        raise ValueError("Invalid report format: 'data.result' is empty")

    phones: list[str] = []

    for index, item in enumerate(result):
        if not isinstance(item, dict):
            logging.warning("Skipped item at index %s: expected object, got %s", index, type(item).__name__)
            continue

        phone = item.get("msisdn")

        if phone is None:
            logging.warning("Skipped item at index %s: missing 'msisdn'", index)
            continue

        phone = str(phone).strip()

        if not phone:
            logging.warning("Skipped item at index %s: empty 'msisdn'", index)
            continue

        phones.append(phone)

    unique_phones = sorted(set(phones))

    if not unique_phones:
        raise ValueError("No valid phone numbers found in report")

    return unique_phones


def save_to_excel(phones: list[str], filename: str) -> None:
    output_path = Path(filename)

    try:
        df = pd.DataFrame({
            "phone_number": phones
        })

        df.to_excel(output_path, index=False)

    except PermissionError as error:
        raise PermissionError(
            f"Cannot write Excel file '{output_path}'. File may be open or write permission is denied."
        ) from error
    except OSError as error:
        raise OSError(
            f"Failed to write Excel file '{output_path}'. Check disk space, file path, and permissions."
        ) from error
    except ImportError as error:
        raise ImportError(
            "Failed to write Excel file: missing Excel engine. Make sure 'openpyxl' is installed."
        ) from error
    except Exception as error:
        raise RuntimeError(
            f"Unexpected error while creating Excel file '{output_path}': {error}"
        ) from error

    if not output_path.exists():
        raise FileNotFoundError(f"Excel file was not created: {output_path}")

    if output_path.stat().st_size == 0:
        raise ValueError(f"Excel file was created but it is empty: {output_path}")


def send_email_with_attachment(
    smtp_host: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str,
    receiver_email: str,
    attachment_path: str,
) -> None:
    file_path = Path(attachment_path)

    if not file_path.exists():
        raise FileNotFoundError(f"Attachment file not found: {file_path}")

    if not file_path.is_file():
        raise ValueError(f"Attachment path is not a file: {file_path}")

    if file_path.stat().st_size == 0:
        raise ValueError(f"Attachment file is empty: {file_path}")

    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Test report: unique phone numbers"

    message.set_content(
        "Здравствуйте.\n\n"
        "Во вложении Excel-файл с уникальными номерами телефонов из отчёта.\n\n"
        "Сообщение отправлено автоматически."
    )

    try:
        file_data = file_path.read_bytes()
    except OSError as error:
        raise OSError(f"Failed to read attachment file '{file_path}'") from error

    message.add_attachment(
        file_data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=file_path.name,
    )

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.starttls()
            smtp.login(sender_email, sender_password)
            smtp.send_message(message)

    except smtplib.SMTPAuthenticationError as error:
        raise PermissionError(
            "SMTP authentication failed. Check sender email and app password."
        ) from error
    except smtplib.SMTPRecipientsRefused as error:
        raise ValueError(
            f"SMTP refused receiver email: {receiver_email}"
        ) from error
    except smtplib.SMTPSenderRefused as error:
        raise ValueError(
            f"SMTP refused sender email: {sender_email}"
        ) from error
    except smtplib.SMTPException as error:
        raise RuntimeError(f"SMTP error while sending email: {error}") from error
    except OSError as error:
        raise ConnectionError(
            f"Failed to connect to SMTP server {smtp_host}:{smtp_port}"
        ) from error


def main() -> None:
    setup_logging()

    config: Config | None = None

    try:
        config = load_config()

        logging.info("Script started")
        send_telegram_message(
            config.telegram_token,
            config.telegram_chat_id,
            "Скрипт запущен"
        )

        report = fetch_report(config.report_url, config.api_key)
        logging.info("Report fetched successfully")

        phones = extract_unique_phones(report)
        logging.info("Unique phone numbers found: %s", len(phones))

        save_to_excel(phones, OUTPUT_FILE)
        logging.info("Excel file created: %s", OUTPUT_FILE)

        send_email_with_attachment(
            smtp_host=config.smtp_host,
            smtp_port=config.smtp_port,
            sender_email=config.sender_email,
            sender_password=config.sender_password,
            receiver_email=config.receiver_email,
            attachment_path=OUTPUT_FILE,
        )
        logging.info("Email with Excel attachment sent to %s", config.receiver_email)

        send_telegram_message(
            config.telegram_token,
            config.telegram_chat_id,
            (
                "Скрипт завершён успешно.\n"
                f"Уникальных номеров: {len(phones)}\n"
                f"Excel-файл: {OUTPUT_FILE}\n"
                f"Отправлено на почту: {config.receiver_email}"
            )
        )

    except Exception as error:
        logging.exception("Script failed")

        if config is not None:
            send_telegram_message(
                config.telegram_token,
                config.telegram_chat_id,
                f"Ошибка при выполнении скрипта: {error}"
            )

        raise


if __name__ == "__main__":
    main()