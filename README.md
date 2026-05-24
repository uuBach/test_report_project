# Test Report Parser

Python-скрипт для получения тестового отчёта через API, извлечения уникальных номеров телефонов, сохранения результата в Excel-файл и отправки Excel-файла на почту.

## Что делает скрипт

- выполняет GET-запрос к API;
- использует API-ключ из `.env`;
- получает URL отчёта из `.env`;
- проверяет HTTP-ошибки, ошибки подключения и невалидный JSON;
- валидирует структуру отчёта;
- извлекает уникальные номера телефонов из поля `msisdn`;
- сохраняет результат в файл `unique_phones.xlsx`;
- отправляет Excel-файл на указанную почту;
- ведёт логирование в консоль и файл `app.log`;
- отправляет Telegram-уведомления о запуске, успешном завершении и ошибках.

## Структура проекта

```text
test_report_project/
├── main.py
├── .env.example
├── .gitignore
├── README.md
└── requirements.txt
```

Файлы `.env`, `app.log`, `unique_phones.xlsx` и папка `venv/` не должны попадать в GitHub.

## Установка

Создать виртуальное окружение:

```bash
python3 -m venv venv
```

Активировать виртуальное окружение:

```bash
source venv/bin/activate
```

Установить зависимости:

```bash
pip install -r requirements.txt
```

## Настройка

Создайте файл `.env` по примеру `.env.example`:

```env
REPORT_URL=https://example.com/report
API_KEY=your_api_key_here

TELEGRAM_BOT_TOKEN=your_telegram_bot_token_here
TELEGRAM_CHAT_ID=your_telegram_chat_id_here

SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SENDER_EMAIL=your_sender_email_here
SENDER_PASSWORD=your_sender_app_password_here
RECEIVER_EMAIL=receiver_email_here
```

Важно: файл `.env` содержит секретные данные и не должен попадать в GitHub.

## Запуск

```bash
python main.py
```

## Результат

После успешного запуска скрипт:

1. получает отчёт через API;
2. извлекает уникальные номера телефонов;
3. создаёт Excel-файл:

```text
unique_phones.xlsx
```

4. отправляет этот Excel-файл на почту, указанную в `RECEIVER_EMAIL`;
5. записывает логи в файл:

```text
app.log
```

## Обработка ошибок

Скрипт явно обрабатывает основные ошибки:

- отсутствующие переменные окружения;
- невалидный API-ключ;
- ошибки подключения к API;
- timeout при запросе к API;
- HTTP-ошибки;
- невалидный JSON в ответе API;
- некорректную структуру отчёта;
- отсутствие поля `data.result`;
- отсутствие валидных номеров телефонов;
- ошибки при создании Excel-файла;
- ошибки SMTP при отправке письма;
- ошибки Telegram-уведомлений.

Если возникает ошибка, она записывается в лог и отправляется Telegram-уведомление.