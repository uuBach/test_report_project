# Test Report Parser

Python-скрипт для получения тестового отчёта через API, извлечения уникальных номеров телефонов и сохранения результата в Excel-файл.

## Что делает скрипт

- выполняет GET-запрос к API;
- использует API-ключ из `.env`;
- извлекает уникальные номера телефонов из поля `msisdn`;
- сохраняет результат в файл `unique_phones.xlsx`;
- ведёт логирование в консоль и файл `app.log`;
- отправляет Telegram-уведомления о запуске, успешном завершении и ошибках.

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
API_KEY=your_api_key_here
TELEGRAM_BOT_TOKEN=your_telegram_bot_token_here
TELEGRAM_CHAT_ID=your_telegram_chat_id_here
```

Важно: файл `.env` не должен попадать в GitHub.

## Запуск

```bash
python main.py
```

## Результат

После успешного запуска будет создан файл:

```text
unique_phones.xlsx
```

Логи сохраняются в файл:

```text
app.log
```
