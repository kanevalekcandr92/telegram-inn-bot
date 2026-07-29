#!/usr/bin/env python3
"""
Telegram Webhook Server для обработки запросов по ИНН
Интеграция с Giga Cowork для генерации документов

Автор: Giga Cowork AI Agent
Дата: 2026-07-29
Версия: 1.0
"""

import os
import json
import logging
import asyncio
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any

from flask import Flask, request, jsonify
import requests
import pandas as pd

# ==================== КОНФИГУРАЦИЯ ====================

# Токен бота от BotFather
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "ВАШ_ТОКЕН_ОТ_BOTFATHER")

# URL для webhook (ваш сервер)
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "https://your-server.pythonanywhere.com")

# Giga Cowork API (если есть HTTP API для вызова агента)
GIGA_COWORK_API_URL = os.getenv("GIGA_COWORK_API_URL", "")
GIGA_COWORK_API_KEY = os.getenv("GIGA_COWORK_API_KEY", "")

# Пути к файлам
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
LOGS_DIR = BASE_DIR / "logs"

# Создаём директории
DATA_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# Файлы
HISTORY_FILE = DATA_DIR / "dialog_history.xlsx"
ERROR_LOG = LOGS_DIR / "error.log"
ACCESS_LOG = LOGS_DIR / "access.log"

# ==================== ЛОГИРОВАНИЕ ====================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(ACCESS_LOG, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

error_logger = logging.getLogger('errors')
error_logger.addHandler(logging.FileHandler(ERROR_LOG, encoding='utf-8'))
error_logger.setLevel(logging.ERROR)

# ==================== ВАЛИДАЦИЯ ИНН ====================

def validate_inn(inn: str) -> tuple[bool, str]:
    """
    Проверка ИНН на корректность
    
    Args:
        inn: ИНН для проверки
        
    Returns:
        (True, "ОК") или (False, "Ошибка")
    """
    inn = str(inn).strip()
    
    # Проверка длины
    if len(inn) not in [10, 12]:
        return False, "ИНН должен содержать 10 или 12 цифр"
    
    # Проверка на цифры
    if not inn.isdigit():
        return False, "ИНН должен содержать только цифры"
    
    # Проверка контрольных сумм
    if len(inn) == 10:
        # ИНН организации
        coeffs = [2, 4, 10, 3, 5, 9, 4, 6, 8]
        checksum = sum(int(inn[i]) * coeffs[i] for i in range(9)) % 11
        checksum = checksum % 10
        if checksum != int(inn[9]):
            return False, "Некорректный ИНН организации (контрольная сумма)"
    else:
        # ИНН физического лица
        coeffs1 = [7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        coeffs2 = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]
        
        checksum1 = sum(int(inn[i]) * coeffs1[i] for i in range(10)) % 11
        checksum1 = checksum1 % 10
        
        checksum2 = sum(int(inn[i]) * coeffs2[i] for i in range(11)) % 11
        checksum2 = checksum2 % 10
        
        if checksum1 != int(inn[10]) or checksum2 != int(inn[11]):
            return False, "Некорректный ИНН физического лица (контрольная сумма)"
    
    return True, "ИНН корректен"


# ==================== ГЕНЕРАЦИЯ HTML ====================

def generate_html_report(inn: str, client_name: str, request_date: str, inn_type: str) -> str:
    """
    Генерация HTML-отчёта по ИНН
    
    Args:
        inn: ИНН клиента
        client_name: Имя клиента
        request_date: Дата запроса
        inn_type: Тип ИНН (Физ. лицо / Организация)
        
    Returns:
        HTML-строка
    """
    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Отчёт по ИНН {inn}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        .container {{
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{
            font-size: 28px;
            margin-bottom: 10px;
        }}
        .header p {{
            font-size: 14px;
            opacity: 0.9;
        }}
        .content {{
            padding: 40px;
        }}
        .info-block {{
            background: #f8f9fa;
            border-left: 4px solid #4CAF50;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 0 8px 8px 0;
        }}
        .info-block h3 {{
            color: #4CAF50;
            margin-bottom: 15px;
            font-size: 18px;
        }}
        .info-row {{
            display: flex;
            justify-content: space-between;
            padding: 10px 0;
            border-bottom: 1px solid #e0e0e0;
        }}
        .info-row:last-child {{
            border-bottom: none;
        }}
        .info-label {{
            font-weight: 600;
            color: #555;
        }}
        .info-value {{
            color: #333;
            font-weight: 500;
        }}
        .status-badge {{
            display: inline-block;
            padding: 6px 12px;
            background: #4CAF50;
            color: white;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }}
        .footer {{
            background: #f8f9fa;
            padding: 20px 40px;
            text-align: center;
            font-size: 12px;
            color: #888;
            border-top: 1px solid #e0e0e0;
        }}
        .footer p {{
            margin: 5px 0;
        }}
        .watermark {{
            opacity: 0.1;
            font-size: 48px;
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%) rotate(-45deg);
            pointer-events: none;
        }}
        @media print {{
            body {{
                background: white;
                padding: 0;
            }}
            .container {{
                box-shadow: none;
            }}
        }}
        @media (max-width: 600px) {{
            .content {{
                padding: 20px;
            }}
            .info-row {{
                flex-direction: column;
                gap: 5px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="watermark">GIGA COWORK</div>
        
        <div class="header">
            <h1>📊 Отчёт по ИНН</h1>
            <p>Автоматически сгенерированный документ</p>
        </div>
        
        <div class="content">
            <div class="info-block">
                <h3>📋 Основная информация</h3>
                <div class="info-row">
                    <span class="info-label">ИНН:</span>
                    <span class="info-value">{inn}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Тип:</span>
                    <span class="info-value">{inn_type}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Статус проверки:</span>
                    <span class="status-badge">✅ Проверен</span>
                </div>
            </div>
            
            <div class="info-block">
                <h3>👤 Данные клиента</h3>
                <div class="info-row">
                    <span class="info-label">Имя:</span>
                    <span class="info-value">{client_name}</span>
                </div>
                <div class="info-row">
                    <span class="info-label">Дата запроса:</span>
                    <span class="info-value">{request_date}</span>
                </div>
            </div>
            
            <div class="info-block">
                <h3>ℹ️ Примечание</h3>
                <p style="color: #666; font-size: 14px;">
                    Данный отчёт сгенерирован автоматически на основе предоставленных данных.
                    Для получения дополнительной информации обратитесь в службу поддержки.
                </p>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>Giga Cowork</strong> — Платформа автоматизации бизнеса</p>
            <p>Документ сгенерирован: {request_date}</p>
            <p>© 2026 Все права защищены</p>
        </div>
    </div>
</body>
</html>"""
    
    return html


# ==================== TELEGRAM API ====================

def send_telegram_message(chat_id: int, text: str, parse_mode: str = "HTML") -> bool:
    """
    Отправка текстового сообщения в Telegram
    
    Args:
        chat_id: ID чата
        text: Текст сообщения
        parse_mode: Режим парсинга (HTML/Markdown)
        
    Returns:
        True если успешно
    """
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    data = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": parse_mode
    }
    
    try:
        response = requests.post(url, json=data, timeout=10)
        result = response.json()
        
        if result.get("ok"):
            logger.info(f"Сообщение отправлено в чат {chat_id}")
            return True
        else:
            error_logger.error(f"Ошибка отправки сообщения: {result}")
            return False
            
    except Exception as e:
        error_logger.error(f"Исключение при отправке сообщения: {e}")
        return False


def send_telegram_document(chat_id: int, file_path: str, caption: str = "") -> bool:
    """
    Отправка документа в Telegram
    
    Args:
        chat_id: ID чата
        file_path: Путь к файлу
        caption: Подпись к документу
        
    Returns:
        True если успешно
    """
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    
    try:
        with open(file_path, 'rb') as f:
            files = {'document': f}
            data = {
                'chat_id': chat_id,
                'caption': caption,
                'parse_mode': 'HTML'
            }
            response = requests.post(url, files=files, data=data, timeout=30)
            result = response.json()
            
            if result.get("ok"):
                logger.info(f"Документ отправлен в чат {chat_id}")
                return True
            else:
                error_logger.error(f"Ошибка отправки документа: {result}")
                return False
                
    except Exception as e:
        error_logger.error(f"Исключение при отправке документа: {e}")
        return False


# ==================== СОХРАНЕНИЕ ИСТОРИИ ====================

def save_to_history(data: Dict[str, Any]) -> None:
    """
    Сохранение данных диалога в Excel-файл
    
    Args:
        data: Данные для сохранения
    """
    df_new = pd.DataFrame([data])
    
    try:
        if HISTORY_FILE.exists():
            df_existing = pd.read_excel(HISTORY_FILE)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        else:
            df_combined = df_new
        
        df_combined.to_excel(HISTORY_FILE, index=False)
        logger.info(f"История сохранена в {HISTORY_FILE}")
        
    except Exception as e:
        error_logger.error(f"Ошибка сохранения истории: {e}")


# ==================== ОБРАБОТКА ЗАПРОСА ====================

def process_inn_request(update: Dict[str, Any]) -> None:
    """
    Обработка запроса с ИНН
    
    Args:
        update: Данные обновления от Telegram
    """
    message = update.get('message', {})
    chat_id = message.get('chat', {}).get('id')
    user = message.get('from', {})
    user_name = f"{user.get('first_name', '')} {user.get('last_name', '')}".strip()
    username = user.get('username', '')
    inn_text = message.get('text', '').strip()
    message_id = message.get('message_id')
    
    logger.info(f"Получен запрос от {user_name} (@{username}): ИНН={inn_text}")
    
    # Валидация ИНН
    is_valid, validation_msg = validate_inn(inn_text)
    
    if not is_valid:
        error_text = f"""
⚠️ <b>Ошибка проверки ИНН</b>

{validation_msg}

<b>Требования к ИНН:</b>
• 10 цифр — для организаций
• 12 цифр — для физических лиц
• Только цифры, без пробелов и символов

Пожалуйста, проверьте и отправьте ИНН снова.
        """.strip()
        send_telegram_message(chat_id, error_text)
        return
    
    # Определение типа ИНН
    inn_type = "Физическое лицо" if len(inn_text) == 12 else "Организация"
    
    # Генерация отчёта
    request_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    html_content = generate_html_report(inn_text, user_name, request_date, inn_type)
    
    # Сохранение HTML-файла
    file_name = f"report_{inn_text}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    file_path = DATA_DIR / file_name
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    logger.info(f"HTML-отчёт сохранён: {file_path}")
    
    # Отправка документа
    caption = f"""
✅ <b>Ваш отчёт по ИНН готов!</b>

ИНН: <code>{inn_text}</code>
Тип: {inn_type}
Дата: {request_date}

Документ сгенерирован автоматически.
    """.strip()
    
    send_telegram_document(chat_id, str(file_path), caption)
    
    # Сохранение истории
    history_data = {
        'Дата': request_date,
        'Chat ID': chat_id,
        'User ID': user.get('id'),
        'Имя': user_name,
        'Username': f"@{username}" if username else "-",
        'ИНН': inn_text,
        'Тип ИНН': inn_type,
        'Статус': 'Успешно',
        'Файл': file_name
    }
    save_to_history(history_data)
    
    logger.info(f"Запрос от {user_name} обработан успешно")


# ==================== FLASK ПРИЛОЖЕНИЕ ====================

app = Flask(__name__)


@app.route('/webhook', methods=['POST'])
def webhook():
    """
    Обработчик webhook от Telegram
    """
    try:
        update = request.get_json()
        
        if not update:
            return jsonify({'status': 'error', 'message': 'Empty request'}), 400
        
        # Логирование входящего запроса
        logger.info(f"Получен webhook: {json.dumps(update, ensure_ascii=False)[:500]}")
        
        # Проверка, что это сообщение (не другие типы обновлений)
        if 'message' not in update:
            return jsonify({'status': 'ok', 'message': 'Not a message'}), 200
        
        message = update['message']
        
        # Игнорируем сообщения от ботов
        if message.get('from', {}).get('is_bot', False):
            return jsonify({'status': 'ok', 'message': 'Bot message ignored'}), 200
        
        # Проверка, что есть текст
        if 'text' not in message:
            return jsonify({'status': 'ok', 'message': 'No text'}), 200
        
        text = message['text'].strip()
        
        # Игнорируем команды
        if text.startswith('/'):
            if text == '/start':
                welcome_text = """
👋 <b>Добро пожаловать в Giga Hipot Bot!</b>

Я помогу вам получить отчёт по ИНН.

<b>Как использовать:</b>
1. Отправьте мне ИНН (10 или 12 цифр)
2. Я проверю корректность
3. Вы получите HTML-отчёт

<b>Пример:</b>
<code>123456789012</code>

Отправьте ИНН для начала работы!
                """.strip()
                send_telegram_message(message['chat']['id'], welcome_text)
            elif text == '/help':
                help_text = """
ℹ️ <b>Помощь</b>

Этот бот автоматически генерирует отчёты по ИНН.

<b>Команды:</b>
/start — Начать работу
/help — Эта справка
/status — Статус обработки

<b>Требования к ИНН:</b>
• 10 цифр — организации
• 12 цифр — физические лица
• Только цифры

Просто отправьте ИНН в чат!
                """.strip()
                send_telegram_message(message['chat']['id'], help_text)
            return jsonify({'status': 'ok', 'message': 'Command processed'}), 200
        
        # Обработка ИНН
        process_inn_request(update)
        
        return jsonify({'status': 'ok', 'message': 'Processed'}), 200
        
    except Exception as e:
        error_logger.error(f"Ошибка обработки webhook: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/health', methods=['GET'])
def health():
    """
    Проверка работоспособности сервера
    """
    return jsonify({
        'status': 'ok',
        'timestamp': datetime.now().isoformat(),
        'version': '1.0'
    }), 200


@app.route('/set-webhook', methods=['GET'])
def set_webhook():
    """
    Установка webhook в Telegram (для удобства)
    """
    webhook_url = f"{WEBHOOK_URL}/webhook"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/setWebhook"
    data = {"url": webhook_url}
    
    try:
        response = requests.post(url, json=data, timeout=10)
        result = response.json()
        
        if result.get("ok"):
            return jsonify({
                'status': 'ok',
                'message': f'Webhook установлен: {webhook_url}',
                'details': result
            }), 200
        else:
            return jsonify({
                'status': 'error',
                'message': 'Ошибка установки webhook',
                'details': result
            }), 400
            
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500


# ==================== ЗАПУСК ====================

if __name__ == '__main__':
    logger.info("🚀 Запуск Telegram Webhook Server v1.0")
    logger.info(f"📁 Директория данных: {DATA_DIR}")
    logger.info(f"📁 Директория шаблонов: {TEMPLATES_DIR}")
    logger.info(f"📁 Директория логов: {LOGS_DIR}")
    
    # Запуск сервера
    # Для production используйте gunicorn: gunicorn -w 4 -b 0.0.0.0:5000 app:app
    app.run(host='0.0.0.0', port=5000, debug=True)
