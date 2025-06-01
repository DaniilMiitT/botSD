import os
import sys

import requests
from fastapi import APIRouter
from fastapi.responses import Response

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.dirname(SCRIPT_DIR))
# Импорт функции для проверки новых версий приложений
from check_new_app_versions import check_new_app_versions

av = APIRouter()

# Обработчик GET-запроса: /av/{chatId}
@av.get('/av/{chatId}')
async def get_new_app_versions(chatId: str):
    # Получаем текст с информацией о новых версиях приложений
    text = check_new_app_versions()
    
    send_text_url = 'https://api.msg.tass.ru/bot/v1/messages/sendText'

    body = {
        'token': os.getenv('TOKEN'),
        'chatId': chatId,
        'text': text,
        'parseMode': 'MarkdownV2'
    }

    requests.get(url=send_text_url, params=body, headers={'Content-Type': 'application/json;charset=utf-8'})

    with open(file='AppVersionLog.json', mode='r', encoding='utf-8') as f:
        response = f.read()
        f.close()
# Возвращаем содержимое файла как JSON-ответ
    return Response(content=response, media_type='application/json')
