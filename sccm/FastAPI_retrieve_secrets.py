import io
import os
import subprocess

import requests
from fastapi import APIRouter
from pydantic import BaseModel
from pykeepass import PyKeePass

kp = APIRouter()
bios = APIRouter()
kubetoken = APIRouter()

# Модель данных для POST-запросов
class Search(BaseModel):
    search: str
    password: str
    chat_id: str


@kp.post('/kp/')
def find_keepass_entry(search: Search):
    # Учетные данные Nextcloud (где хранится база Keepass)
    ncuser = "keepass"
    ncpassword = os.getenv('KEEPASS_PASSWORD')
    url = "https://disk.tass.ru/remote.php/webdav/USP_passwords.kdbx"
    # Проверка авторизации
    if search.password != ncpassword:
        return "Запрос не авторизован паролем от keepass!"
# Загрузка базы данных из Nextcloud
    print("Downloading Keepass db")
    db_dl = requests.get(url, auth=(ncuser, ncpassword)).content

    print("Loading Keepass db into memory")
    db = io.BytesIO(db_dl)
# Открытие базы с помощью PyKeePass
    print("Open the Keepass database in memory using PyKeePass")
    kpdb = PyKeePass(db, password=os.getenv('KEEPASS_PASSWORD'))
# Поиск записей по названию 
    find = kpdb.find_entries(title=f'(?i){search.search}', regex=True)
# Формирование ответа
    result = '\n'.join([f"{i.title}: {i.password}" for i in find])

    db.close()
# Отправка результата в чат
    send_message_from_fastapi(result, search.chat_id)

    return result or f"Записи по поиску {search.search} не найдены!"


@bios.post('/bios/')
def get_bios_password(search: Search):
    token = os.getenv('BIOS_TOKEN')
    headers = {
        "X-App-Token": token,
        "Content-Type": "application/json",
    }
    # Подготовка тела запроса для получения пароля
    data = {"_action": "GET", "name": search.search}
    request = requests.post('https://msk-spoon-app02.corp.tass.ru/appbios/api', json=data, headers=headers,
                            verify=False)
    # Обработка ответа
    data = request.json()['OK']
    result = "\n".join([f"{key}: {value}" for key, value in data.items()])
    # Отправка результата в чат
    send_message_from_fastapi(result, search.chat_id)

    return result or f"Записи по поиску {search.search} не найдены!"


def send_message_from_fastapi(text, chat_id):
    send_text_url = 'https://api.msg.tass.ru/bot/v1/messages/sendText'
    send_actions_url = 'https://api.msg.tass.ru/bot/v1/chats/sendActions'

    body = {
        'token': os.getenv('TOKEN'),
        'chatId': chat_id,
        'text': text or "Записи по поиску не найдены!",
    }

    actions = {
        'token': os.getenv('TOKEN'),
        'chatId': chat_id,
        'actions': ''
    }
  # Имитирует "пишет сообщение..." в чате
    requests.get(url=send_actions_url, params=actions, headers={'Content-Type': 'application/json;charset=utf-8'})
    # Отправляет само сообщение
    requests.get(url=send_text_url, params=body, headers={'Content-Type': 'application/json;charset=utf-8'})


@kubetoken.post('/kubetoken/')
def get_kubernetes_dashboard_token(chat_id: Search):
    ssh_user = os.getenv('SSH_USER')
    ssh_private_key = "/app/.ssh/id_rsa"# Путь к приватному ключу SSH
    chat_id = chat_id.chat_id
# Команда для создания токена через SSH
    command = [
        "ssh", "-i", ssh_private_key, "-o", "StrictHostKeyChecking=no", f"{ssh_user}@m2t-k3s01.corp-sd.tass.ru",
        "kubectl create token -n kubernetes-dashboard admin-user"
    ]

    print(command)

    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        token = result.stdout.strip()
        send_message_from_fastapi(token, chat_id)
        return token
    except subprocess.CalledProcessError as e:
        print("Error obtaining token:", e.stderr)
        return None
