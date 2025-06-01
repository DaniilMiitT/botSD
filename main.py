import io
import ipaddress
import json
import logging.config
import os
import re
import time

import requests
from bot.bot import Bot
from bot.event import EventType
from bot.filter import Filter
from bot.handler import CommandHandler, MessageHandler, BotButtonCommandHandler
from pydantic import BaseModel

#импорт внутренних модулей 

import WMI.wmi_actions as winrm
from AD.ad_blocked_users import ad_unblock_user, get_ad_blocked_user
from AD.laps import get_laps_password
from AD.ldap_auth_vkteams import auth_vkteams_user
from AD.user_info import get_ad_user_info, format_user_info
from IM_API.im_api_caller import im_search_call_by_executor, im_set_call_field_value, \
    im_get_user_by_mail, im_add_note, im_set_call_state, im_remove_object
from SQL.onec_report import form_onec_report
from SQL.sccm_hardware_report import form_sccmhw_report
from SQL.sql_actions import redis_read, redis_insert, im_get_call_sql, im_get_call_notes
from SQL.tassovec_report import form_tassovec_report



logging.config.fileConfig("logging.ini")

#получение токена для бота из переменных окружения

TOKEN = os.getenv("TOKEN")
API_URL = "https://api.msg.tass.ru/bot/v1"

bot = Bot(token=TOKEN, api_url_base=API_URL, is_myteam=True)


class Search(BaseModel):
    search: str
    password: str
    chat_id: str


def validate_ip_address(address):
    try:
        ip = ipaddress.ip_address(address)
        return str(ip)
    except ValueError:
        return False

#отправление запроса к апи штормвола для проверки на блокировки IP
    
def check_sw_ip(ip):
    sw_uri = "https://swunlock.itass.local/api/check"
    response = requests.post(sw_uri, data={'ip': ip}, verify=False)
    return response

#удаление HTML-тегов и спецсимволов из текста

def striphtml(data):
    htmltagsremove = re.compile(r'<.*?>')
    htmlspecialremove = re.compile(r'&.*?;')
    result = htmltagsremove.sub(' ', data)
    result = htmlspecialremove.sub(' ', result)
    return result


def shield_mdv2_formatting_symbols(data):
    return re.sub(r'([~\"\'*_{}\[\]()+\-.!#])', r'\\\1', data)


def format_im_message(text):
    text = "\r\n>".join([line for line in text.splitlines() if line.strip()])
    result = shield_mdv2_formatting_symbols(text.split("С уважением,")[0].split('Обращение № IM-CL-')[0])
    if not result.endswith("\r\n") and not result.endswith("\n"):
        result += "\r\n"

    return result

#извлечение номеров заявок из текста

def parse_call_numbers(text):
    numbers = re.findall(r'\d{5,}(?:-\d{5,})?', text)

    
    parsed_numbers = []
    for number in numbers:
        if '-' in number:
            start, end = map(int, number.split('-'))
            parsed_numbers.extend(range(start, end + 1))
        else:
            parsed_numbers.append(int(number))

    return list(set(parsed_numbers))

#Формирование карточки заявки для отображения в боте

def im_form_call_msg(callinfo):
    if callinfo['Removed'] != 1:
        btn_text = "Восстановить" if callinfo['EntityStateName'] == "Выполнена" else "Заполнить решение" \
            if callinfo['EntityStateName'] == "В работе" and not callinfo['Solution'] else "Выполнить" \
            if callinfo['EntityStateName'] == "В работе" else "В работу"

        action = "writesolution" if btn_text == "Заполнить решение" else \
            "callAccomplished2Line" if callinfo['EntityStateName'] == "В работе" else "callOpened2Line"

        style = 'base' if btn_text == 'Восстановить' else 'primary'

        if btn_text == 'Выполнить' or btn_text == 'Заполнить решение':
            ikm = json.dumps([[{"text": btn_text, "callbackData": f"action_{action}_{callinfo['ID']}", "style": style},
                               {"text": "Изм. решение", "callbackData": f"action_writesolution_{callinfo['ID']}",
                                "style": 'attention'}],
                              [{"text": "На ожидание", "callbackData": f"action_callWaiting2Line_{callinfo['ID']}",
                                "style": 'primary'},
                               {"text": "Вопрос клиенту",
                                "callbackData": f"action_callWaitingInformation_{callinfo['ID']}",
                                "style": 'primary'}],
                              [{"text": "Возвр. в группу", "callbackData": f"action_callTo2Line_{callinfo['ID']}",
                                "style": 'base'},
                               {"text": "Возвр. диспетчерам", "callbackData": f"action_callTo1Line_{callinfo['ID']}",
                                "style": 'base'}],
                              [{"text": "Обновить инфо", "callbackData": f"action_refresh_{callinfo['ID']}",
                                "style": 'base'}],
                              [{"text": "Переписка", "callbackData": f"notes_{callinfo['ID']}",
                                "style": 'base'}]
                              ])
        elif btn_text == 'В работу' and callinfo['EntityStateName'] != 'Направлена в группу диспетчеров':
            ikm = json.dumps([[{"text": btn_text, "callbackData": f"action_{action}_{callinfo['ID']}", "style": style},
                               {"text": "Вернуть диспетчерам", "callbackData": f"action_callTo1Line_{callinfo['ID']}",
                                "style": 'attention'}],
                              [{"text": "Обновить инфо", "callbackData": f"action_refresh_{callinfo['ID']}",
                                "style": 'base'}],
                              [{"text": "Переписка", "callbackData": f"notes_{callinfo['ID']}",
                                "style": 'base'}]
                              ])
        elif callinfo['EntityStateName'] == 'Направлена в группу диспетчеров':
            ikm = json.dumps(
                [[{"text": "В рабочую группу", "callbackData": f"action_callTo2Line_{callinfo['ID']}", "style": style}],
                 [{"text": "Обновить инфо", "callbackData": f"action_refresh_{callinfo['ID']}", "style": 'base'}],
                 [{"text": "Переписка", "callbackData": f"notes_{callinfo['ID']}",
                   "style": 'base'}]
                 ])
        
        else:
            ikm = json.dumps([[{"text": btn_text, "callbackData": f"action_{action}_{callinfo['ID']}", "style": style}],
                              [{"text": "Обновить инфо", "callbackData": f"action_refresh_{callinfo['ID']}",
                                "style": 'base'}],
                              [{"text": "Переписка", "callbackData": f"notes_{callinfo['ID']}",
                                "style": 'base'}]
                              ])
    else:
        ikm = json.dumps(
            [[{"text": "Обновить инфо", "callbackData": f"action_refresh_{callinfo['ID']}", "style": 'base'}],
             [{"text": "Переписка", "callbackData": f"notes_{callinfo['ID']}",
               "style": 'base'}]
             ])

    calltype = "Инцидент" if "Инцидент" in callinfo['CallTypeFullName'] else "Заявка"
    title = (f"*\\[{callinfo['EntityStateName']}\\] - {calltype}: {callinfo['Number']}* "
             f"- {shield_mdv2_formatting_symbols(callinfo['CallSummaryName'])}\r\n")
    service = callinfo['ServiceAttendanceFullName'].replace("\\", "-") if callinfo['ServiceAttendanceFullName'] else \
        callinfo['ServiceItemFullName'].replace("\\", "-")
    client = f"*Клиент:* {callinfo['ClientFullName']}\r\n" \
             f"*Должность:* {callinfo['PositionName']}\r\n"
    # contacts = f"*Контакты:* {callinfo['Phone']}, {shield_mdv2_formatting_symbols(callinfo['Email'])}\r\n"
    contact_info = [callinfo['Phone'], shield_mdv2_formatting_symbols(callinfo['Email'])]
    contacts = ", ".join([i for i in contact_info if i])
    client_division_name = callinfo['ClientSubdivisionName'].replace(" \\", ",")
    client_division = f"*Подразделение:* {client_division_name}\r\n"
    description = f"*Описание:*\r\n>{format_im_message(callinfo['Description'])}" if callinfo['Description'] else ""
    group = f"*Группа:* {callinfo['QueueName'] or 'Нет'}\r\n"
    executor = f"*Исполнитель:* {callinfo['ExecutorFullName'] or 'Нет'}\r\n"
    solution = f"*Решение:*\r\n>{format_im_message(callinfo['Solution'])}" if callinfo[
        'Solution'] else '*Решение:* Не заполнено\r\n'
    if callinfo['Note']:
        lastmessage = (f"\r\n*Последнее сообщение по заявке:*\r\n"
                       f">*{callinfo['UtcDate'].strftime('%d.%m.%Y %H:%M')} - {callinfo['UserName']}:*\r\n"
                       f">{format_im_message(callinfo['Note'])}")
    else:
        lastmessage = ""

    text = f'{title}' \
           f'{"*Сервис:* " + shield_mdv2_formatting_symbols(service)}\r\n' \
           f"{client}" \
           f"*Контакты:* {contacts}\r\n" \
           f"{client_division}" \
           f"{description}" \
           f"{group}{executor}" \
           f"{solution}" \
           f"{lastmessage}" \
           f"{os.getenv('IM_API_URI')}?callNumber={callinfo['Number']}"

    if callinfo['Removed'] == 1:
        text = "*Эта заявка удалена!*\r\n" + text + "\r\n*Эта заявка удалена!*"
    return {'ikm': ikm, 'text': text}

# Формирование списка заявок для инженера
def im_form_engineer_call_list(userlastname, statename):
    text = f'Заявки в статусе {statename}:\r\n'
    for i in im_search_call_by_executor(userlastname):
        if i['EntityStateName'] == statename:
            text += f"{i['Number']} - {i['Summary']}\r\n"


# Главный обработчик сообщений от пользователей
def message_cb(bot, event):
    
    sqlread = redis_read(event.data['from']['userId'])
    if sqlread and sqlread['state'] == 'writesolution':
        # Пользователь заполняет решение по заявке
        print(im_set_call_field_value(sqlread['callid'], 'Call.Solution', event.text))

        msg = im_form_call_msg(im_get_call_sql(callid=sqlread['callid']))

        bot.edit_text(chat_id=event.from_chat, msg_id=sqlread['msgid'], text=msg['text'],
                      inline_keyboard_markup="{}".format(msg['ikm']), parse_mode="MarkdownV2")

        if event.chat_type == 'group':
            bot.delete_messages(chat_id=event.from_chat, msg_id=event.msgId)

        data = {
            'userid': event.data['from']['userId'],
            'state': 'idle',
        }
        redis_insert(data)
# примечание или сообщение
    elif sqlread and (sqlread['state'] == 'writenote' or sqlread['state'] == 'writemsg'):
        text = im_get_user_by_mail(event.data['from']['userId'])['FullName'] + ": <br>" + event.text
        note_type = 0 if sqlread['state'] == 'writenote' else 1
        print(im_add_note(sqlread['callid'], text, note_type))

        time.sleep(1)

        data = {
            'userid': event.data['from']['userId'],
            'state': 'idle',
        }
        redis_insert(data)

        print(im_set_call_state(sqlread['callstate'], sqlread['callid']))

        time.sleep(1)

        msg = im_form_call_msg(im_get_call_sql(callid=sqlread['callid']))

        bot.edit_text(chat_id=event.from_chat, msg_id=sqlread['msgid'], text=msg['text'],
                      inline_keyboard_markup="{}".format(msg['ikm']), parse_mode="MarkdownV2")

        if event.chat_type == 'group':
            bot.delete_messages(chat_id=event.from_chat, msg_id=event.msgId)

 # отправление списка ПК для добавления в коллекцию KES
            
    elif event.text != "/kes" and sqlread and sqlread['state'] == 'kes_wait_for_pc_list':
        data = {
            'userid': event.data['from']['userId'],
            'state': 'idle',
        }
        redis_insert(data)

        pc_list = event.text.split('\n')
        

        bot.send_text(event.from_chat, "Добавляем ПК в коллекцию установки KNA и KES, ожидайте...")
        collectionid = 'CM100416'  

        result = winrm.add_pc_to_sccm_collection_winrm(collectionid=collectionid, computers=pc_list)
        print(result)

        with io.StringIO() as file:
            file.write(str(result))
            file.name = "cm_col_op.log"
            file.seek(0)
            text = "Произошла ошибка во время добавления ПК!\r\n\r\n" if result[0] == 1 else ""

            text += (f"Проверьте коллекцию {collectionid}, должны быть добавлены ПК:\r\n"
                     f"{', '.join(pc_list)}\r\n\r\nВывод команды для диагностики во вложении")

            bot.send_file(event.from_chat, caption=text, file=file)
            file.close()

     # Проверка IP на правильность
    ipcheck = validate_ip_address(event.text)
    if ipcheck:
        group = "SDBot_SWIPCheck"
        bot.send_text(chat_id=event.from_chat, text="Авторизация пользователя...")
        auth = auth_vkteams_user(event.data['from']['userId'], group)
    else:
        auth = False

    if ipcheck and auth:
        bot.send_text(chat_id=event.from_chat, text="Проверяем IP-адрес через API Stormwall...")

        response = check_sw_ip(event.text)
        if response.text == "False":
            bot.send_text(chat_id=event.from_chat, text="IP " + event.text + " *не заблокирован* в SW",
                          parse_mode="MarkdownV2")
        elif response.text == "True":
            bot.send_text(chat_id=event.from_chat,
                          text="IP " + event.text + " *заблокирован* в SW! Напишите в поддержку ДЦТ!",
                          parse_mode="MarkdownV2")
        else:
            text = "Error in stormwall API Call!\r\n" + "Error code: *" + str(response.status_code) + \
                   "*\r\nError message:\r\n*" + str(response.reason) + "*\r\nAPI URL: " + str(response.url)
            bot.send_text(chat_id=event.from_chat, text=text, parse_mode="MarkdownV2")

    elif ipcheck and not auth:
        text = f"Вы не состоите в группе AD <b>{group}</b>, напишите на sd@tass.ru!"
        bot.send_text(chat_id=event.from_chat, text=text, parse_mode="HTML")


    # Запрос на проверку новых версий приложений
    elif event.text == "/av" or event.text == "/appversions":
        bot.send_text(chat_id=event.from_chat, text="Проверяем новые версии приложений, займёт до 30 секунд...",
                      parse_mode="MarkdownV2")
        try:
            requests.get(f"http://{os.getenv('WORKER_API_URL')}:5000/av/{event.from_chat}", timeout=0.0000000001)
        except requests.exceptions.ReadTimeout:
            pass
        bot.send_actions(chat_id=event.from_chat, actions=['typing'])
    else:
        pass

#обработка команды удаления заявки 
def im_del_req_cb(bot, event):
    print('im_del_req_cb')
    group = "SDBot_IM_DeleteCall"
    bot.send_text(chat_id=event.from_chat, text=f"Авторизация пользователя по группе {group}...")
    auth = auth_vkteams_user(event.data['from']['userId'], group)
    if auth:
        callnumbers = parse_call_numbers(event.data['parts'][0]['payload']['message']['text']
                                         if 'parts' in event.data.keys() else event.text)
        if callnumbers:
            # удаляет дублирующие номера
            callnumbers = list(set(callnumbers))
            for number in callnumbers:
                callinfo = im_get_call_sql(callnumber=number)
                if callinfo == f"{number}: Заявка с таким номером не найдена!":
                    bot.send_text(chat_id=event.from_chat, text=callinfo)
                 # Если заявка найдена и еще не удалена
                elif callinfo['Removed'] != 1:
                    call_id = str(callinfo['ID'])
                    print(call_id)
                    msgid = event.data['msgId']
                     # Создание кнопки "удалить?"
                    ikm = json.dumps(
                        [[{"text": "Удалить?", "callbackData": "delete_" + call_id + "_" + msgid,
                           "style": "attention"}]])
                    description = callinfo.get('Summary') or callinfo.get('CallSummaryName') or callinfo.get(
                        'Description', 'Описание не найдено')
                    text = f"Удаление заявки: {callinfo['Number']} {striphtml(description)}"
                    bot.send_text(chat_id=event.from_chat, text=text, inline_keyboard_markup="{}".format(ikm))
                else:
                    bot.send_text(chat_id=event.from_chat,
                                  text=f"Заявка уже удалена!\r\n{callinfo['Number']}: {callinfo['CallSummaryName']}")
        else:
            text = 'Не удалось найти номер заявки!'
            bot.send_text(chat_id=event.from_chat, text=text)
    else:
        text = f'Вы не состоите в группе AD {group}'
        bot.send_text(chat_id=event.from_chat, text=text)

  # Обработка нажатия на кнопку "Удалить"
def im_delete_call_cb(bot, event):
    print(event.data['callbackData'])
    callback_data = event.data['callbackData'].split("_")
    call_id = callback_data[1]
    msgid = callback_data[2]
    
    im_remove_object(call_id)
    bot.answer_callback_query(query_id=event.data['queryId'], text="Удалено из ИМ", show_alert=False)
    if event.chat_type == 'group':
        bot.delete_messages(chat_id=event.from_chat, msg_id=msgid)
    bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                  text=event.data['message']['text'] + "\r\nУдалено!")

# Получаем информацию о заявке по номеру
def im_get_callinfo_cb(bot, event):
    parsed_text = event.data['parts'][0]['payload']['message']['text'] if 'parts' in event.data.keys() else event.text
    print(parsed_text)
    if parsed_text.startswith("https://files-n.msg.tass.ru") and "\n" in parsed_text:
        parsed_text = parsed_text.split('\n')[1]
    elif parsed_text.startswith("https://files-n.msg.tass.ru"):
        return
    # Извлекаем номер заявки
    callnumber = parse_call_numbers(parsed_text)
    print(callnumber)
    for n in callnumber:
        if n and (len(str(n)) == 6 or "/call" in parsed_text):
            callinfo = im_get_call_sql(callnumber=n)
            if callinfo != f"{n}: Заявка с таким номером не найдена!":
                msg = im_form_call_msg(callinfo)

                bot.send_text(chat_id=event.from_chat, text=msg['text'], inline_keyboard_markup="{}".format(msg['ikm']),
                              parse_mode="MarkdownV2")
            else:
                bot.send_text(chat_id=event.from_chat, text=callinfo)

# обработка кнопок связанных с действиями по заявкам 
def im_call_action_cb(bot, event):
    stateid = event.data['callbackData'].split("_")[1]
    if stateid == 'writesolution':
        bot.answer_callback_query(query_id=event.data['queryId'],
                                  text="Напишите и отправьте текст решения в сообщении боту, "
                                       "после чего появится кнопка Выполнить", show_alert=True)
        # сохраняем состояние во временное хранилище
        data = {
            "userid": event.data['from']['userId'],
            "callid": event.data['callbackData'].split("_")[2],
            "msgid": event.data['message']['msgId'],
            "state": "writesolution",
        }
        redis_insert(data)

        callid = event.data['callbackData'].split("_")[2]
        ikm = json.dumps([[{"text": "Отменить", "callbackData": f"action_cancel_{callid}", "style": 'attention'}]])
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=event.data['message']['text'],
                      inline_keyboard_markup="{}".format(ikm), parse_mode='MarkdownV2')
    # перевод заявки диспетчеру
    elif stateid == 'callTo1Line':
        bot.answer_callback_query(query_id=event.data['queryId'],
                                  text="Напишите и отправьте текст для диспетчера в сообщении боту, "
                                       "после чего заявка отправится в группу диспетчеров", show_alert=True)
        data = {
            "userid": event.data['from']['userId'],
            "callid": event.data['callbackData'].split("_")[2],
            "msgid": event.data['message']['msgId'],
            "state": "writenote",
            "callstate": stateid,
        }
        redis_insert(data)

        callid = event.data['callbackData'].split("_")[2]
        ikm = json.dumps([[{"text": "Отменить", "callbackData": f"action_cancel_{callid}", "style": 'attention'}]])
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=event.data['message']['text'],
                      inline_keyboard_markup="{}".format(ikm), parse_mode='MarkdownV2')
    # перевод заявки в ожидание на 2 линии
    elif stateid == 'callWaiting2Line':
        bot.answer_callback_query(query_id=event.data['queryId'],
                                  text="Напишите и отправьте причину ожидания для пользователя в сообщении боту, "
                                       "после чего заявка перейдет в Ожидание", show_alert=True)
        data = {
            "userid": event.data['from']['userId'],
            "callid": event.data['callbackData'].split("_")[2],
            "msgid": event.data['message']['msgId'],
            "state": "writemsg",
            "callstate": stateid,
        }
        redis_insert(data)

        callid = event.data['callbackData'].split("_")[2]
        ikm = json.dumps([[{"text": "Отменить", "callbackData": f"action_cancel_{callid}", "style": 'attention'}]])
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=event.data['message']['text'],
                      inline_keyboard_markup="{}".format(ikm), parse_mode='MarkdownV2')
    # перевод в статус "ожидание информации"
    elif stateid == 'callWaitingInformation':
        bot.answer_callback_query(query_id=event.data['queryId'],
                                  text="Напишите и отправьте текст вопроса пользователю в сообщении боту, "
                                       "после чего заявка перейдет в Ожидание информации", show_alert=True)
        data = {
            "userid": event.data['from']['userId'],
            "callid": event.data['callbackData'].split("_")[2],
            "msgid": event.data['message']['msgId'],
            "state": "writemsg",
            "callstate": stateid,
        }
        redis_insert(data)

        callid = event.data['callbackData'].split("_")[2]
        ikm = json.dumps([[{"text": "Отменить", "callbackData": f"action_cancel_{callid}", "style": 'attention'}]])
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=event.data['message']['text'],
                      inline_keyboard_markup="{}".format(ikm), parse_mode='MarkdownV2')
    # обновление информации о заявке
    elif stateid == 'refresh':
        callid = event.data['callbackData'].split("_")[2]
        msg = im_form_call_msg(im_get_call_sql(callid=callid))
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=msg['text'], inline_keyboard_markup="{}".format(msg['ikm']), parse_mode='MarkdownV2')
        bot.answer_callback_query(query_id=event.data['queryId'], text="Информация о заявке обновлена!")

        # отмена предыдущего действия 
    elif stateid == 'cancel':
        data = {
            'userid': event.data['from']['userId'],
            'state': 'idle',
        }
        redis_insert(data)

        callid = event.data['callbackData'].split("_")[2]
        msg = im_form_call_msg(im_get_call_sql(callid=callid))
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=msg['text'], inline_keyboard_markup="{}".format(msg['ikm']), parse_mode='MarkdownV2')
       # Обработка всех остальных состояний
    else:
        callid = event.data['callbackData'].split("_")[2]
        callinfo = im_get_call_sql(callid=callid)

        if callinfo['EntityStateName'] == "Выполнена":
            print(im_add_note(callid, "Заявка восстнаовлена исполнителем", 1))

        if stateid == 'callTo2Line':
            print(im_add_note(callid, "Заявка возвращена в группу исполнителем", 0))

        set_executor = ''
        set_accomplisher = ''
        if stateid == 'callOpened2Line':
            set_executor = im_set_call_field_value(callid, 'Call.Executor',
                                                   im_get_user_by_mail(event.data['from']['userId']))
        if stateid == 'callAccomplished2Line':
            set_accomplisher = im_set_call_field_value(callid, 'Call.Accomplisher',
                                                       im_get_user_by_mail(event.data['from']['userId']))

        set_state = im_set_call_state(stateid, callid)

        time.sleep(1)

        if set_state['Result'] == 0 and set_executor == 'Set Call.Executor success':
            bot.answer_callback_query(query_id=event.data['queryId'], text="Заявка взята в работу!", show_alert=False)
        elif set_state['Result'] == 0 and set_accomplisher == 'Set Call.Accomplisher success':
            bot.answer_callback_query(query_id=event.data['queryId'], text="Заявка выполнена!", show_alert=False)
        else:
            bot.answer_callback_query(query_id=event.data['queryId'], text=set_state['Message'], show_alert=True)

        time.sleep(1)

        msg = im_form_call_msg(im_get_call_sql(callid=callid))

        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'],
                      text=msg['text'], inline_keyboard_markup="{}".format(msg['ikm']), parse_mode='MarkdownV2')
        
#сообщение со списком заблокированных пользователей 
def locked_users_msg():
    buttons = []
    users = get_ad_blocked_user()
    print(users)
    if users:
        for user in users:
            buttons.append([{"text": user, "callbackData": "ad_unblock_" + user, "style": "primary"}])
        ikm = json.dumps(buttons)
        return {'text': 'Заблокированные пользователи:', 'ikm': ikm}
    else:
        return {'text': 'Заблокированные пользователи не найдены!'}
    
# Команда /unblock — показывает список заблокированных пользователей
def ad_unblock_user_req(bot, event):
    bot.send_actions(chat_id=event.from_chat, actions=['typing'])
    msg = locked_users_msg()
    if msg['text'] == 'Заблокированные пользователи:':
        ikm = msg['ikm']
        bot.send_text(chat_id=event.from_chat, text=msg['text'], inline_keyboard_markup=ikm)
    else:
        bot.send_text(chat_id=event.from_chat, text=msg['text'])
    bot.send_actions(chat_id=event.from_chat, actions=[''])

# Обработка нажатия на кнопку "Разблокировать"
def ad_unblock_user_cb(bot, event):
    name = event.data['callbackData'].split("_")[2]
    if ad_unblock_user(name):
        bot.answer_callback_query(query_id=event.data['queryId'], text="Пользователь разблокирован!")
    else:
        bot.answer_callback_query(query_id=event.data['queryId'], text="Ошибка при разблокировке пользователя!")
        
    # Обновляем список пользователей на экране
    msg = locked_users_msg()
    if msg['text'] == 'Заблокированные пользователи:':
        ikm = msg['ikm']
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'], text=msg['text'],
                      inline_keyboard_markup=ikm)
    else:
        bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'], text=msg['text'])

# Команда /laps <имя_компьютера> — выдает пароль LAPS
def laps_cb(bot, event):
    group = "SDBot_SWIPCheck"
    bot.send_text(chat_id=event.from_chat, text="Авторизация пользователя...")
    auth = auth_vkteams_user(event.data['from']['userId'], group)
    rec_text = event.data.get('parts') and event.data['parts'][0]['payload']['message']['text'] or event.text or ""
     # Форматируем и отправляем пароль
    if auth:
        text = shield_mdv2_formatting_symbols(get_laps_password(rec_text.split(" ")[1]))
        bot.send_text(chat_id=event.from_chat, text=text, parse_mode="MarkdownV2")


@bot.command_handler(command='kes')
def kes_cb(bot, event):
    bot.send_text(event.from_chat, "Введите список ПК, разделенный переносом строки")

    data = {
        "userid": event.data['from']['userId'],
        "state": "kes_wait_for_pc_list", # Ждем от сотрудника список
    }
    redis_insert(data)

#команда для отчета по тассовцу
@bot.command_handler(command='tassovecreport')
def tassovecreport_cb(bot, event):
    form_tassovec_report()
    with open('tassovec.xlsx', 'rb') as f:
        f.read()
        f.seek(0)
        text = "Ваш отчет готов!"
        print(f)
        bot.send_file(event.from_chat, caption=text, file=f)
        f.close()
        os.remove("tassovec.xlsx")

# Команда /report1c — ежедневный отчет 1С
@bot.command_handler(command='report1c')
def onecreport_cb(bot, event):
    form_onec_report()
    with open('1c_daily.xlsx', 'rb') as f:
        f.read()
        f.seek(0)
        text = "Ваш отчет готов!"
        print(f)
        bot.send_file(event.from_chat, caption=text, file=f)
        f.close()
        os.remove("1c_daily.xlsx")

# Команда /hwreport — отчет SCCM
@bot.command_handler(command='hwreport')
def hwreport_cb(bot, event):
    form_sccmhw_report()
    with open('sccm_hw_report.xlsx', 'rb') as f:
        f.read()
        f.seek(0)
        text = "Ваш отчет готов!"
        print(f)
        bot.send_file(event.from_chat, caption=text, file=f)
        f.close()
        os.remove("sccm_hw_report.xlsx")

# Команда /callnote или кнопка — показать переписку по заявке
@bot.command_handler(command='callnote')
@bot.button_handler(filters=Filter.callback_data_regexp("notes*"))
def get_call_notes(bot, event):
    if event.type == EventType.NEW_MESSAGE and "/callnote" in event.text:
        callid = im_get_call_sql(callnumber=event.text.split("/callnote")[1].strip())['ID']
    else:
        callid = event.data['callbackData'].split("_")[1]
    notes = im_get_call_notes(callid)
    if notes:
        # Формируем текст переписки
        text = f"*Переписка по заявке № {notes[0]['Number']}:*\r\n"
        i = 0
        for n in notes:
            i += 1
            text += f"*Сообщение {i}:*\r\n" \
                    f">*{n['UtcDate'].strftime('%d.%m.%Y %H:%M')} - {n['UserName']}:*\r\n" \
                    f">{format_im_message(n['Note'])}" \
                    f"============\r\n"
    else:
        text = "Переписка по заявке отсутствует!"

    print(text)
    # Отправляем или показываем текст
    if event.type == EventType.CALLBACK_QUERY and text == "Переписка по заявке отсутствует!":
        bot.answer_callback_query(query_id=event.data['queryId'], text=text, show_alert=True)
    elif event.type == EventType.CALLBACK_QUERY:
        bot.answer_callback_query(query_id=event.data['queryId'], text='')
        bot.send_text(chat_id=event.from_chat, text=text, parse_mode='MarkdownV2')
    else:
        bot.send_text(chat_id=event.from_chat, text=text, parse_mode='MarkdownV2')


@bot.command_handler(command='img')
def img_test(bot, event):
    file = "0x89504E470D0A1A0A0000000D4948445200000010000000100802000000909168360000000467414D410000B18F0BFC6105000000097048597300000EC400000EC401952B0E1B0000002D74455874536F66747761726500437265617465642062792066436F6465722047726170686963732050726F636573736F727FC3ED5F0000020F49444154384F63F88F15FCFBFFEF1F94890650347C7FF7F3D282671B32EFCFF27F3C39E0F5ECCCCFFB16FCFEF21E2A0B01080DAF0F3EDB177460AAEDC650953633C90213C9620F95C945B657DB837E5E3A04550304500D1F0F3D38EDB53A45B59297CD858BD5098E78D85CDC54A737787DBB7CF02F442548C3BFF75F1F05CFCD55CD47568A8CDC55A775057DF9F20EE42D90866F0B0F1CB36DE06573064973BA70280473288470CA84C03500ED29B7BD7070C177A886AF45932AD5132072CC33BA189E1D60787481E1FE7DD6A46EE649FB38450380E23E2AFD8B335F8335FCFBF725A93C503610A281695B37C3AB0D0CCFF6333C3ECF70E70EC3ED771CB261407173C982B9FEF780C10DD2F03D272F48D10FAA61671DC3BB790CAFD6313CDFC7F0E00CC3FDBB1CB220B75949E6AE08B806D6F0FFFF8FF6EA1A13A80DCCF3B3193EF531BC9BCBF06A2D6B7A39F3CC859CA2BE40F118D5FABD9997A07EF8B363C5E5B8505E3690062E4E670E156F0E155F4E79903A080286C712DBC57716DE836AF8FFE9EDAFF200570537B80A3494A35A783E78E3CFF73F601AFEFCFAD3E8E1A4881265100434BB4C35FD9ED72C60CC82544235FCF8FCB7CBE6508AD3E50CCFF389013516BE41CADEC060A8D68D3CEB52FC3EA4F7DBC19B60C52000D6F0F7E7FF6DE5FF7676FF3DB9FECFEE45BF7A8B7E96247DCFCBFE5ED3FE73D5CE7FEF3E8355420158032600266EACE9FBFF7F00E66CC0D658C6865E0000000049454E44AE426082"
    bot.send_file(caption='123', file=file, chat_id=event.from_chat)

# /kp <поиск> поиск пароля в Keepass
@bot.command_handler(command='kp')
def keepass_handler(bot, event):
    bot.send_actions(chat_id=event.from_chat, actions=['typing'])
    group = "SDBot_Keepass_Reader"
    auth = auth_vkteams_user(event.data['from']['userId'], group)
    if not auth or event.data['chat']['type'] != 'private':
        return
    search = Search(search=event.data['text'].split('/kp ')[1].replace("_", r"\_").strip(),
                    password=os.getenv('KEEPASS_PASSWORD'),
                    chat_id=event.from_chat)
    search = search.model_dump_json()
    headers = {"Content-Type": "application/json; charset=UTF-8"}
    try:
        requests.post(f"http://{os.getenv('WORKER_API_URL')}:5000/kp", data=search.encode(), headers=headers,
                      timeout=1)
    except requests.exceptions.ReadTimeout:
        pass

# /bios <поиск> запрос биос пароля
@bot.command_handler(command='bios')
def bios_handler(bot, event):
    print(event.data['chat']['type'])
    bot.send_actions(chat_id=event.from_chat, actions=['typing'])
    group = "SDBot_Keepass_Reader"
    auth = auth_vkteams_user(event.data['from']['userId'], group)
    if not auth or event.data['chat']['type'] != 'private':
        return
    search = Search(search=event.data['text'].split('/bios ')[1].replace("_", r"\_").strip(),
                    password='notneeded', chat_id=event.from_chat)
    search = search.model_dump_json()
    headers = {"Content-Type": "application/json; charset=UTF-8"}
    try:
        requests.post(f"http://{os.getenv('WORKER_API_URL')}:5000/bios", data=search.encode(), headers=headers,
                      timeout=1)
    except requests.exceptions.ReadTimeout:
        pass

# /user +логин получить информацию о пользователе из AD
@bot.command_handler(command='user')
def user_info_handler(bot, event) -> None:
    users = get_ad_user_info(event.text.split('/user ')[1])
    i = 0
    for u in users:
        user_attributes = json.loads(u.entry_to_json())['attributes']
        text = "".join(format_user_info(user_attributes))
        msg = format_user_info(user_attributes)[0]
        ikm = json.dumps(
            [[{"text": "Группы пользователя", "callbackData": "groups___" + text,
               "style": "base"}]])
        if i < 5:
            bot.send_text(text=msg, parse_mode=None, chat_id=event.from_chat, inline_keyboard_markup="{}".format(ikm))
        i += 1

#кнопка для показа групп пользователя (тк их много)
@bot.button_handler(filters=Filter.callback_data_regexp("groups_*"))
def groups_button_handler(bot, event):
    text = event.data['callbackData'].split("___")[1]
    bot.edit_text(chat_id=event.from_chat, msg_id=event.data['message']['msgId'], text=text, parse_mode=None)


@bot.command_handler(command='kubetoken')
def kubetoken_handler(bot, event):
    bot.send_actions(chat_id=event.from_chat, actions=['typing'])
    chat_id = Search(search='notneeded', password='notneeded', chat_id=event.from_chat).model_dump_json().encode()
    headers = {"Content-Type": "application/json; charset=UTF-8"}
 # Отправляет POST-запрос на API воркера, чтобы получить токен
    try:
        requests.post(f"http://{os.getenv('WORKER_API_URL')}:5000/kubetoken", data=chat_id, headers=headers,
                      timeout=1)
    except requests.exceptions.ReadTimeout:
        pass

# Регистрируем все обработчики (команды, кнопки, сообщения и т.д.)
def launch_handlers(bot):
    bot.dispatcher.add_handler(
        BotButtonCommandHandler(callback=im_delete_call_cb, filters=Filter.callback_data_regexp("delete_*")))
    bot.dispatcher.add_handler(
        BotButtonCommandHandler(callback=im_call_action_cb, filters=Filter.callback_data_regexp("action_*")))
    bot.dispatcher.add_handler(
        BotButtonCommandHandler(callback=ad_unblock_user_cb, filters=Filter.callback_data_regexp("ad_unblock_*")))

    bot.dispatcher.add_handler(MessageHandler(callback=message_cb))

    bot.dispatcher.add_handler(MessageHandler(filters=Filter.forward, callback=im_get_callinfo_cb))
    bot.dispatcher.add_handler(MessageHandler(filters=Filter.reply, callback=im_get_callinfo_cb))

    bot.dispatcher.add_handler(CommandHandler(command="call", callback=im_get_callinfo_cb))
    bot.dispatcher.add_handler(CommandHandler(command="unblock", callback=ad_unblock_user_req))
    bot.dispatcher.add_handler(CommandHandler(command="delcall", callback=im_del_req_cb))
    bot.dispatcher.add_handler(CommandHandler(command="laps", callback=laps_cb))


def main():
    launch_handlers(bot)
    bot.start_polling()
    bot.idle()


if __name__ == '__main__':
    main()
