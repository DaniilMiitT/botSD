from datetime import datetime

from ldap3 import ALL_ATTRIBUTES

from AD.ad_blocked_users import OU, ldap_auth

# функция поиска пользователей в ад
def get_ad_user_info(search):
    conn = ldap_auth()
    print(conn)

    
    search_query = f"""
                    (&
                      (objectClass=user)
                      (|
                        (name=*{search}*)
                        (sAMAccountName=*{search}*)
                      )
                    )
                    """
# Поиск по AD с получением всех атрибутов
    conn.search(search_base=OU, search_filter=search_query, search_scope='SUBTREE',
                attributes=ALL_ATTRIBUTES)

    results = conn.entries # Сохраняем найденные записи

    conn.unbind()

    return results

# Функция форматирования информации о пользователе для удобного вывода
def format_user_info(user):
    # Получение основных данных пользователя
    name = ', '.join(user.get('name', ['Нет']))
    login = ', '.join(user.get('sAMAccountName', ['Нет']))
    email = ', '.join(user.get('mail', ['Нет']))
    phone = ', '.join(user.get('telephoneNumber', ['Нет']))
    mobile = ', '.join(user.get('mobile', ['Нет']))
    title = ', '.join(user.get('title', ['Нет']))
    mail_addresses = ', '.join(user.get('proxyAddresses', ['Нет']))
    mail_addresses = mail_addresses.lower().replace("smtp:", "")
    department_full = ', '.join(user.get('TASS-1C-DepartmentPath', ['Нет']))

    # Проверка статуса учетной записи
    uac = user.get('userAccountControl')
    
    status = 'Включена' if uac[0] == 512 else 'Отключена' if uac[0] == 514 else f'Непон: {uac[0]}'
    # Проверка блокировки
    lockout_timestamp = ', '.join(user.get('lockoutTime', ['Нет']))
    bad_pwd_count = user.get('badPwdCount', [0])[0]
    lockout_time = datetime.fromisoformat(lockout_timestamp).strftime("%d.%m.%Y в %H:%M:%S") \
        if "Нет" not in lockout_timestamp else "Нет"
    bad_pwd_timestamp = ', '.join(user.get('badPasswordTime', ['Нет']))
    bad_pwd_time = datetime.fromisoformat(bad_pwd_timestamp).strftime("%d.%m.%Y в %H:%M:%S") \
        if "Нет" not in bad_pwd_timestamp else "Нет"
    if "Нет" not in lockout_time and "01.01.1601" not in lockout_time:
        lock = (
            f"Да\n"
            f"Время блока: {lockout_time}\n"
            f"Неудачных попыток ввода пароля: {bad_pwd_count}\n"
            f"Время неудачного ввода пароля: {bad_pwd_time}"
        )
    else:
        lock = "Нет"
        # Дополнительные атрибуты
    ad_path = ', '.join(user.get('distinguishedName'))
    dismissal_date = ', '.join(user.get('TASS-DismissalDate', ['Нет']))
    dismissed = f"Да\nДата увольнения: {dismissal_date}" if "УВОЛЕНЫ" in ad_path else "Нет"
    
    pwd_last_set_timestamp = ', '.join(user.get('pwdLastSet', ['Нет']))
    pwd_last_set = datetime.fromisoformat(pwd_last_set_timestamp).strftime("%d.%m.%Y в %H:%M:%S") \
        if "Нет" not in lockout_timestamp else "Нет"
    fmc = ', '.join(user.get('TASS-Fmc', ['Нет']))
    emp_num = ', '.join(user.get('employeeNumber', ['Нет']))
    
    maid_db = ', '.join(user.get('homeMDB', ['Нет']))
    maid_db = maid_db.split("=")[1].split(",")[0] if "=" in maid_db else maid_db
    
    last_logon_timestamp = ', '.join(user.get('lastLogonTimestamp', ['Нет']))
    last_logon = datetime.fromisoformat(last_logon_timestamp).strftime("%d.%m.%Y в %H:%M:%S") \
        if "Нет" not in last_logon_timestamp else "Нет"

    # Обработка списков общих почтовых ящиков
    shared_mb_list = user.get('msExchDelegateListBL', ['Нет'])
    shared_mb = []
    for i in shared_mb_list:
        shared_mb_address = (i.split("=")[1].split(",")[0]) if "=" in i else i
        shared_mb.append("- " + shared_mb_address)
    shared_mb = "\n".join(sorted(shared_mb))
    user_sid = ', '.join(user.get('objectSid'))
    tg_id = user.get('TASS-TelegramUser', ["Нет"])
    tg_id = ", ".join(map(str, tg_id))
    expire_timestamp = ", ".join(user.get('accountExpires', ["Нет"]))
    expire_time = datetime.fromisoformat(expire_timestamp).strftime("%d.%m.%Y в %H:%M:%S") \
        if "Нет" not in expire_timestamp else expire_timestamp
    expire_time = "Никогда" if "01.01.1601" in expire_time else expire_time
    group_list = user.get('memberOf', ['Нет'])
    groups_unsorted = []
    for g in group_list:
        # Разделяем DN на части
        parts = g.split(",")
        # Собираем части пути до 'OU=Groups'
        ou_parts = []
        cn = None
        for part in parts:
            if part.startswith("CN="):
                cn = part.split("=")[1]
            elif part.startswith("OU="):
                ou = part.split("=")[1]
                if ou == "Groups":
                    break  # Останавливаемся, если дошли до 'OU=Groups'
                ou_parts.append(ou)
        # Если CN и OU найдены, формируем нужный формат
        ou_parts.reverse()
        if cn and ou_parts:
            groups_unsorted.append(f"{'\\'.join(ou_parts)}: {cn}")
    groups = "Членство в группах:\n" + "\n".join(sorted(groups_unsorted)) if groups_unsorted else ""
    vpn_group = [i for i in groups.splitlines() if i.startswith("VPN")] or ['Нет']
    vpn_ticket = ', '.join(user.get('TASS-VpnAccessTicket', ['Нет']))
    vpn = f"{", ".join(vpn_group)}\nЗаявка VPN: https://sd.corp.tass.ru/inframanager/?callNumber={vpn_ticket}" if "Нет" not in vpn_group else "VPN: Нет"
    distribution_lists = []
    for g in groups.splitlines():
        if g.startswith("Distribution"):
            distribution_lists.append("- " + g)
    distribution_lists = ("\n".join(sorted(distribution_lists))).replace("Distribution: ",
                                                                         "") if distribution_lists else "- Нет"
    location = ", ".join(user.get('l', ["Нет"]))
    street_address = ", ".join(user.get('streetAddress', []))
    office = ", ".join(user.get('physicalDeliveryOfficeName', []))
    is_working = "Да" if user.get('TASS-IsWorking') == [True] else "Нет"
    remote_work = "Да" if user.get('TASS-RemoteWork') == [True] else "Нет"
    kedo = "Да" if user.get('TASS-KDO') == [True] else "Нет"
    td = "Да" if user.get('TASS-TdEmployed') == [True] else "Нет"
    gph = "Да" if user.get('TASS-GphEmployed') == [True] else "Нет"
    ext = "Да" if user.get('TASS-External') == [True] else "Нет"
    bday = ", ".join(user.get('TASS-Birthday', ["Нет"]))

    # Возвращаем список из 2х строк: подробной информации и групп

    return [f"""
ФИО: {name}
Логин: {login}, почта: {email}
Должность: {title}
Подразделение: {department_full}
{vpn}
================================
Статус уз: {status}
Последний вход: {last_logon}
Последняя смена пароля: {pwd_last_set}
В уволенных: {dismissed}
Блокировка: {lock}
Истекает: {expire_time}
================================
Флаг TASS-IsWorking: {is_working}
Удаленная работа: {remote_work}
Согласие КЭДО: {kedo}
Трудовой договор: {td}
ГПХ договор: {gph}
Внешний пользователь: {ext}
================================
Телефон: {phone}
Мобильный: {mobile}
Адрес: {", ".join([location, street_address, office]).rstrip(", ")}
TelegramID: {tg_id}
FMC: {fmc}
День рождения: {bday}
Адреса почты: {mail_addresses}
БД почты: {maid_db}
================================
Общие пя:\n{shared_mb}
Группы рассылки:\n{distribution_lists}
================================
Путь в AD: {ad_path}
Хэш СНИЛС: {emp_num}
SID: {user_sid}
================================
""", groups]


