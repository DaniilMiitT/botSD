import json
import os
import re
import tempfile

import pefile
import requests.exceptions
from bs4 import BeautifulSoup
from packaging import version

from sccm_actions import sccm_get_app_versions, sccm_get_latest_app_version


def get_file_version(url):
    # Загружает исполняемый файл по URL, извлекает и возвращает его версию
    print(f"Start getting file version from {url}")
    response = requests.get(url, allow_redirects=True)
    response.raise_for_status()  # Проверка на ошибки HTTP 

    final_url = response.url
    print(f"Final download URL: {final_url}")

    # Проверка типа содержимого (должен быть .exe)
    content_type = response.headers.get('Content-Type', '').lower()
    if 'application/octet-stream' not in content_type and 'application/x-msdownload' not in content_type:
        raise ValueError(f"Unexpected Content-Type: {content_type}. Expected an executable file.")

    #Сохраняем содержимое во временный файл
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(response.content)
        temp_file_path = temp_file.name

    try:
        #Извлекаем версию из PE-файла
        pe = pefile.PE(temp_file_path)

        # Используем фиксированную структуру версии, если доступна
        if hasattr(pe, 'VS_FIXEDFILEINFO'):
            fixed_info = pe.VS_FIXEDFILEINFO[0]
            version_ms = fixed_info.FileVersionMS
            version_ls = fixed_info.FileVersionLS
            version = f"{(version_ms >> 16)}.{(version_ms & 0xFFFF)}.{(version_ls >> 16)}.{(version_ls & 0xFFFF)}"
            return version
        else:
            # Если нет, пробуем через строковую таблицу
            for file_info in pe.FileInfo:
                if file_info.Key.decode() == 'StringFileInfo':
                    for string_table in file_info.StringTable:
                        if b'FileVersion' in string_table.entries:
                            return string_table.entries[b'FileVersion'].decode('utf-8').strip()
            raise ValueError("FileVersion entry not found in the executable's metadata.")
    except pefile.PEFormatError:
        raise ValueError("Downloaded file is not a valid Windows executable.")
    finally:
        os.remove(temp_file_path)# Удаляем временный файл


def check_new_app_versions():
    print("Getting sccm apps...")
    sccm_apps = sccm_get_app_versions()# Получаем версии из SCCM
    # Получение версий с сайта Ninite
    print("Loading ninite...")
    page = requests.get("https://ninite.com")
    parsedhtml = BeautifulSoup(page.text, 'lxml')

    result = {}
    # Парсинг версий по ключевым описаниям
    for text in parsedhtml.find_all('p'):
        if "Fast Browser by Google" in text.text:
            result['Google Chrome'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if "Great Video Player" in text.text:
            result['VLC'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if "Extensible Browser" in text.text:
            result['Firefox'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if "Video decoders plus Media Player Classic" in text.text:
            result['K-Lite Mega Codec Pack'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if "Password Manager" in text.text:
            result['KeePass'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if "Video Conference" in text.text:
            result['ZOOM'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if text.text.startswith("Programming Language"):
            result['Python'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if text.text.startswith("Great Compression App"):
            result['7-Zip'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if text.text.startswith("Internet Telephone"):
            result['Skype'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if text.text.startswith("SCP Client"):
            result['WinSCP'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if text.text.startswith("Music/Media Manager"):
            result['iTunes'] = {'Version': re.sub("[^0-9.]", "", text.text)}
        if text.text.startswith("Remote Access Tool"):
            result['TeamViewer'] = {'Version': re.sub("[^0-9.]", "", text.text)}

        # Дополнительный парсинг по названиям
        for i in parsedhtml.find_all(string=re.compile("Notepad")):
            result['Notepad++'] = {'Version': re.sub("[^0-9.]", "", str(i.findNext('p')))}
        for i in parsedhtml.find_all(string=re.compile("FileZilla")):
            result['FileZilla'] = {'Version': re.sub("[^0-9.]", "", str(i.findNext('p')))}
        for i in parsedhtml.find_all(string=re.compile("Krita")):
            result['Krita'] = {'Version': re.sub("[^0-9.]", "", str(i.findNext('p')))}
        for i in parsedhtml.find_all(string=re.compile("Paint.NET")):
            result['Paint.NET'] = {
                'Version': re.sub("[^0-9.]", "", str(i.findNext('p')).replace(" (requires .NET 4.5)", ""))}
        for i in parsedhtml.find_all(string=re.compile("GIMP")):
            result['GIMP'] = {'Version': re.sub("[^0-9.]", "", str(i.findNext('p')))}
        for i in parsedhtml.find_all(string=re.compile("Inkscape")):
            result['Inkscape'] = {'Version': re.sub("[^0-9.]", "", str(i.findNext('p')))}

    # Adobe Reader
    print("Adobe Reader")
    adobe_reader = requests.get("https://www.adobe.com/devnet-docs/acrobatetk/tools/ReleaseNotesDC")
    adobe_reader = BeautifulSoup(adobe_reader.text, 'lxml')
    adobe_reader_version = []
    for i in adobe_reader.find_all('a'):
        if "update" in i.text:
            adobe_reader_version.append(i.text)
            adobe_reader = re.split(" ", adobe_reader_version[0])[0]
    result['Acrobat Reader DC'] = {'Version': re.sub("[^0-9.]", "", adobe_reader)}


    # AnyConnect 
    print("Anyconnect")
    anyconnect_uri = "https://www.cisco.com/c/en/us/td/docs/security/vpn_client/anyconnect/anyconnect410/release/notes/release-notes-anyconnect-4-10.html"

    try:
        class Proxy:
            user = os.getenv("PROXY_USER")
            password = os.getenv("PROXY_PASSWORD")
            url = os.getenv("PROXY_URL")
            port = os.getenv("PROXY_PORT")

        proxy = {
            "http": f"http://{Proxy.user}:{Proxy.password}@{Proxy.url}:{Proxy.port}",
            "https": f"http://{Proxy.user}:{Proxy.password}@{Proxy.url}:{Proxy.port}"
        }
        r = requests.get(anyconnect_uri, verify=False, timeout=5, proxies=proxy)
        print(r.status_code, r.url)
        anyconnect = BeautifulSoup(r.text, 'lxml')
        for i in anyconnect.find_all(string=re.compile("AnyConnect(.*)New Features"), limit=1):
            result['Anyconnect'] = {'Version': re.sub("[^0-9.]", "", i)}
    except requests.exceptions.ConnectionError as e:
        print({f'Error in contacting {anyconnect_uri}': str(e)})
        result['Anyconnect'] = {'Version': '0.0.0.0'}

    # Telegram
    print('Telegram')
    telegram = requests.get("https://api.github.com/repos/{}/releases/latest".format("telegramdesktop/tdesktop"))
    telegram = json.loads(telegram.content)['tag_name']
    result['Telegram'] = {'Version': re.sub("[^0-9.]", "", telegram)}

    # Kaspersky Endpoint Security
    print('KES')
    kes = requests.get("https://support.kaspersky.ru/kes12/new")
    kes = BeautifulSoup(kes.text, "lxml")
    kes_version = kes.find("p", class_="introheading")
    result['Kaspersky Endpoint Security'] = {
        'Version': kes_version.get_text().split(" ")[1].strip()}

    # Kaspersky Security Center
    print('KNA')
    ksc = requests.get("https://support.kaspersky.ru/15899")
    ksc = BeautifulSoup(ksc.text, "lxml")
    ksc_version = ksc.find_all('div', 'cont', limit=1)
    result['Агент администрирования Kaspersky'] = {
        'Version': ksc_version[0].find_all_next("p", limit=1)[0].text.split(" ")[-1]}

    # Справки БК
    print('Справки БК')
    gscx_params = {
        'key': os.getenv("GOOLE_SEARCH_API_KEY"),
        'cx': os.getenv("GOOGLE_SEARCH_CX"),
        'q': "справки бк",
        'num': 10  # Количество результатов для проверки
    }
    gscx_url = "https://www.googleapis.com/customsearch/v1"

    try:
        response = requests.get(gscx_url, params=gscx_params)
        data = response.json()
        sbk_version = data['items'][0]['snippet']
        match = re.search(r'(\d+\.\d+\.\d+)', sbk_version)
        if match:
            sbk_version = match.group(1)
        else:
            sbk_version = "0.0.0.0"
    except Exception as e:
        print(e)
        sbk_version = "0.0.0.0"

    result['Справки БК'] = {'Version': str(sbk_version)}

    # Nextcloud
    print('Nextcloud')
    ncc = requests.get("https://api.github.com/repos/{}/releases/latest".format("nextcloud/desktop"))
    ncc = json.loads(ncc.content)['tag_name']
    result['Nextcloud'] = {'Version': re.sub("[^0-9.]", "", ncc)}

    # R7-Office
    print('R7 Office')
    r7o = requests.get("https://support.r7-office.ru/desktop_editors/general_de/release-notes-desktop-editors/")
    r7o = BeautifulSoup(r7o.text, 'html.parser')
    r7o = r7o.find('h2', string=re.compile(r"^Версия"))
    r7o_version = r7o.text.split(" ")[1].strip()
    result['Р7-Офис'] = {'Version': str(r7o_version)}

    # Yandex Browser
    print('Yandex Browser')
    download_url = "https://browser.yandex.ru/download?partner_id=corp-common"
    try:
        yb_version = get_file_version(download_url)
        print(f"File Version: {yb_version}")
        result['Yandex Browser'] = {'Version': str(yb_version)}
    except Exception as e:
        print(f"Error: {e}")

    # Формируем текст сообщения и сравниваем с SCCM
    result = dict(sorted(result.items()))

    print(result)

    text = "*New versions:*\r\n"
    for k, v in result.items():
        sccm = sccm_get_latest_app_version(sccm_apps, k)
        if sccm and "Таймаут" not in v['Version'] and version.parse(v['Version']) > version.parse(str(sccm['Version'])):
            text += k + ": *" + v['Version'] + f"* (SCCM: *{sccm['Version']}*)\r\n"
            result[k]['SCCM'] = str(sccm['Version'])
        else:
            text += k + ": *" + v['Version'] + "*\r\n"
# Сохраняем результат в JSON-файл
    with open('AppVersionLog.json', 'w') as outfile:
        json.dump(result, outfile)

    return str(text)

