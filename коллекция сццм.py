import winrm
import os


# # WMI
# c = wmi.WMI(computer=computer, user=user[0], password=password)
#
# os = c.Win32_OperatingSystem()
#
# for i in os:
#     print(i)
# # os.Reboot()

# WINRM
# computers = ['m2t-client01', 'mskt-client02']


def add_pc_to_sccm_collection_winrm(collectionid, computers):
    # URL сервера WinRM (SCCM сервер)
    server = 'http://msk-sccm-ss02.corp.tass.ru:5985/wsman'
     # Логин пользователя с доступом к SCCM через WinRM
    user = 'sccm_sd_service@corp.tass.ru'
    # Пароль берём из переменных окружения для безопасности
    password = os.getenv("AD_USER_PASSWORD")
    # Создаём сессию WinRM с аутентификацией NTLM, игнорируя валидацию сертификата
    session = winrm.Session(server, auth=(user, password), transport='ntlm', server_cert_validation='ignore')

    result = session.run_ps(
        # Формируем команду запуска PowerShell скрипта, передаём список компьютеров и ID коллекции SCCM
    # computers - список, склеиваем через запятую, collectionid - строка с ID коллекции
       pc_command = f". c:\\scripts\\PythonWINRM\\add_pcs_to_sccm_collection.ps1 -computers {','.join(computers)} -collectionid '{collectionid}'"
    )  

    return result.status_code, result.std_out, result.std_err

# (f". c:\\scripts\\test.ps1 -computers {','.join(computers)} -collectionid '{'CM100416'}'")
# (add_pc_to_sccm_collection_winrm(user, password, server, 'CM100416', computers))
