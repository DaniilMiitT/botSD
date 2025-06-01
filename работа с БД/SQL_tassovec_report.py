import os

import pandas # Импорт библиотеки для работы с таблицами и Excel
import pymssql# Импорт библиотеки для подключения к MS SQL Server

# Функция для формирования Excel-отчета по заявкам, связанным с тассОВЦЕМ
def form_tassovec_report():
    server = 'msk-sd-sdesk.corp.tass.ru\\general01'# Имя SQL-сервера
    database = 'ServiceDesk2'
    ad_user = os.getenv("AD_USER_NAME")
    ad_password = os.getenv("AD_USER_PASSWORD")
    # SQL-запрос для выборки данных по заявкам за прошлый месяц, связанных с тассОВЦЕМ
    query = (
        "SELECT Number as 'Номер', UtcDateRegistered as 'Дата регистрации', c.EntityStateName as 'Статус',  cl.FullName as 'Клиент', d.fullname as 'Подразделение клиента', CallSummaryName as 'Краткое описание', Description as 'Описание', Solution as 'Решение', u.FullName as 'Исполнитель', cs.ServiceName as 'Сервис', cs.ServiceItemOrAttendanceName as 'Услуга' from Call c \
         LEFT JOIN CallService cs ON c.CallServiceID = cs.ID \
         LEFT JOIN view_UserFullName u ON c.ExecutorID = u.ID \
         LEFT JOIN view_UserFullName cl ON c.ClientID = cl.ID \
         LEFT JOIN view_DepartmentFullName d ON d.id = c.ClientSubdivisionID \
         WHERE (cs.ServiceName LIKE '%ТАССОВЕЦ%' OR cs.ServiceItemOrAttendanceName LIKE '%КЭП%') \
         AND MONTH(c.UtcDateRegistered) = MONTH(DATEADD(MONTH, -1, GETDATE())) \
         AND YEAR(c.UtcDateRegistered) = YEAR(DATEADD(MONTH, -1, GETDATE())) \
         AND NOT Removed = 1 \
         ORDER BY Number Asc")
# Устанавливаем соединение с SQL-сервером
    connection = pymssql.connect(server, ad_user, ad_password, database)
# Выполняем SQL-запрос и сохраняем результат
    df = pandas.read_sql_query(query, connection)
# Создаем Excel-файл и записываем туда данные
    writer = pandas.ExcelWriter('tassovec.xlsx')
    df.to_excel(writer, sheet_name='Заявки', index=False)
 # Закрываем Excel-файл и соединение с базой
    writer.close()
    connection.close()
