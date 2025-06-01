import os

import pandas
import pymssql

# Функция формирования ежедневного отчета по заявкам, связанным с 1С
def form_onec_report():
    # Параметры подключения к базе данных ServiceDesk
    server = 'msk-sd-sdesk.corp.tass.ru\\general01'
    database = 'ServiceDesk2'
    ad_user = os.getenv("AD_USER_NAME")
    ad_password = os.getenv("AD_USER_PASSWORD")
 # SQL-запрос: выборка данных о заявках за вчерашний день, связанных с 1С
    query = """
        SELECT Number as 'Номер', UtcDateRegistered as 'Дата регистрации', c.EntityStateName as 'Статус',
        cl.FullName as 'Клиент', d.fullname as 'Подразделение клиента', CallSummaryName as 'Краткое описание',
        Description as 'Описание', Solution as 'Решение', u.FullName as 'Исполнитель', q.Name as 'Группа',
        cs.ServiceName as 'Сервис', cs.ServiceItemOrAttendanceName as 'Услуга' from Call c
        LEFT JOIN CallService cs ON c.CallServiceID = cs.ID
        LEFT JOIN view_UserFullName u ON c.ExecutorID = u.ID
        LEFT JOIN view_UserFullName cl ON c.ClientID = cl.ID
        LEFT JOIN view_DepartmentFullName d ON d.id = c.ClientSubdivisionID
        LEFT JOIN view_Queue q ON q.ID = c.QueueID
        WHERE (cs.ServiceItemOrAttendanceName LIKE '%1С%' or q.Name LIKE '%1С%')
        AND DAY(c.UtcDateRegistered) = DAY(DATEADD(DAY, -1, GETDATE()))
        AND MONTH(c.UtcDateRegistered) = MONTH(DATEADD(DAY, -1, GETDATE()))
        AND YEAR(c.UtcDateRegistered) = YEAR(DATEADD(DAY, -1, GETDATE()))
        AND NOT Removed = 1
        ORDER BY Number Asc
        """
# Установка соединения с базой данных
    connection = pymssql.connect(server, ad_user, ad_password, database)
 # Выполнение SQL-запроса и загрузка результатов
    df = pandas.read_sql_query(query, connection)
# Запись результатов в Excel-файл
    writer = pandas.ExcelWriter('1c_daily.xlsx')
    df.to_excel(writer, sheet_name='Заявки', index=False)

    writer.close()
    connection.close()
