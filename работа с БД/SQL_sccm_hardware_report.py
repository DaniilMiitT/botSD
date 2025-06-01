import os

import pandas
import pymssql

# Функция формирования отчета из базы SCCM
def form_sccmhw_report():
     # Данные для подключения к БД
    server = 'msk-sccm-ss02.corp.tass.ru'
    database = 'CM_CM1'
    ad_user = os.getenv("AD_USER_NAME")
    ad_password = os.getenv("AD_USER_PASSWORD")

# SQL-запрос: сбор информации по компьютерам в коллекции CM100068 за последние 30 дней
    query = """
        DECLARE @Today AS DATE
        DECLARE @CollectionID nvarchar(8)
        SET @Today = GETDATE()

        DECLARE @BackInTime AS DATE
        SET @BackInTime = DATEADD(DAY, -30, @Today )
        SET @CollectionID = 'CM100068'

        SELECT DISTINCT
         --dfc.CollectionID,
         --SYS.ResourceID,
         SYS.Name0 'Name',
         --SYS.AD_Site_Name0 'ADSite',
         CS.UserName0 'User Name',
         CASE
         WHEN U.TopConsoleUser0 = '-1' OR U.TopConsoleUser0 IS NULL THEN 'N/A'
         ELSE U.TopConsoleUser0
         END AS TopUser,
         REPLACE((REPLACE((REPLACE((REPLACE((REPLACE((REPLACE (OS.Caption0, 'Microsoft Windows','Win')),'Enterprise','EE') ),'Standard','ST')),'Microsoft®','')),'Server','SRV')),'Windows','Win') OS,
         --REPLACE (OS.CSDVersion0,'Service Pack','SP') 'Service Pack',
         CS.Manufacturer0 'Manufacturer',
         CS.Model0 Model,
         BIOS.SerialNumber0 'Serial Number',
         CONVERT (DATE,BIOS.ReleaseDate0) AS BIOSDate,
         BIOS.SMBIOSBIOSVersion0 AS BIOSVersion,
         CPU.[CPU Name],
         --CPU.Manufacturer AS 'CPU Man.',
         --CPU.[Number of CPUs] AS '# of CPUs',
         --CPU.[Number of Cores per CPU] AS '# of Cores per CPU',
         CPU.[Logical CPU Count] AS 'Logical CPU Count',
         --(SELECT CONVERT(DATE,SYS.Creation_Date0)) 'Managed Date',
         SUM(ISNULL(RAM.Capacity0,0)) 'Memory (MB)',
         COUNT(RAM.ResourceID) '# Memory Slots',
         REPLACE (cs.SystemType0,'-based PC','') 'Type',
         SUM(D.Size0) / 1024 AS 'Disk Size GB',
         CONVERT(VARCHAR(26), OS.LastBootUpTime0, 100) AS 'Last Reboot Date/Time',
         CONVERT(VARCHAR(26), OS.InstallDate0, 101) AS 'Install Date',
         --CONVERT(VARCHAR(26), WS.LastHWScan, 101) AS 'Last Hardware Inventory',
         CONVERT(VARCHAR(26), CH.LastOnline, 101) AS 'Last Seen Online',
         --SYS.Client_Version0 as 'SCCM Agent Version',
         --US.ScanTime AS ' Windows Updates Scan Time' ,
         --US.LastErrorCode AS ' Windows Updates Last Error Code' ,
         --US.LastScanPackageLocation AS ' Windows Updates Last Package Location' ,
         CASE SE.ChassisTypes0
         WHEN '1' THEN 'Other'
         WHEN '2' THEN 'Unknown'
         WHEN '3' THEN 'Desktop'
         WHEN '4' THEN 'Low Profile Desktop'
         WHEN '5' THEN 'Pizza Box'
         WHEN '6' THEN 'Mini Tower'
         WHEN '7' THEN 'Tower'
         WHEN '8' THEN 'Portable'
         WHEN '9' THEN 'Laptop'
         WHEN '10' THEN 'Notebook'
         WHEN '11' THEN 'Hand Held'
         WHEN '12' THEN 'Docking Station'
         WHEN '13' THEN 'All in One'
         WHEN '14' THEN 'Sub Notebook'
         WHEN '15' THEN 'Space-Saving'
         WHEN '16' THEN 'Lunch Box'
         WHEN '17' THEN 'Main System Chassis'
         WHEN '18' THEN 'Expansion Chassis'
         WHEN '19' THEN 'SubChassis'
         WHEN '20' THEN 'Bus Expansion Chassis'
         WHEN '21' THEN 'Peripheral Chassis'
         WHEN '22' THEN 'Storage Chassis'
         WHEN '23' THEN 'Rack Mount Chassis'
         WHEN '24' THEN 'Sealed-Case PC'
         ELSE 'Undefinded'
         END AS 'PC Type'
        FROM
         v_R_System SYS
         INNER JOIN (
         SELECT
         Name0,
         MAX(Creation_Date0) AS Creation_Date
         FROM
         dbo.v_R_System
         GROUP BY
         Name0
         ) AS CleanSystem
         ON SYS.Name0 = CleanSystem.Name0 AND SYS.Creation_Date0 = CleanSystem.Creation_Date
         LEFT JOIN v_GS_COMPUTER_SYSTEM CS
         ON SYS.ResourceID=cs.ResourceID
         LEFT JOIN v_GS_PC_BIOS BIOS
         ON SYS.ResourceID=bios.ResourceID
         LEFT JOIN (
         SELECT
         A.ResourceID,
         MAX(A.[InstallDate0]) AS [InstallDate0]
         FROM
         v_GS_OPERATING_SYSTEM A
         GROUP BY
         A.ResourceID
         ) AS X
         ON SYS.ResourceID = X.ResourceID
         INNER JOIN v_GS_OPERATING_SYSTEM OS
         ON X.ResourceID=OS.ResourceID AND X.InstallDate0 = OS.InstallDate0
         LEFT JOIN v_GS_PHYSICAL_MEMORY RAM
         ON SYS.ResourceID=ram.ResourceID
         LEFT OUTER JOIN dbo.v_GS_LOGICAL_DISK D
         ON SYS.ResourceID = D.ResourceID AND D.DriveType0 = 3
         LEFT OUTER JOIN v_GS_SYSTEM_CONSOLE_USAGE_MAXGROUP U
         ON SYS.ResourceID = U.ResourceID
         LEFT JOIN dbo.v_GS_SYSTEM_ENCLOSURE SE ON SYS.ResourceID = SE.ResourceID
         LEFT JOIN dbo.v_GS_ENCRYPTABLE_VOLUME En ON SYS.ResourceID = En.ResourceID
         LEFT JOIN dbo.v_GS_WORKSTATION_STATUS WS ON SYS.ResourceID = WS.ResourceID
         LEFT JOIN v_CH_ClientSummary CH
         ON SYS.ResourceID = CH.ResourceID
         LEFT JOIN (
         SELECT
         DISTINCT(CPU.SystemName0) AS [System Name],
         CPU.Manufacturer0 AS Manufacturer,
         CPU.ResourceID,
         CPU.Name0 AS [CPU Name],
         COUNT(CPU.ResourceID) AS [Number of CPUs],
         CPU.NumberOfCores0 AS [Number of Cores per CPU],
         CPU.NumberOfLogicalProcessors0 AS [Logical CPU Count]
         FROM [dbo].[v_GS_PROCESSOR] CPU
         GROUP BY
         CPU.SystemName0,
         CPU.Manufacturer0,
         CPU.Name0,
         CPU.NumberOfCores0,
         CPU.NumberOfLogicalProcessors0,
         CPU.ResourceID
         ) CPU
         ON CPU.ResourceID = SYS.ResourceID
         LEFT JOIN v_UpdateScanStatus US
         ON US.ResourceID = SYS.ResourceID
         inner join dbo.v_FullCollectionMembership dfc
         on dfc.ResourceID = sys.ResourceID
        WHERE SYS.obsolete0=0 AND SYS.client0=1 AND SYS.obsolete0=0 AND SYS.active0=1 AND dfc.CollectionID = @CollectionID and
         CH.LastOnline BETWEEN @BackInTime AND GETDATE()
         GROUP BY
         dfc.CollectionID,
         SYS.Creation_Date0 ,
         SYS.Name0 ,
         SYS.ResourceID ,
         SYS.AD_Site_Name0 ,
         CS.UserName0 ,
         REPLACE((REPLACE((REPLACE((REPLACE((REPLACE((REPLACE (OS.Caption0, 'Microsoft Windows','Win')),'Enterprise','EE') ),'Standard','ST')),'Microsoft®','')),'Server','SRV')),'Windows','Win'),
         REPLACE (OS.CSDVersion0,'Service Pack','SP'),
         CS.Manufacturer0 ,
         CS.Model0 ,
         BIOS.SerialNumber0 ,
         REPLACE (cs.SystemType0,'-based PC','') ,
         CONVERT(VARCHAR(26), OS.LastBootUpTime0, 100) ,
         CONVERT(VARCHAR(26), OS.InstallDate0, 101) ,
         CONVERT(VARCHAR(26), WS.LastHWScan, 101),
         CASE
         WHEN U.TopConsoleUser0 = '-1' OR U.TopConsoleUser0 IS NULL THEN 'N/A'
         ELSE U.TopConsoleUser0
         END,
         CPU.Manufacturer,
         CPU.[Number of CPUs] ,
         CPU.[Number of Cores per CPU],
         CPU.[Logical CPU Count],
         CPU.[CPU Name],
         US.ScanTime ,
         US.LastErrorCode ,
         US.LastScanPackageLocation ,
         CASE SE.ChassisTypes0
         WHEN '1' THEN 'Other'
         WHEN '2' THEN 'Unknown'
         WHEN '3' THEN 'Desktop'
         WHEN '4' THEN 'Low Profile Desktop'
         WHEN '5' THEN 'Pizza Box'
         WHEN '6' THEN 'Mini Tower'
         WHEN '7' THEN 'Tower'
         WHEN '8' THEN 'Portable'
         WHEN '9' THEN 'Laptop'
         WHEN '10' THEN 'Notebook'
         WHEN '11' THEN 'Hand Held'
         WHEN '12' THEN 'Docking Station'
         WHEN '13' THEN 'All in One'
         WHEN '14' THEN 'Sub Notebook'
         WHEN '15' THEN 'Space-Saving'
         WHEN '16' THEN 'Lunch Box'
         WHEN '17' THEN 'Main System Chassis'
         WHEN '18' THEN 'Expansion Chassis'
         WHEN '19' THEN 'SubChassis'
         WHEN '20' THEN 'Bus Expansion Chassis'
         WHEN '21' THEN 'Peripheral Chassis'
         WHEN '22' THEN 'Storage Chassis'
         WHEN '23' THEN 'Rack Mount Chassis'
         WHEN '24' THEN 'Sealed-Case PC'
         ELSE 'Undefinded'
         END,
         CONVERT (DATE,BIOS.ReleaseDate0) ,
         BIOS.SMBIOSBIOSVersion0 ,
         SYS.Client_Version0 ,
         CONVERT(VARCHAR(26) ,CH.LastOnline, 101)
         ORDER BY SYS.Name0
        """
 # Подключение к базе данных
    connection = pymssql.connect(server, ad_user, ad_password, database)
# Выполнение SQL-запроса и получение результата в виде DataFrame
    df = pandas.read_sql_query(query, connection)
# Запись результата в Excel
    writer = pandas.ExcelWriter('sccm_hw_report.xlsx')
    df.to_excel(writer, sheet_name='Col CM100068', index=False)
# Закрытие Excel-файла и соединения с БД
    writer.close()
    connection.close()
