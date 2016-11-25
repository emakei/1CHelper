SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

DECLARE @dbid AS smallint;
use [{My Database}];
SET @dbid=DB_ID();

/*кто кого*/
SELECT DB_NAME(pr1.dbid) AS 'DB'
      ,pr1.spid AS 'ID жертвы'
      ,RTRIM(pr1.loginame) AS 'Login жертвы'
      ,pr2.spid AS 'ID виновника'
      ,RTRIM(pr2.loginame) AS 'Login виновника'
      ,pr1.program_name AS 'программа жертвы'
      ,pr2.program_name AS 'программа виновника'
      ,txt.[text] AS 'Запрос виновника'
FROM   MASTER.dbo.sysprocesses pr1(NOLOCK)
       JOIN MASTER.dbo.sysprocesses pr2(NOLOCK)
            ON  (pr2.spid = pr1.blocked)
       OUTER APPLY sys.[dm_exec_sql_text](pr2.[sql_handle]) AS txt
WHERE  pr1.blocked <> 0

/* Кто что блокирует */
SELECT s.[nt_username]
      ,request_session_id
      ,tran_locks.[request_status]
      ,rd.[Description] + ' (' + tran_locks.resource_type + ' ' + tran_locks.request_mode + ')' [Object]
      ,txt_blocked.[text]
      ,COUNT(*) [COUNT]
FROM   sys.dm_tran_locks AS tran_locks WITH (NOLOCK)
       JOIN sys.sysprocesses AS s WITH (NOLOCK)
            ON  tran_locks.request_session_id = s.[spid]
       JOIN (
                SELECT 'KEY' AS sResource_type
                      ,p.[hobt_id] AS [id]
                      ,QUOTENAME(o.name) + '.' + QUOTENAME(i.name) AS [Description]
                FROM   sys.partitions p
                       JOIN sys.objects o
                            ON  p.object_id = o.object_id
                       JOIN sys.indexes i
                            ON  p.object_id = i.object_id
                            AND p.index_id = i.index_id
                UNION ALL
                SELECT 'RID' AS sResource_type
                      ,p.[hobt_id] AS [id]
                      ,QUOTENAME(o.name) + '.' + QUOTENAME(i.name) AS [Description]
                FROM   sys.partitions p
                       JOIN sys.objects o
                            ON  p.object_id = o.object_id
                       JOIN sys.indexes i
                            ON  p.object_id = i.object_id
                            AND p.index_id = i.index_id
                UNION ALL
                SELECT 'PAGE'
                      ,p.[hobt_id]
                      ,QUOTENAME(o.name) + '.' + QUOTENAME(i.name)
                FROM   sys.partitions p
                       JOIN sys.objects o
                            ON  p.object_id = o.object_id
                       JOIN sys.indexes i
                            ON  p.object_id = i.object_id
                            AND p.index_id = i.index_id
               
                UNION ALL
                SELECT 'OBJECT'
                      ,o.[object_id]
                      ,QUOTENAME(o.name)
                FROM   sys.objects o
            ) AS RD
            ON  RD.[sResource_type] = tran_locks.resource_type
            AND RD.[id] = tran_locks.resource_associated_entity_id
       OUTER APPLY sys.[dm_exec_sql_text](s.[sql_handle]) AS txt_Blocked
WHERE  (
           tran_locks.request_mode = 'X'
           AND tran_locks.resource_type = 'OBJECT'
       )
       OR  tran_locks.[request_status] = 'WAIT'
GROUP BY
       s.[nt_username]
      ,request_session_id
      ,tran_locks.[request_status]
      ,rd.[Description] + ' (' + tran_locks.resource_type + ' ' + tran_locks.request_mode + ')'
      ,txt_blocked.[text]
ORDER BY
       6 DESC
       

IF EXISTS ( SELECT  Name

            FROM    tempdb..sysobjects

            WHERE   name LIKE '#LOCK_01_01%' )
    DROP TABLE #LOCK_01_01


CREATE TABLE #LOCK_01_01

    (

      spid INT,

      dbid INT,

      ObjId INT,

      IndId SMALLINT,

      Type VARCHAR(20),

      Resource VARCHAR(50),

      Mode VARCHAR(20),

      Status VARCHAR(20)

    )

INSERT  INTO #LOCK_01_01

EXEC sp_lock


select OBJECT_NAME(ObjId) as [Имя объекта], Mode [Тип блокировки (код)],

CASE
     WHEN Mode='Sch-S' THEN 'Блокировка стабильности схемы. Гарантирует, что элемент схемы, такой как таблица или индекс, не будет удален до тех пор, пока сеанс связи удерживает блокировку стабильности схемы на данный элемент схемы;'

 WHEN Mode='Sch-М' THEN '= Блокировка изменения схемы. Должен поддерживаться любым сеансом связи, во время которого предполагается изменить схему данного ресурса. Гарантирует, что другие сеансы не имеют ссылок на обозначенный объект;'

 WHEN Mode='S' THEN 'S = Коллективная блокировка. Удерживающему сеансу предоставлен коллективный доступ к ресурсу;'

 WHEN Mode='U' THEN 'U = Блокировка обновления. Указывает блокировку обновления, полученную на ресурсы, которые со временем могут быть обновлены. Используется для предотвращения общей формы взаимоблокировки, которая возникает, когда множество сеансов блокируют ресурсы для потенциального обновления в последующее время;'

 WHEN Mode='X' THEN 'X = Монопольная блокировка. Удерживающему сеансу предоставлен исключительный доступ к ресурсу;'

 WHEN Mode='IS' THEN 'IS = Блокировка с намерением коллективного доступа. Указывает намерение поместить S блокировки на некоторые подчиненные ресурсы в иерархии блокировок;'

 WHEN Mode='IU' THEN 'IU = Блокировка с намерением обновления. Указывает намерение поместить U блокировки на некоторые подчиненные ресурсы в иерархии блокировок;'

 WHEN Mode='IX' THEN 'IX = Блокировка с намерением монопольного доступа. Указывает намерение поместить X блокировки на некоторые подчиненные ресурсы в иерархии блокировок;'

 WHEN Mode='SIU' THEN 'SIU = Коллективная блокировка с намерением обновления. Указывает коллективный доступ к ресурсу с намерением получения блокировок обновления на подчиненные ресурсы в иерархии блокировок;'

 WHEN Mode='SIX' THEN 'SIX = Коллективная блокировка с намерением монопольного доступа. Указывает коллективный доступ к ресурсу с намерением получения монопольных блокировок на подчиненные ресурсы в иерархии блокировок;'

 WHEN Mode='UIX' THEN 'UIX = Блокировка обновления с намерением монопольного доступа. Указывает блокировку обновления ресурса с намерением получения монопольных блокировок на подчиненные ресурсы в иерархии блокировок;'

 WHEN Mode='BU' THEN 'BU = Блокировка массового обновления. Используется для массовых операций;'
     --[ ELSE else_result_expression ]
END as [Тип блокировки]

,syspr.spid, syspr.dbid, syspr.open_tran, syspr.status, syspr.hostprocess, syspr.loginame, syspr.hostname

 From
#LOCK_01_01

inner join master.dbo.sysprocesses as syspr
on syspr.spid = #LOCK_01_01.spid and syspr.dbid = #LOCK_01_01.dbid

where
#LOCK_01_01.Type = 'TAB'
and
#LOCK_01_01.dbid = @dbid     

/* Чем занят сервер*/
SELECT s.[spid]
      ,s.[loginame]
      ,s.[open_tran]
      ,s.[blocked]
      ,s.[waittime]
      ,s.[cpu]
      ,s.[physical_io]
      ,s.[memusage]
       INTO #sysprocesses
FROM   sys.[sysprocesses] s

WAITFOR DELAY '00:00:01'

SELECT txt.[text]
      ,s.[spid]
      ,s.[loginame]
      ,s.[hostname]
      ,DB_NAME(s.[dbid]) [db_name]
      ,SUM(s.[waittime] -ts.[waittime]) [waittime]
      ,SUM(s.[cpu] -ts.[cpu]) [cpu]
      ,SUM(s.[physical_io] -ts.[physical_io]) [physical_io]
      ,s.[program_name]
FROM   sys.[sysprocesses] s
       JOIN #sysprocesses ts
            ON  s.[spid] = ts.[spid]
            AND s.[loginame] = ts.[loginame]
       OUTER APPLY sys.[dm_exec_sql_text](s.[sql_handle]) AS txt
WHERE  s.[cpu] -ts.[cpu]
       + s.[physical_io] -ts.[physical_io]
       > 500
       OR  (s.[waittime] -ts.[waittime]) > 3000
GROUP BY
       txt.[text]
      ,s.[spid]
      ,s.[loginame]
      ,s.[hostname]
      ,DB_NAME(s.[dbid])
      ,s.[program_name]
ORDER BY
       [physical_io] DESC
       
DROP TABLE #sysprocesses