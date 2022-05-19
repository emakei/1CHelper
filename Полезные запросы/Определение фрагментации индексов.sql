
select N'erp' batabase_name, * from sys.dm_db_index_physical_stats(DB_ID(N'erp'), NULL, NULL, NULL , 'DETAILED') where avg_fragmentation_in_percent >= 20 order by avg_fragmentation_in_percent desc
go

select N'accbe' batabase_name, * from sys.dm_db_index_physical_stats(DB_ID(N'accbe'), NULL, NULL, NULL , 'DETAILED') where avg_fragmentation_in_percent >= 20 order by avg_fragmentation_in_percent desc
go