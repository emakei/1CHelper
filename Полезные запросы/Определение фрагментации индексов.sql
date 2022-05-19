
select N'BelGAZ' batabase_name, * from sys.dm_db_index_physical_stats(DB_ID(N'BelGAZ'), NULL, NULL, NULL , 'DETAILED') where avg_fragmentation_in_percent >= 20 order by avg_fragmentation_in_percent desc
go

select N'accbel' batabase_name, * from sys.dm_db_index_physical_stats(DB_ID(N'accbel'), NULL, NULL, NULL , 'DETAILED') where avg_fragmentation_in_percent >= 20 order by avg_fragmentation_in_percent desc
go