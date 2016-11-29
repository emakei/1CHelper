select
	r.query_hash
	,DB_NAME(r.database_id) as dbname
	,bf.drive
	,wt.wait_type
	,wt.wait_duration_ms
	,wt.resource_description
	,s.session_id
	,wt.blocking_session_id
from sys.dm_exec_requests r
inner join sys.dm_exec_sessions s on s.session_id = r.session_id
left join sys.dm_os_waiting_tasks wt on wt.session_id = r.session_id
left join (select distinct database_id, SUBSTRING(physical_name,1,2) as drive from sys.master_files) bf on r.database_id = bf.database_id
where
	r.session_id > 50
	and r.status in ('running', 'suspended')
order by
	query_hash
	,dbname
	,drive
