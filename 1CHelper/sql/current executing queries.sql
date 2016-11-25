select
	--s.login_name
	--,s.host_name
	--,s.program_name
	/*,*/
	r.query_hash
	,DB_NAME(r.database_id) as dbname
	,bf.drive
	,t.text
	,substring(t.text, r.statement_start_offset / 2 + 1, case when r.statement_end_offset = -1 then len(t.text) else (r.statement_end_offset / 2 + 1) end) as stmt
	,wt.wait_type
	,wt.wait_duration_ms
	,wt.resource_description
	,s.session_id
	--,s.host_process_id
	,wt.blocking_session_id
	--,bs.login_name
	--,bs.program_name
	,substring(bt.text, r.statement_start_offset / 2 + 1, case when br.statement_end_offset = -1 then len(bt.text) else (br.statement_end_offset / 2 + 1) end) as blocking_stmt
	--,cast(p.query_plan as xml) as query_plan
	--,*
from sys.dm_exec_requests r
inner join sys.dm_exec_sessions s on s.session_id = r.session_id
left join sys.dm_os_waiting_tasks wt on wt.session_id = r.session_id
left join sys.dm_exec_sessions bs on bs.session_id = wt.blocking_session_id
left join sys.dm_exec_requests br on br.session_id = bs.session_id
left join (select distinct database_id, SUBSTRING(physical_name,1,2) as drive from sys.master_files) bf on r.database_id = bf.database_id
cross apply sys.dm_exec_sql_text(r.sql_handle) t
cross apply sys.dm_exec_text_query_plan(r.plan_handle, r.statement_start_offset, r.statement_end_offset) p
outer apply sys.dm_exec_sql_text(br.sql_handle) bt
where
	r.session_id > 50
	and r.status in ('running', 'suspended')
	-- только ожидающие
	--and wt.wait_duration_ms is not null
	-- для конкретной базы данных
	--and DB_NAME(r.database_id) = '{My Database}'
order by
	query_hash
	,dbname
	,drive