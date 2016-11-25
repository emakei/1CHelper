select *
from sys.dm_os_latch_stats
order by wait_time_ms

select *
from sys.dm_os_wait_stats
order by wait_time_ms desc