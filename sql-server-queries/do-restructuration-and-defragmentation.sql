USE tempdb
SET NOCOUNT ON
SET QUOTED_IDENTIFIER ON
GO

IF OBJECT_ID('handdleFragmentationIndexes', 'P') IS NOT NULL
BEGIN
  DROP  PROCEDURE handdleFragmentationIndexes
  PRINT 'handdleFragmentationIndexes already exist and was DROPED'
END
ELSE
  PRINT 'handdleFragmentationIndexes is not exist and will be CREATED'
GO

CREATE PROCEDURE handdleFragmentationIndexes(@debugMode bit = 0, @mode varchar(50) = 'DETAILED', @databaseName varchar(50) = 'ALL')
AS
BEGIN
  
  -- @mode =
  -- LIMITED (default): This mode is the fastest and scans the smallest number of pages. It scans all pages for a heap, but only scans the parent-level pages, which means, the pages above the leaf level, for an index
  -- SAMPLED: This mode returns statistics base on a one percent sample of the entire page in the index or heap. If the index or heap has fewer tha 10 000 pages, DETAILED mode is used instead of SAMPLED.
  -- DETAILED: Detailed mode scans all pages and returns all statistics. Be careful, from LIMITED to SAMPLED to DETAILED, the mode are progressively slower, because more work is performed in each. In my script I am useng this one.

  SET NOCOUNT ON

  DECLARE @objectid int
  DECLARE @indexid int
  DECLARE @partitioncount bigint
  DECLARE @schemaname sysname
  DECLARE @objectname sysname
  DECLARE @indexname sysname
  DECLARE @partitionnum bigint
  DECLARE @partitions bigint
  DECLARE @frag float
  DECLARE @spaceUsed float
  DECLARE @command varchar(MAX)
  DECLARE @command2 varchar(MAX)
  DECLARE @commandStatistics varchar(8000)
  DECLARE @database_id int
  DECLARE @t1 DATETIME
  DECLARE @t2 DATETIME
  DECLARE @time_diff int
  
  DECLARE @db_id int
  DECLARE @db_name varchar(50)
  DECLARE @allow_page_locks int
  
  SET @t1 = GETDATE()
  
  CREATE TABLE #med_info_index(object_id int, index_id int, partition_number int, avg_fragmentation_in_percent float, [db_name] varchar(150), avg_page_space_used_in_percent float)
  CREATE TABLE #med_info_index_detail([objectid] int, [indexid] int, [partitionnum] int, [frag] float, [objectname] varchar(150), [schemaname] varchar(150), [db_name] varchar(150), [indexname] varchar(150), [partitioncount] int, [allow_page_locks] int, [avg_page_space_used_in_percent] float)

  -- STEP 1
  PRINT 'Step 1 starting... ' + CAST(GETDATE() AS nvarchar(max))
  PRINT '- Get all indexes with a fragmentation over than 20% for database: ' + @databaseName

  IF @databaseName = 'ALL'
    SET @database_id = NULL
  ELSE
    SELECT @database_id = ISNULL(DB_ID(@databaseName), 0)

  INSERT #med_info_index(object_id, index_id, partition_number, avg_fragmentation_in_percent, [db_name], [avg_page_space_used_in_percent])
  	SELECT object_id, index_id, partition_number, avg_fragmentation_in_percent, db_name(database_id) as [db_name], ISNULL(avg_page_space_used_in_percent,100)
  	FROM sys.dm_db_index_physical_stats(@database_id, NULL, NULL, NULL, @mode)
  	WHERE avg_fragmentation_in_percent >= 20
  		AND database_id IN (SELECT database_id
  		    		    FROM sys.databases
  				    WHERE [state] = 0 /* ONLINE */ AND is_read_only = 0 AND database_id > 4 /* SKIP SYSTEM DB */
  				    -- Documentation : http://msdn.microsoft.com/en-us/library/ms178534.aspx
  				    )
  
  -- STEP 2
  PRINT 'Step 2 starting... ' + CAST(GETDATE() AS nvarchar(max))
  PRINT '- Get all objects details about the index with a fragmentation over 20% for database: ' + @databaseName
  
  DECLARE dbList CURSOR FOR
  	SELECT object_id, index_id, partition_number, avg_fragmentation_in_percent, [db_name], avg_page_space_used_in_percent
  	FROM #med_info_index
  FOR READ ONLY
  
  OPEN dbList
  
  FETCH NEXT FROM dbList INTO @objectid, @indexid, @partitionnum, @frag, @db_name, @spaceused
  WHILE @@FETCH_STATUS = 0
  BEGIN
    SET @command2 = 'DECLARE @objectname varchar(50), @indexname varchar(50), @schemaname varchar(50), @partitioncount int, @allow_page_locks int'
        			+ ''
        			+ 'SELECT @objectname = o.name, @schemaname = s.name '
    				+ 'FROM [' + @db_name + '].sys.objects AS o '
  					+ 'JOIN sys.schema AS s ON s.schema_id = o.schema_id'
  					+ 'WHERE o.object_id = ' + CAST(@objectid as varchar(50))
  					+ ''
  					+ 'SELECT @indexname = name, @allow_page_locks = allow_page_locks '
  					+ 'FROM [' + @db_name + '].sys.indexes'
  					+ 'WHERE object_id = ' + CAST(@objectid as varchar(50)) + ' AND index_id = ' + CAST(@indexid as varchar(50))
  					+ 'SELECT @partitioncount = count(*) '
  					+ 'FROM [' + @db_name + '].sys.partitions '
  					+ 'WHERE object_id = ' + CAST(@objectid as varchar(50)) + ' AND index_id = ' + CAST(@indexid as varchar(50))
  					+ ''
  					+ 'INSERT INTO #med_info_index_detail '
  					+ '([objectid],[indexid],[partitionnum],[frag],[objectname],[schemaname],[indexname],[partitioncount],[db_name],[allow_page_locks],[avg_page_space_used_in_percent])'
  					+ 'VALUES ("' + CAST(@objectid as varchar(50)) + '", "' + CAST(@indexid as varchar(50)) + '", "' + CAST(@partitionnum as varchar(50)) + '", "' + CAST(@frag as varchar(50))
  					+ '", @objectname, @schemaname, @indexname, @partitioncount, "' + @db_name + '", @allow_page_locks, "' + CAST(@spaceUsed as varchar(50)) + '")'
  					+ ''
  
    EXEC(@command2)

    FETCH NEXT FROM dbList INTO @objectid, @indexid, @partitionnum, @frag, @db_name, @spaceUsed
  END
  	
  CLOSE dbList
  DEALLOCATE dbList
  
  -- STEP 3
  PRINT 'Step 3 starting... ' + CAST(GETDATE() AS nvarchar(50))
  PRINT '- Reorgonazing and rebuilding all index in the previous list'
  
  DECLARE DefragList CURSOR FOR
    SELECT objectid, indexid, partitionnum, frag, objectname, schemaname, indexname, partitioncount, db_name, allow_page_locks, avg_page_space_used_in_percent
    FROM #med_info_index_detail
    ORDER BY frag DESC
  FOR READ ONLY
  
  OPEN DefragList
  
  FETCH NEXT FROM DefragList
  INTO @objectid, @indexid, @partitionnum, @frag, @objectname, @schemaname, @indexname, @partitioncount, @db_name, @allow_page_locks, @spaceUsed
  
  WHILE @@FETCH_STATUS = 0
  BEGIN
    IF (@indexid) IS NOT NULL
    BEGIN
      -- 30 is an arbitrary decision point at which to switch between reorganizing and rebuilding
  
      IF @objectname IS NOT NULL AND @indexname IS NOT NULL AND @partitioncount IS NOT NULL
      BEGIN
        IF ((@frag >= 20.0 AND @frag < 30.0) OR (@spaceUsed < 75.0 AND @spaceUsed > 60.0)) AND @allow_page_locks = 1
        BEGIN
          SELECT @command = 'USE [' + @db_name + ']; ALTER INDEX [' + @indexname + '] ON ' + @schemaname + '.[' + @objectname + '] REORGANIZE'
  
	  IF @partitioncount > 1
	    SELECT @command = @command + ' PARTITION=' + CONVERT(CHAR, @partitionnum)

	    SELECT @commandStatistics = 'USE [' + @db_name + '; UPDATE STATISTICS ' + @schemaname + '.[' + @objectname + '](' + @indexname +') WITH FULLSCAN;'

	    IF @debugMode = 1 PRINT GETDATE() + '' + @commandStatistics
	    ELSE
  	    BEGIN
	      PRINT CAST(GETDATE() AS nvarchar(50)) + '' + @commandStatistics
	      EXEC(@commandStatistics)
	    END

	    IF @debugMode = 1 PRINT @command
  	    ELSE
	    BEGIN
	      PRINT CAST(GETDATE() AS nvarchar(50)) + '' + @command
	      EXEC(@command)
	    END
        END -- IF ((@frag >= 20.0 AND @frag < 30.0) OR (@spaceUsed < 75.0 AND @spaceUsed > 60.0)) AND @allow_page_locks = 1

        IF (@frag >= 30.0) OR (@spaceUsed < 60.0)
        BEGIN
          SELECT @command = 'USE [' + @db_name + ']; ALTER INDEX [' + @indexname + '] ON ' + @schemaname + '.[' + @objectname + '] REBUILD'

          IF @partitioncount = 1
            BEGIN
            SELECT @command = @command + ' PARTITION=' + CONVERT(CHAR, @partitionnum)

            -- If it's a partition we update statistics manually
            SELECT @commandStatistics = 'USE [' + @db_name + ']; UPDATE STATISTICS ' + @schemaname + '.[' + @objectname + '](' + @indexname + ') WITH FULLSCAN;'

            IF @debugMode = 1
            PRINT @commandStatistics
              ELSE
            BEGIN
              PRINT CAST(GETDATE() AS nvarchar(50)) + '' + @commandStatistics
              EXEC(@commandStatistics)
            END
          END

          IF @debugMode = 1 PRINT @command
          ELSE
          BEGIN
            PRINT CAST(GETDATE() AS nvarchar(50)) + '' + @command
            EXEC(@command)
          END
        END -- IF (@frag >= 30.0) OR (@spaceUsed < 60.0)
      END -- IF @objectname IS NOT NULL AND @indexname IS NOT NULL AND @patritioncount IS NOT NULL
    END -- IF (@indexid) IS NOT NULL

  FETCH NEXT FROM DefragList INTO @objectid, @indexid, @partitionnum, @frag, @objectname, @schemaname, @indexname, @partitioncount, @db_name, @allow_page_locks, @spaceUsed

  END -- WHILE @@FETCH_STATUS = 0

  CLOSE DefragList
  DEALLOCATE DefragList
  PRINT 'Ster 3 end'

  -- STEP 4
  PRINT 'Step 4 starting... ' + CAST(GETDATE() AS nvarchar(50))
  PRINT 'Remove temp table'

  DROP TABLE #med_info_index
  DROP TABLE #med_info_index_detail
  PRINT 'Step 4 end'

  SET @t2 = GETDATE()
  SELECT @time_diff = DATEDIFF(SECOND, @t1, @t2)

  PRINT 'Execute time is ' + CAST(@time_diff / 60 / 60 / 24 % 7 AS nvarchar(50)) + ' days ' + CAST(@time_diff / 60 / 60 % 24 AS nvarchar(50)) + ' hours ' + CAST(@time_diff / 60 % 60 AS nvarchar(50)) + ' minutes and ' + CAST(@time_diff % 60 AS nvarchar(50)) + ' seconds'

END -- handdleFragmentationIndexes

GO

EXECUTE handdleFragmentationIndexes 1, 'LIMITED', 'TEST'

GO

DBCC FREEPROCCACHE
DROP PROCEDURE handdleFragmentationIndexes
