USE master
GO


IF OBJECT_ID('mtc_printExecutionTime', 'P') IS NOT NULL
BEGIN
  DROP PROCEDURE mtc_printExecutionTime
  PRINT 'mtc_printExecutionTime already exist and was DROPED'
END
ELSE
  PRINT 'mtc_printExecutionTime is not exists and will be created'
GO

CREATE PROCEDURE mtc_printExecutionTime(@t1 DATETIME)
AS
BEGIN
  SET NOCOUNT ON

  DECLARE @t2 DATETIME = GETDATE()
  DECLARE @time_diff int

  SELECT @time_diff = DATEDIFF(SECOND, @t1, @t2)

  PRINT 'Execute time is ' + CAST(@time_diff / 60 / 60 / 24 % 7 AS nvarchar(50)) + ' days ' + CAST(@time_diff / 60 / 60 % 24 AS nvarchar(50)) + ' hours ' + CAST(@time_diff / 60 % 60 AS nvarchar(50)) + ' minutes and ' + CAST(@time_diff % 60 AS nvarchar(50)) + ' seconds'
END


IF OBJECT_ID('mtc_createMaintenanceDatabaseIfNeed', 'P') IS NOT NULL
BEGIN
  DROP  PROCEDURE mtc_createMaintenanceDatabaseIfNeed
  PRINT 'mtc_createMaintenanceDatabaseIfNeed already exist and was DROPED'
END
ELSE
  PRINT 'mtc_createMaintenanceDatabaseIfNeed is not exist and will be CREATED'
GO

CREATE PROCEDURE mtc_createMaintenanceDatabaseIfNeed
AS
BEGIN
  IF DB_ID('mtc_Maintenance') IS NULL
  BEGIN
    PRINT 'Creating mtc_Maintenance batabase'
    CREATE DATABASE mtc_Maintenance
  END
  ELSE
  BEGIN
    PRINT 'Found mtc_Maintenance database'
  END

  IF OBJECT_ID('mtc_Maintenance.dbo.med_info_index') IS NULL
  BEGIN
    PRINT 'Creating table mtc_Maintenance.dbo.med_info_index'
    CREATE TABLE mtc_Maintenance.dbo.med_info_index(object_id int, index_id int, partition_number int, avg_fragmentation_in_percent float, [db_name] varchar(150), avg_page_space_used_in_percent float)
  END
  ELSE
  BEGIN
    PRINT 'Found table mtc_Maintenance.dbo.med_info_index'
  END

  IF OBJECT_ID('mtc_Maintenance.dbo.med_info_index_detail') IS NULL
  BEGIN
    PRINT 'Creating table mtc_Maintenance.dbo.med_info_index_detail'
    CREATE TABLE mtc_Maintenance.dbo.med_info_index_detail([objectid] int, [indexid] int, [partitionnum] int, [frag] float, [objectname] varchar(150), [schemaname] varchar(150), [db_name] varchar(150), [indexname] varchar(150), [partitioncount] int, [allow_page_locks] int, [avg_page_space_used_in_percent] float)
  END
  ELSE
  BEGIN
    PRINT 'Found table mtc_Maintenance.dbo.med_info_index'
  END
END


IF OBJECT_ID('mtc_getAllIndexesWithAFragmentationOverThan20Percent', 'P') IS NOT NULL
BEGIN
  DROP PROCEDURE mtc_getAllIndexesWithAFragmentationOverThan20Percent
  PRINT 'mtc_getAllIndexesWithAFragmentationOverThan20Percent already exist and was DROPED'
END
ELSE
  PRINT 'mtc_getAllIndexesWithAFragmentationOverThan20Percent is not exist and will be CREATED'
GO

CREATE PROCEDURE mtc_getAllIndexesWithAFragmentationOverThan20Percent(@databaseName varchar(50) = 'ALL', @mode varchar(50) = 'DETAILED')
AS
BEGIN

  -- @mode =
  -- LIMITED (default): This mode is the fastest and scans the smallest number of pages. It scans all pages for a heap, but only scans the parent-level pages, which means, the pages above the leaf level, for an index
  -- SAMPLED: This mode returns statistics base on a one percent sample of the entire page in the index or heap. If the index or heap has fewer tha 10 000 pages, DETAILED mode is used instead of SAMPLED.
  -- DETAILED: Detailed mode scans all pages and returns all statistics. Be careful, from LIMITED to SAMPLED to DETAILED, the mode are progressively slower, because more work is performed in each. In my script I am useng this one.

  SET NOCOUNT ON

  DECLARE @database_id int
  DECLARE @objectid int
  DECLARE @indexid int
  DECLARE @partitionnum bigint
  DECLARE @db_name varchar(50)
  DECLARE @spaceUsed float
  DECLARE @command2 varchar(MAX)
  DECLARE @frag float
  DECLARE @t1 DATETIME = GETDATE()
  
  PRINT 'Starting... ' + CAST(GETDATE() AS nvarchar(max))
  PRINT '- Get all indexes with a fragmentation over than 20% for database: ' + @databaseName

  IF @databaseName = 'ALL'
    SET @database_id = NULL
  ELSE
    SELECT @database_id = ISNULL(DB_ID(@databaseName), 0)

  INSERT mtc_Maintenance.dbo.med_info_index(object_id, index_id, partition_number, avg_fragmentation_in_percent, [db_name], [avg_page_space_used_in_percent])
  	SELECT object_id, index_id, partition_number, avg_fragmentation_in_percent, db_name(database_id) as [db_name], ISNULL(avg_page_space_used_in_percent,100)
  	FROM sys.dm_db_index_physical_stats(@database_id, NULL, NULL, NULL, @mode)
  	WHERE avg_fragmentation_in_percent >= 20
  		AND database_id IN (SELECT database_id
  		    		    FROM sys.databases
  				    WHERE [state] = 0 /* ONLINE */ AND is_read_only = 0 AND database_id > 4 /* SKIP SYSTEM DB */
  				    -- Documentation : http://msdn.microsoft.com/en-us/library/ms178534.aspx
  				    )  -- STEP 1
  PRINT 'Step 1 starting... ' + CAST(GETDATE() AS nvarchar(max))
  PRINT '- Get all indexes with a fragmentation over than 20% for database: ' + @databaseName

  IF @databaseName = 'ALL'
    SET @database_id = NULL
  ELSE
    SELECT @database_id = ISNULL(DB_ID(@databaseName), 0)

  INSERT mtc_Maintenance.dbo.med_info_index(object_id, index_id, partition_number, avg_fragmentation_in_percent, [db_name], [avg_page_space_used_in_percent])
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
  	FROM mtc_Maintenance.dbo.med_info_index
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
  					+ 'INSERT INTO med_info_index_detail '
  					+ '([objectid],[indexid],[partitionnum],[frag],[objectname],[schemaname],[indexname],[partitioncount],[db_name],[allow_page_locks],[avg_page_space_used_in_percent])'
  					+ 'VALUES ("' + CAST(@objectid as varchar(50)) + '", "' + CAST(@indexid as varchar(50)) + '", "' + CAST(@partitionnum as varchar(50)) + '", "' + CAST(@frag as varchar(50))
  					+ '", @objectname, @schemaname, @indexname, @partitioncount, "' + @db_name + '", @allow_page_locks, "' + CAST(@spaceUsed as varchar(50)) + '")'
  					+ ''
  
    EXEC(@command2)

    FETCH NEXT FROM dbList INTO @objectid, @indexid, @partitionnum, @frag, @db_name, @spaceUsed
  END
  	
  CLOSE dbList
  DEALLOCATE dbList
  
  EXECUTE mtc_printExecutionTime @t1
END


IF OBJECT_ID('mtc_reorgonizeAndRebuildIndexes', 'P') IS NOT NULL
BEGIN
  DROP PROCEDURE mtc_reorgonizeAndRebuildIndexes
  PRINT 'mtc_reorgonizeAndRebuildIndexes already exist and was DROPED'
END
ELSE
  PRINT 'mtc_reorgonizeAndRebuildIndexes is not exists and will be created'
GO

CREATE PROCEDURE mtc_reorgonizeAndRebuildIndexes(@debugMode bit = 0, @mode varchar(50) = 'DETAILED', @timeLimit DATETIME = NULL)
AS
BEGIN
  
  -- @mode =
  -- LIMITED (default): This mode is the fastest and scans the smallest number of pages. It scans all pages for a heap, but only scans the parent-level pages, which means, the pages above the leaf level, for an index
  -- SAMPLED: This mode returns statistics base on a one percent sample of the entire page in the index or heap. If the index or heap has fewer tha 10 000 pages, DETAILED mode is used instead of SAMPLED.
  -- DETAILED: Detailed mode scans all pages and returns all statistics. Be careful, from LIMITED to SAMPLED to DETAILED, the mode are progressively slower, because more work is performed in each. In my script I am useng this one.

  -- TODO: add MaxDOP option for defragmentation and rebuild

  SET NOCOUNT ON

  DECLARE @objectid int
  DECLARE @indexid int
  DECLARE @partitionnum bigint
  DECLARE @frag float
  DECLARE @objectname sysname
  DECLARE @schemaname sysname
  DECLARE @indexname sysname
  DECLARE @partitioncount bigint
  DECLARE @db_name varchar(50)
  DECLARE @allow_page_locks int
  DECLARE @spaceUsed float
  DECLARE @commandStatistics varchar(8000)
  DECLARE @command varchar(MAX)
  DECLARE @t1 DATETIME = GETDATE()
  DECLARE @time_diff int

  PRINT 'Starting... ' + CAST(GETDATE() AS nvarchar(50))
  PRINT '- Reorgonazing and rebuilding all index in the previous list'
  
  DECLARE DefragList CURSOR FOR
    SELECT objectid, indexid, partitionnum, frag, objectname, schemaname, indexname, partitioncount, db_name, allow_page_locks, avg_page_space_used_in_percent
    FROM mtc_Maintenance.dbo.med_info_index_detail
    ORDER BY frag DESC
  FOR READ ONLY
  
  OPEN DefragList
  
  FETCH NEXT FROM DefragList
  INTO @objectid, @indexid, @partitionnum, @frag, @objectname, @schemaname, @indexname, @partitioncount, @db_name, @allow_page_locks, @spaceUsed
  
  WHILE @@FETCH_STATUS = 0
  BEGIN

    IF @timeLimit IS NOT NULL
    BEGIN
      SELECT @time_diff = DATEDIFF(SECOND, @t1, GETDATE())
      IF @time_diff >= @timeLimit
        BREAK
    END

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
  
  EXEC mtc_printExecutionTime @t1
END
