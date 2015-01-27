DECLARE @BaseDirectory VARCHAR(MAX)

SET @BaseDirectory = 'F:\SP_File_Backup'
-- SET ROOT PATH for files


DECLARE
    @Stream VARBINARY(MAX),
    @Directory VARCHAR(MAX),
    @File VARCHAR(MAX),
    @ObjectToken INT,
    @Subdir AS VARCHAR(MAX),
    @FullDirectory AS VARCHAR(MAX),
    @FullName AS VARCHAR(MAX)	

-- Get all documents (Location, FileName, Stream)
DECLARE DocumentsCursor CURSOR FAST_FORWARD FOR 
        SELECT DirName, LeafName, Content FROM dbo.AllDocs AS doc INNER JOIN dbo.AllDocStreams AS stream ON doc.Id = stream.Id

OPEN DocumentsCursor 
FETCH NEXT FROM DocumentsCursor INTO @Directory, @File, @Stream 

WHILE @@FETCH_STATUS = 0
    BEGIN
			
		SET @FullDirectory = @BaseDirectory + '\' + REPLACE(@Directory, '/', '\')
				
		-- Create directory	
		DECLARE @Sql AS nVARCHAR(MAX)
		SET @Sql = 'EXEC master.dbo.xp_create_subdir ''' + @FullDirectory + ''''
		EXEC sp_Executesql @Sql
			
		-- Prepare filename
		SET @FullName =  @FullDirectory + '\' + @File

		-- Create file and write data
        EXEC sp_OACreate 'ADODB.Stream', @ObjectToken OUTPUT
        EXEC sp_OASetProperty @ObjectToken, 'Type', 1
        EXEC sp_OAMethod @ObjectToken, 'Open'
        EXEC sp_OAMethod @ObjectToken, 'Write', NULL, @Stream
        EXEC sp_OAMethod @ObjectToken, 'SaveToFile', NULL, @FullName, 2
        EXEC sp_OAMethod @ObjectToken, 'Close'
        EXEC sp_OADestroy @ObjectToken

        FETCH NEXT FROM DocumentsCursor INTO @Directory, @File, @Stream 
    END 

CLOSE DocumentsCursor
DEALLOCATE DocumentsCursor

