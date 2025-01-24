

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Total_ReIndex]
AS 
    DECLARE @DataFile NVARCHAR(50)
    DECLARE @LogFile NVARCHAR(50)

    SELECT  @DataFile = [name] /*fileid ,, filename, size, growth, status, maxsize */
    FROM    dbo.sysfiles
    WHERE   fileid = 1 --(status & 0x40) <> 0 AND
    SELECT  @LogFile = [name] /*fileid ,, filename, size, growth, status, maxsize*/
    FROM    dbo.sysfiles
    WHERE   fileid = 2 --(status & 0x40) <> 0 AND

--PRINT @DataFile
--PRINT @LogFile
    DECLARE @Str NVARCHAR(200)
--(N'Total_Data')
    SET @Str = N'DBCC SHRINKFILE (N''' + RTRIM(LTRIM(@DataFile)) + '''' + ')'
--PRINT @Str
    EXECUTE sp_executesql @str


    SET @Str = N'DBCC SHRINKFILE (N''' + RTRIM(LTRIM(@LogFile)) + '''' + ')'
    EXECUTE sp_executesql @str
--PRINT @Str

    DBCC DBREINDEX ('tfacd', '' , 0)
    DBCC DBREINDEX ('tfacm', '' , 0)
    DBCC DBREINDEX ('tfacd2', '' , 0)
    DBCC DBREINDEX ('tCASh', '' , 0)
    DBCC DBREINDEX ('tCust', '' , 0)
    DBCC DBREINDEX ('tFacCard', '' , 0)
    DBCC DBREINDEX ('tFacCash', '' , 0)
    DBCC DBREINDEX ('tFacCheque', '' , 0)
    DBCC DBREINDEX ('tFacCredit', '' , 0)
    DBCC DBREINDEX ('tFacLoan', '' , 0)
    DBCC DBREINDEX ('tGood', '' , 0)
    DBCC DBREINDEX ('tGoodLevel1', '' , 0)
    DBCC DBREINDEX ('tGoodLevel2', '' , 0)
    DBCC DBREINDEX ('tInventory', '' , 0)
    DBCC DBREINDEX ('tHistory', '' , 0)
    DBCC DBREINDEX ('tInventory_Good', '' , 0)
    DBCC DBREINDEX ('tPer', '' , 0)
    DBCC DBREINDEX ('tRepFacEditM', '' , 0)
    DBCC DBREINDEX ('tStation_Inventory_Good', '' , 0)
    DBCC DBREINDEX ('tStations', '' , 0)
    DBCC DBREINDEX ('tSupplier', '' , 0)
    DBCC DBREINDEX ('tUser', '' , 0)
    DBCC DBREINDEX ('tHavaleM', '' , 0)
    DBCC DBREINDEX ('tHavaleD', '' , 0)

--    SET @Str = N'DBCC SHRINKFILE (N''' + RTRIM(LTRIM(@DataFile)) + '''' + ')'
----PRINT @Str
--    EXECUTE sp_executesql @str
--
--
--    SET @Str = N'DBCC SHRINKFILE (N''' + RTRIM(LTRIM(@LogFile)) + '''' + ')'
--    EXECUTE sp_executesql @str
--===============================================
DECLARE @nvcDate NVARCHAR(8)
SET @nvcDate =  dbo.Get_ShamsiDate_For_Current_Shift(getdate())
--PRINT @nvcDate

DELETE FROM dbo.tblSamar_TableUsage WHERE nvcUsedDate <> @nvcDate


GO
