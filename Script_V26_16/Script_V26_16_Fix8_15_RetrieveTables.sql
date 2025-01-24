

--Script V26_16_Fix8_RetriveTables
--93/09/10


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER    PROCEDURE [dbo].[RetriveTable_Branch]
    (
      @Branch INT ,
      @TableControl BIT
    )
AS 

DECLARE @CurrentDay NVARCHAR(8) 
SET @CurrentDay = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())    

    IF @TableControl = 0 
        SELECT  [No] ,
                [Name] AS TableDescription , 
                CASE WHEN [No] IN (SELECT No FROM dbo.[FN_NoEmptyTables](@CurrentDay)) THEN 0 ELSE 1 END AS Empty , -- Empty ,
                 NumberOfChair , dbo.tPartitions.*
        FROM    tTable  INNER JOIN tPartitions ON tPartitions.PartitionID = tTable.PartitionID AND  tPartitions.Branch = tTable.Branch
	  WHERE tPartitions.Branch = @Branch
        ORDER BY dbo.tTable.[No] , dbo.tPartitions.PartitionID

    IF @TableControl = 1 
        SELECT  [No] ,
                [Name] AS TableDescription , 1 AS  Empty ,NumberOfChair, dbo.tPartitions.*
        FROM    tTable INNER JOIN tPartitions ON tPartitions.PartitionID = tTable.PartitionID AND  tPartitions.Branch = tTable.Branch
	   WHERE [No] NOT  IN (SELECT No FROM dbo.FN_NoEmptyTables(@CurrentDay))  --  tTable.Empty = 1
	        AND  tPartitions.Branch = @Branch
        ORDER BY dbo.tTable.[No] , dbo.tPartitions.PartitionID
--===============================================


GO


