
--Script_V26_16_Fix8_4TableBusyTime.SQL
--نمایش تایم دقیق زمان میز پر
--حتی اگر به روزهای بعد کشیده شود
--93/08/04


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  VIEW [dbo].[vwSamar_TableUsage_BusyTable]
AS  SELECT  intTableUsedNo,
            intBranch,
            intReseveDetailNo,
            nvcAssignTime,
            nvcStartTime,
            nvcEndTime,
            intTableNo,
            bitIsValid,
			( CAST(SUBSTRING(dbo.shamsi(GETDATE()), 4, 2) AS INT) - 1 * 30
			+ CAST(SUBSTRING(dbo.shamsi(GETDATE()), 7, 2) AS INT) ) * 1440
			+ ( DATEPART(HOUR, GETDATE()) * 60 + DATEPART(minute, GETDATE()) ) 
			- ( CAST(SUBSTRING(nvcUsedDate, 4, 2) AS INT) - 1 * 30
			+ CAST(SUBSTRING(nvcUsedDate, 7, 2) AS INT) )  * 1440
			- ( CAST(SUBSTRING(nvcStartTime, 1, 2) AS INT) * 60
			+ CAST(SUBSTRING(nvcStartTime, 4, 2) AS INT) ) AS MinuteUseDiff,
            nvcUsedDate
    FROM    dbo.tblSamar_TableUsage
    WHERE   ( nvcEndTime IS NULL )
            AND ( bitIsValid = 1 )

--=========================================================================
GO




ALTER   PROCEDURE [Get_tblSamar_TableUsage_BusyTable]

@Branch INT  

AS

SELECT  DISTINCT   t.intTableNo AS NO , t.MinuteUseDiff , MAX(intTableUsedNo) , dbo.tTable.PartitionID , dbo.tTable.nvcMaxUseTime , dbo.tTable.Name AS TableDescription
    FROM    vwSamar_TableUsage_BusyTable t
    INNER JOIN dbo.tTable ON t.intBranch = dbo.tTable.Branch AND t.intTableNo = dbo.tTable.No
    WHERE  t.intBranch=@Branch --AND t.nvcUsedDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())
    GROUP BY intTableNo , MinuteUseDiff , PartitionID , nvcMaxUseTime ,  dbo.tTable.Name
    ORDER BY PartitionID , No




GO


