


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   PROCEDURE [dbo].[GetServePlaceSellDetail]

    @SystemDate NVARCHAR(50),
    @SystemDay NVARCHAR(50),
    @SystemTime NVARCHAR(50),
    @Date1 VARCHAR(10),
    @Date2 VARCHAR(10),
    @User1 INT,
    @User2 INT,
    @Station1 INT,
    @Station2 INT ,
    @Time1 NVARCHAR(5),
    @Time2 NVARCHAR(5),
    @Status1 INT ,
    @Branch1 INT ,
    @Branch2 INT 
AS 
    DECLARE @tmp1 INT
    DECLARE @tmp2 NVARCHAR(50)
    DECLARE @Time3 NVARCHAR(50)
    DECLARE @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @time2

    IF @User2 < @User1 
        BEGIN 
            SET @tmp1 = @User2
            SET @User2 = @User1
            SET @User1 = @tmp1	
        END	

    IF @Time2 < @Time1 
        BEGIN
		/*SET @tmp2 = @Time2
		SET @Time2 = @Time1
		SET @Time1 = @tmp2*/
            SET @Time3 = '00:00'
            SET @Time4 = '24:00'
        END
DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
	
    SELECT  @SystemDate AS SystemDate,
            @SystemDay AS SystemDay,
            @SystemDay + ' ' + @SystemDate + '    ' + @SystemTime AS Sysdate,
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @Station1 AS FromStation,
            @Station2 AS ToStationID,
            @User1 AS FromUser,
            @User2 AS ToUser,
            dbo.ViewServePlaceSellDetail.* , dbo.tBranch.nvcBranchName
            , @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2
    FROM    dbo.ViewServePlaceSellDetail
            INNER JOIN tServePlace ON tServePlace.intServePlace = ViewServePlaceSellDetail.intServePlace
            INNER JOIN dbo.tBranch ON ViewServePlaceSellDetail.Branch = tBranch.Branch
    WHERE   
             [date] >= @Date1
            AND [date] <= @Date2
            AND ViewServePlaceSellDetail.[User] >= @User1 
            AND ViewServePlaceSellDetail.[User] <= @User2
			AND ( ( ViewServePlaceSellDetail.[Time] >= @Time1
            AND ViewServePlaceSellDetail.[Time] <= @Time4
			)
			OR ( ViewServePlaceSellDetail.[Time] <= @Time2
               AND ViewServePlaceSellDetail.[Time] >= @Time3
             )
			)
            AND dbo.ViewServePlaceSellDetail.status = @status1
            AND ViewServePlaceSellDetail.StationId >= @Station1
            AND ViewServePlaceSellDetail.StationId <= @Station2
            AND ViewServePlaceSellDetail.Branch >= @Branch1
            AND ViewServePlaceSellDetail.Branch <= @Branch2
            
    ORDER BY dbo.ViewServePlaceSellDetail.[No]


GO

