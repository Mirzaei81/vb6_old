


IF COL_LENGTH('tblTotal_PrintFich','nvcPrintDate') IS NULL
BEGIN
	ALTER TABLE dbo.tblTotal_PrintFich
	ADD nvcPrintDate DATETIME NULL 
END

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Update_tblTotal_printFich
	(
	 @intPrintFichNo INT ,
	 @TypeFlag BIT ,
	 @nvcError NVARCHAR(255)
	 )

AS


IF @TypeFlag = 0
	UPDATE [dbo].[tblTotal_PrintFich]  
		SET IsPrinted=1 , nvcError = @nvcError , nvcPrintDate = GETDATE()
	 	WHERE intPrintFichNo=@intPrintFichNo 

ELSE
	DELETE FROM  [dbo].[tblTotal_PrintFich]  
	 	WHERE intPrintFichNo=@intPrintFichNo 
	

Return 1



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

---------------------------------------------------------
-----------------گزارش فروش درصدی  و ساعتی


ALTER   PROCEDURE [dbo].[GetPercentInvoicePerHourInfo]
    (
      @intLanguage INT = 0,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50),
      @Branch1 INT,
      @Branch2 INT
    )
AS 
    DECLARE @tmp NVARCHAR(50)
    declare @Time3 NVARCHAR(50)
    declare @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @Time2
	

    IF @Time2 < @Time1 
        BEGIN
		/*SET @tmp        = @Time2
		SET @Time2  = @Time1
		SET @Time1 = @tmp*/
            SET @Time3 = '00'
            SET @Time4 = '24'
        END

    IF @Date2 < @Date1 
        BEGIN
            SET @tmp = @Date2
            SET @Date2 = @Date1
            SET @Date1 = @tmp
		
        END

    SET @Time1 = LTRIM(LEFT(@Time1, 2))
    SET @Time2 = LTRIM(LEFT(@Time2, 2))
    SET @Time3 = LTRIM(LEFT(@Time3, 2))
    SET @Time4 = LTRIM(LEFT(@Time4, 2))

    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage = 0 
        SET @TimeTitle = N' ساعت : '
    ELSE 
        SET @TimeTitle = N'Time: '

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
    SELECT  dbo.VwPercentInvoicePerHour.FactorCount,
            dbo.VwPercentInvoicePerHour.SalePriceTotal,
            dbo.VwPercentInvoicePerHour.[Date],
            dbo.VwPercentInvoicePerHour.FromTime,
            dbo.VwPercentInvoicePerHour.ToTime,
            dbo.VwPercentInvoicePerHour.Branch,
            CAST(( dbo.VwPercentInvoicePerHour.SalePriceTotal
                   / t.MySalePriceTotal ) * 100 AS DECIMAL(6, 3)) AS TotalPercent,
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @Time1 AS TimeBefore,
            @Time2 AS TimeAfter,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate
            , @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2 , Branch
    FROM    dbo.VwPercentInvoicePerHour
            INNER JOIN ( SELECT dbo.VwPercentInvoicePerHour.[Date] AS myDate,
                                SUM(dbo.VwPercentInvoicePerHour.SalePriceTotal) AS MySalePriceTotal
                         FROM   dbo.VwPercentInvoicePerHour
                         WHERE  dbo.VwPercentInvoicePerHour.[date] >= @Date1
                                AND dbo.VwPercentInvoicePerHour.[date] <= @Date2
                                AND ( ( dbo.VwPercentInvoicePerHour.FromTime >= @Time1
                                        AND dbo.VwPercentInvoicePerHour.ToTime <= @Time4
                                      )
                                      OR ( dbo.VwPercentInvoicePerHour.ToTime <= @Time2
                                           AND dbo.VwPercentInvoicePerHour.FromTime >= @Time3
                                         )
                                    )
                         GROUP BY dbo.VwPercentInvoicePerHour.[date]
                       ) t ON t.MyDate = dbo.VwPercentInvoicePerHour.[date]
    WHERE   dbo.VwPercentInvoicePerHour.[date] >= @Date1
            AND dbo.VwPercentInvoicePerHour.[date] <= @Date2
            AND ( ( dbo.VwPercentInvoicePerHour.FromTime >= @Time1
                    AND dbo.VwPercentInvoicePerHour.ToTime <= @Time4
                  )
                  OR ( dbo.VwPercentInvoicePerHour.ToTime <= @Time2
                       AND dbo.VwPercentInvoicePerHour.FromTime >= @Time3
                     )
                )
	--AND dbo.VwPercentInvoicePerHour.FromTime >= @Time1 
	--AND dbo.VwPercentInvoicePerHour.ToTime   <=  @Time2
            AND dbo.VwPercentInvoicePerHour.Branch >= @Branch1
            AND dbo.VwPercentInvoicePerHour.Branch <= @Branch2
    ORDER BY dbo.VwPercentInvoicePerHour.[Date] , dbo.VwPercentInvoicePerHour.FromTime


GO




------------------------------------------------------------------------------
--------------------------------------------------------------------------


