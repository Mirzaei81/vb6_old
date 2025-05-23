

--اضافه کردن ترتیب در گزارش فروش صندوق
--93/12/05


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[GetCashInvoice]
    (
      --      @intLanguage INT = 0 ,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @User1 INT,
      @User2 INT,
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50),
      @Station1 INT,
      @Station2 INT ,   
    @Branch1 INT ,
    @Branch2 INT 
    )
AS 
    DECLARE @intLanguage INT
    SET @intLanguage = 0  
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
    DECLARE @TimeTitle NVARCHAR(10)  
    IF @intLanguage = 0 
        SET @TimeTitle = N' ساعت : '  
    ELSE 
        SET @TimeTitle = N'Time: '  

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2

    SELECT  dbo.VwCashInvoice.UID,
            dbo.VwCashInvoice.fullname,
            dbo.VwCashInvoice.pPno,
            dbo.VwCashInvoice.UserFullName,
            dbo.VwCashInvoice.SumPrice,
			ISNULL((SELECT SUM(Bestankar) FROM dbo.tblAcc_Recieved 
			 WHERE Date>=dbo.VwCashInvoice.[Date]
					AND [date] <= dbo.VwCashInvoice.[Date] 
					AND AddUser>=dbo.VwCashInvoice.[User]
					AND AddUser<=dbo.VwCashInvoice.[User]),0)AS SumReceives,
            dbo.VwCashInvoice.[No],
            dbo.VwCashInvoice.[Date],
            dbo.VwCashInvoice.[User],
            dbo.VwCashInvoice.Shift,
            dbo.VwCashInvoice.StationId,
            dbo.VwCashInvoice.[Time],
            dbo.VwCashInvoice.[intSerialNo],
            CASE @intLanguage
              WHEN 0 THEN dbo.VwCashInvoice.Gender
              WHEN 1 THEN dbo.VwCashInvoice.GenderLatin
            END AS Gender,
            dbo.VwCashInvoice.Balance,
            CASE @intLanguage
              WHEN 0 THEN dbo.VwCashInvoice.BalanceDescription
              WHEN 1 THEN dbo.VwCashInvoice.BalanceDescriptionLatin
            END AS BalanceDescription,
            dbo.VwCashInvoice.FullName,
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate,
            VwCashInvoice.Tip,
            VwCashInvoice.[TaxTotal],[DutyTotal],[ServiceTotal],[CarryFeeTotal],[PackingTotal]
    		, TableName , @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2 , Branch , nvcBranchName
    FROM    dbo.VwCashInvoice
    WHERE   [date ] >= @Date1
            AND [date] <= @Date2
            AND dbo.VwCashInvoice.uid >= @User1
            AND uid < = @User2
            AND ( ( [Time] >= @Time1
                    AND [Time] <= @Time4
                  )
                  OR ( [Time] <= @Time2
                       AND [Time] >= @Time3
                     )
                )
            AND dbo.VwCashInvoice.StationID >= @Station1
            AND dbo.VwCashInvoice.StationID <= @Station2
            AND dbo.VwCashInvoice.Branch >= @Branch1
            AND dbo.VwCashInvoice.Branch <= @Branch2
    ORDER BY dbo.VwCashInvoice.[Date] , dbo.VwCashInvoice.No 
 

GO

