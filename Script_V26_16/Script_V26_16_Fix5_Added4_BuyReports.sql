
--Script_V26_16_Fix5_Added4_BuyReports
--گزارشات خرید و حواله
--93/04/30

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

--Script Speed Of Crystal Report 
--افزایش سرعت ریپورت گروهی کالا
-- روی همه ورژن های رستورانی

----------------------------------------------------------------------------------------------------------
ALTER VIEW dbo.ViewRepSellKind
AS
SELECT     TOP 100 PERCENT 
	dbo.tFacD.Amount,dbo.tFacD.GoodCode, dbo.tFacD.FeeUnit, dbo.tGood.NamePrn,
        dbo.tGoodLevel1.Code AS Level1Code, dbo.tGoodLevel1.Description AS Level1Desc, 
        dbo.tGoodLevel2.Code AS Level2Code, dbo.tGoodLevel2.Description AS Level2Desc,
        dbo.tFacM.Status, 
	    dbo.tFacM.DiscountTotal, dbo.tFacM.CarryFeeTotal, 
        dbo.tFacM.SumPrice, 
        dbo.tFacM.StationID, dbo.tFacM.ServiceTotal, dbo.tFacM.PackingTotal, 
        dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.Branch, 
        dbo.tFacD.Discount, dbo.tFacM.AccountYear  ,
	    tBranch.nvcBranchName,tUnitGood.Description as UnitGoodDescription
		,tfacm.TaxTotal , tfacm.DutyTotal , dbo.tFacM.Owner AS Supplier , tGood.BarCode
		, dbo.tfacd.Discount as Discountrate
FROM         dbo.tFacM INNER JOIN
			dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND dbo.tFacM.Branch = dbo.tFacD.Branch INNER JOIN
			dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code INNER JOIN
			dbo.tBranch ON dbo.tfacm.Branch = dbo.tBranch.Branch INNER JOIN
			dbo.tGoodLevel1 ON dbo.tGood.Level1 = dbo.tGoodLevel1.Code INNER JOIN
			dbo.tUnitGood on dbo.tgood.unit=dbo.tUnitGood.code inner join	
			dbo.tGoodLevel2 ON dbo.tGood.Level2 = dbo.tGoodLevel2.Code  
		     
WHERE     (dbo.tFacM.Recursive = 0)
ORDER BY dbo.tFacm.intSerialNo



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[GetBuyKindInfo]
    (
      --@intLanguage INT = 0 ,
      @SystemDate NVARCHAR(50) ,
      @SystemDay NVARCHAR(50) ,
      @SystemTime NVARCHAR(50) ,
      @Date1 NVARCHAR(50) ,
      @Date2 NVARCHAR(50) ,
      @Sup1 INT ,
      @Sup2 INT ,
      @station1 INT ,
      @station2 INT ,
      @Time1 NVARCHAR(50) ,
      @Time2 NVARCHAR(50) ,
      @level11 INT ,
      @level12 INT ,
      @level21 INT ,
      @level22 INT ,
      @Branch1 INT ,
      @Branch2 INT ,     
      @Status1 INT 

    )
AS 

    DECLARE @tmp1 INT
    DECLARE @tmp2 NVARCHAR(50)
    DECLARE @DiscountTotal INT
    DECLARE @CarryFeeTotal INT
    DECLARE @PackingTotal INT
    DECLARE @TaxTotal INT
    DECLARE @DutyTotal INT
    DECLARE @Time3 NVARCHAR(50)
    DECLARE @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @tmp2

    IF @Sup2 < @Sup1 
        BEGIN 
            SET @tmp1 = @Sup2
            SET @Sup2 = @Sup1
            SET @Sup1 = @tmp1	
        END	

    IF @Time2 < @Time1 
        BEGIN
		/*SET @tmp2 = @Time2
		SET @Time2 = @Time1
		SET @Time1 = @tmp2*/
            SET @Time3 = '00:00'
            SET @Time4 = '24:00'
        END
	
    SELECT @DiscountTotal = SUM(DiscountTotal) ,
           @CarryFeeTotal = SUM(CarryFeeTotal) ,
           @PackingTotal  = SUM(PackingTotal) ,
           @TaxTotal      = SUM(TaxTotal) ,
           @DutyTotal     = SUM(DutyTotal)
                           FROM     tfacm
                           WHERE    [date] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = @status1
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2

    DECLARE @TimeTitle NVARCHAR(10)
        SET @TimeTitle = N' ساعت : '

    SELECT  SUM(dbo.ViewRepSellKind.Amount) AS SumAmount ,
            dbo.ViewRepSellKind.FeeUnit ,
            dbo.ViewRepSellKind.FeeUnit * SUM(dbo.ViewRepSellKind.Amount) AS PriceTotal ,
            dbo.ViewRepSellKind.Level1Code ,
            dbo.ViewRepSellKind.Level2Code ,
            dbo.ViewRepSellKind.UnitGoodDescription ,
            @Date1 AS DateBefore ,
            @Date2 AS DateAfter ,
            dbo.ViewRepSellKind.GoodCode ,
            @Sup1 AS FromSupplier ,
            @Sup2 AS ToSupplier ,
            @Time1 AS FromTime ,
            @Time2 AS ToTime ,
            @SystemDay + ' ' + @SystemDate + @TimeTitle + @SystemTime AS Sysdate ,
            @level11 AS FromGoodCodeLvele1 ,
            @level21 AS FromGoodCodeLvele2 ,
            @level12 AS ToGoodCodeLevel1 ,
            @level22 AS ToGoodCodeLevel2 ,
            dbo.ViewRepSellKind.NamePrn AS GoodName ,
			dbo.ViewRepSellKind.Level1Desc AS Level1Description,
            dbo.ViewRepSellKind.Level2Desc AS Level2Description ,
            dbo.ViewRepSellKind.DiscountRate ,
            (dbo.ViewRepSellKind.Discount * dbo.ViewRepSellKind.FeeUnit * SUM(dbo.ViewRepSellKind.Amount)) AS SumDiscount ,
            dbo.ViewRepSellKind.nvcBranchName ,
            @DiscountTotal AS DiscountTotal ,
            @CarryFeeTotal AS CarryFeeTotal ,
            @PackingTotal AS PackingTotal ,
            @Status1 AS Status ,
            ViewRepSellKind.AccountYear ,
            ViewRepSellKind.Branch 
--            ,dbo.ViewRepSellKind.ServePlace
	    , @taxTotal AS TaxTotal , @DutyTotal AS DutyTotal , ViewRepSellKind.Barcode
    FROM    dbo.ViewRepSellKind
    WHERE   dbo.ViewRepSellKind.[date] >= @Date1
            AND dbo.ViewRepSellKind.[date] <= @Date2
	--AND dbo.ViewRepSellKind.[Supplier] IN (SELECT code FROM tSupplier WHERE code  >= @Sup1 AND code < =@Sup2)
            AND dbo.ViewRepSellKind.Supplier >= @Sup1
                  AND dbo.ViewRepSellKind.Supplier <= @Sup2

            AND ( ( dbo.ViewRepSellKind.[Time] >= @Time1
                    AND dbo.ViewRepSellKind.[Time] <= @Time4
                  )
                  OR ( dbo.ViewRepSellKind.[Time] <= @Time2
                       AND dbo.ViewRepSellKind.[Time] >= @Time3
                     )
                )
            AND dbo.ViewRepSellKind.Status = @Status1
            AND dbo.ViewRepSellKind.StationID >= @station1
            AND dbo.ViewRepSellKind.StationID <= @station2
            AND dbo.ViewRepSellKind.Level1Code >= @level11
            AND dbo.ViewRepSellKind.Level1Code <= @level12
            AND dbo.ViewRepSellKind.Level2Code >= @level21
            AND dbo.ViewRepSellKind.Level2Code <= @level22
            AND dbo.ViewRepSellKind.Branch >= @Branch1
            AND dbo.ViewRepSellKind.Branch <= @Branch2
		--	AND dbo.ViewRepSellKind.Balance = 1
    GROUP BY --dbo.ViewRepSellKind.ServePlace ,
            dbo.ViewRepSellKind.GoodCode ,
            dbo.ViewRepSellKind.UnitGoodDescription ,
            dbo.ViewRepSellKind.FeeUnit ,
            dbo.ViewRepSellKind.Level2Desc ,
            dbo.ViewRepSellKind.Level1Desc ,
            dbo.ViewRepSellKind.NamePrn ,
            dbo.ViewRepSellKind.Level1Code ,
            dbo.ViewRepSellKind.Level2Code ,
            dbo.ViewRepSellKind.Discount ,
            ViewRepSellKind.AccountYear ,
            ViewRepSellKind.Branch ,
            ViewRepSellKind.nvcBranchName ,
             ViewRepSellKind.Barcode ,
            dbo.ViewRepSellKind.Discount ,
            dbo.ViewRepSellKind.DiscountRate 
    ORDER BY dbo.ViewRepSellKind.GoodCode



GO


INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  32 ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
          FROM tblTotal_ItemReports_Details WHERE intReportId = 1 AND (Row = 3 OR Row = 4)
          
GO






if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_AllAssignement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_AllAssignement
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

-----------------------------------گزارش حوالهاي صادره------------------


CREATE  PROC Get_AllAssignement(		
	@SystemDate NVARCHAR(20) ,
	@SystemDay NVARCHAR(20) ,
	@SystemTime NVARCHAR(20) ,
	@Date1 NVARCHAR(8),
	@Date2 NVARCHAR(8),
	@Inventory1 INT ,
	@Inventory2 INT)
AS
BEGIN


	SELECT DISTINCT dbo.tFacM.No,
			dbo.tFacM.Date,
			dbo.tFacM.SumPrice,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE  dbo.tFacD.intInventoryNo = dbo.tInventory.InventoryNo ) AS DepartureInventory,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE dbo.tFacD.DestInventoryNo = dbo.tInventory.InventoryNo ) AS DestinationInventory,
			(SELECT nvcFirstName+' '+nvcSurName FROM tper JOIN dbo.tUser ON dbo.tPer.pPno = dbo.tUser.pPno WHERE dbo.tUser.UID=dbo.tFacM.[User]) AS [User],
			dbo.tFacM.NvcDescription
			,@SystemDate AS SystemDate,
			@SystemDay AS SystemDay,
			@SystemTime AS SystemTime
	 FROM dbo.tFacM
		INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
	 WHERE [Status] =6
			AND dbo.tFacD.DestInventoryNo >=@Inventory1
			AND dbo.tFacD.DestInventoryNo <=@Inventory2
			AND dbo.tFacm.Date<=@Date2
			AND dbo.tFacM.Date>=@Date1
	ORDER BY dbo.tFacM.No
END

GO



--exec dbo.Get_AllAssignement N'93/04/30', N'دو شنبه', N'01:57', N'93/04/30', N'93/04/30', 1, 100
--GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_AssignmentDetail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_AssignmentDetail
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


-------------------------------گزارش حواله هاي صادره جزييات-------------
CREATE  PROC	Get_AssignmentDetail(   
	@SystemDate NVARCHAR(20) ,
	@SystemDay NVARCHAR(20) ,
	@SystemTime NVARCHAR(20) ,
	@Date1 NVARCHAR(8),
	@Date2 NVARCHAR(8),
	@Inventory1 INT,
	@Inventory2 INT)
AS	
BEGIN

	SELECT  dbo.tFacM.No,
			dbo.tGood.Name AS GoodName,
			SUM(dbo.tFacD.Amount) AS Amount,
			dbo.tFacD.FeeUnit ,
			dbo.tGood.BuyPrice AS BuyPrice,
			dbo.tGood.FinalPrice AS FinalPrice,
			SUM(dbo.tFacD.Amount*tfacd.FeeUnit) AS SumPrice,
			SUM(dbo.tFacD.Amount* dbo.tGood.BuyPrice) AS SumBuyPrice,
			SUM(dbo.tFacD.Amount*dbo.tGood.FinalPrice) AS SumFinalPrice
			,@SystemDate AS SystemDate,
		    @SystemDay AS SystemDay,
		    @SystemTime AS SystemTime ,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE  dbo.tFacD.intInventoryNo = dbo.tInventory.InventoryNo ) AS DepartureInventory,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE dbo.tFacD.DestInventoryNo = dbo.tInventory.InventoryNo ) AS DestinationInventory
	FROM dbo.tFacM 
		   INNER JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch 
			AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	WHERE dbo.tFacD.DestInventoryNo >= @Inventory1
			AND dbo.tFacD.DestInventoryNo <= @Inventory2
			AND dbo.tFacM.Date>=@Date1
			AND dbo.tFacM.Date<=@Date2
			AND dbo.tFacM.Status=6
	GROUP BY dbo.tFacM.No,
			 dbo.tFacD.GoodCode,
			 dbo.tGood.Name,
			 dbo.tFacD.FeeUnit,
--			 dbo.tFacD.Amount,
			 dbo.tGood.BuyPrice,
			 dbo.tGood.FinalPrice ,
			 dbo.tFacD.intInventoryNo ,
			 dbo.tFacD.DestInventoryNo
	ORDER BY dbo.tFacM.No

END
--
--
--EXEC Get_AssignmentDetail '','','',N'89/02/06',N'89/02/06',1,101

GO

--exec dbo.Get_AssignmentDetail N'93/04/30', N'دو شنبه', N'02:05', N'93/04/30', N'93/04/30', 1, 100
--GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_AssignmentDetailByFeeUnit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_AssignmentDetailByFeeUnit]
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


-------------------------گزارش جزييات حوالهاي صادره كالا------------------------
CREATE  PROC	[dbo].[Get_AssignmentDetailByFeeUnit](   
	@SystemDate NVARCHAR(20) ,
	@SystemDay NVARCHAR(20) ,
	@SystemTime NVARCHAR(20) ,
	@Date1 NVARCHAR(8),
	@Date2 NVARCHAR(8),
	@Inventory1 INT,
	@Inventory2 INT)
AS	
BEGIN

	SELECT  --dbo.tFacM.No,
		dbo.tGood.Name AS GoodName,
		dbo.tGood.TechnicalNo,
		SUM(dbo.tFacD.Amount) AS Amount,
		dbo.tFacD.FeeUnit ,
		dbo.tGood.BuyPrice AS BuyPrice,
		dbo.tGood.FinalPrice AS FinalPrice,
		SUM(dbo.tFacD.Amount*tfacd.FeeUnit) AS SumPrice,
		SUM(dbo.tFacD.Amount* dbo.tGood.BuyPrice) AS SumBuyPrice,
		SUM(dbo.tFacD.Amount*dbo.tGood.FinalPrice) AS SumFinalPrice,
		@SystemDate AS SystemDate,
		@SystemDay AS SystemDay,
		@SystemTime AS SystemTime
	FROM dbo.tFacM 
		JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch 
			AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	WHERE dbo.tFacD.DestInventoryNo >=@Inventory1
			AND dbo.tFacD.DestInventoryNo <=@Inventory2
			AND dbo.tFacM.Date>=@Date1
			AND dbo.tFacM.Date<=@Date2
			AND dbo.tFacM.Status=6
	GROUP BY --dbo.tFacM.No,
			 dbo.tFacD.GoodCode,
			 dbo.tGood.Name,
			 dbo.tGood.TechnicalNo,
			 dbo.tFacD.FeeUnit,
--			 dbo.tFacD.Amount,
			 dbo.tGood.BuyPrice,
			 dbo.tGood.FinalPrice
	ORDER BY dbo.tFacD.FeeUnit

END

GO