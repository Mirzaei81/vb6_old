
--Script Speed Of Crystal Report 
--افزایش سرعت ریپورت گروهی کالا
--V26_15 , V26_16 , V26_17 ,  روی  ورژن  رستورانی

----------------------------------------------------------------------------------------------------------
ALTER  VIEW dbo.ViewRepSellKind
AS
SELECT     TOP 100 PERCENT 
	    dbo.tFacD.Amount,dbo.tFacD.GoodCode, dbo.tFacD.FeeUnit, dbo.tGood.NamePrn,dbo.tGood.BarCode, 
        dbo.tGood.Level1 AS Level1Code, dbo.tGoodLevel1.Description AS Level1Desc,  
        dbo.tGoodLevel2.Description AS Level2Desc, dbo.tGood.Level2 AS Level2Code, 
        dbo.tFacM.Status, 
	    dbo.tFacM.CarryFeeTotal, 
        dbo.tFacM.SumPrice, 
        dbo.tFacM.StationID, dbo.tFacM.ServiceTotal, dbo.tFacM.PackingTotal, 
        dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.Branch, 
        tfacd.Discount as Discountrate, 
        dbo.tFacD.Discount, dbo.tFacM.AccountYear  ,
	    tBranch.nvcBranchName,tUnitGood.Description as UnitGoodDescription
		,tfacm.TaxTotal , tfacm.DutyTotal , 
		dbo.tFacM.ServePlace,(SELECT [Description] FROM tserveplace WHERE intserveplace=tfacm.[ServePlace]) AS serveplacedesc
FROM         dbo.tFacM INNER JOIN
                      dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND dbo.tFacM.Branch = dbo.tFacD.Branch INNER JOIN
                      dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code INNER JOIN
		      dbo.tBranch ON dbo.tfacm.Branch = dbo.tBranch.Branch INNER JOIN
                      dbo.tGoodLevel1 ON dbo.tGood.Level1 = dbo.tGoodLevel1.Code INNER JOIN
		      dbo.tUnitGood on dbo.tgood.unit=dbo.tUnitGood.code inner join	
                      dbo.tGoodLevel2 ON dbo.tGood.Level2 = dbo.tGoodLevel2.Code  
		     
WHERE     (dbo.tFacM.Recursive = 0)
ORDER BY dbo.tFacm.INTSERIALNO



GO


--افزایش سرعت  گزارشات گروهی کالا و ریز کالا  

-------------------------------------------- 'Recipt گزارش فروش ريزكالا وتخفيفات'----------------------------------------------------

ALTER PROCEDURE [dbo].[GetSellKindInfo]
    (
      @SystemDate NVARCHAR(50) ,
      @SystemDay NVARCHAR(50) ,
      @SystemTime NVARCHAR(50) ,
      @Date1 VARCHAR(50) ,
      @Date2 VARCHAR(50) ,
      @User1 INT ,
      @User2 INT ,
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
    DECLARE @ServiceTotal INT
    DECLARE @DiscountTotal INT
    DECLARE @CarryFeeTotal INT
    DECLARE @PackingTotal INT
    DECLARE @TaxTotal INT
    DECLARE @DutyTotal INT
    DECLARE @Time3 NVARCHAR(50)
    DECLARE @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @tmp2

	
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
	
     SELECT   @TaxTotal = SUM(TaxTotal) ,
			  @DutyTotal = SUM(DutyTotal) ,
			  @PackingTotal = SUM(PackingTotal) ,
			  @CarryFeeTotal = SUM(CarryFeeTotal) ,
			  @DiscountTotal = SUM(DiscountTotal) ,
			  @ServiceTotal =  SUM(ServiceTotal)
                          FROM      tfacm
                          WHERE     [date] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = @Status1
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
                                    AND [User] >= @User1 AND [User] <= @User2
                         

    DECLARE @TimeTitle NVARCHAR(10)
        SET @TimeTitle = N' ساعت : '

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
    SELECT  SUM(dbo.ViewRepSellKind.Amount) AS SumAmount ,
            dbo.ViewRepSellKind.FeeUnit ,
            dbo.ViewRepSellKind.FeeUnit * SUM(dbo.ViewRepSellKind.Amount) AS PriceTotal ,
            dbo.ViewRepSellKind.Level1Code ,
            dbo.ViewRepSellKind.Level2Code ,
            dbo.ViewRepSellKind.UnitGoodDescription ,
            @Date1 AS DateBefore ,
            @Date2 AS DateAfter ,
            dbo.ViewRepSellKind.GoodCode ,
            dbo.ViewRepSellKind.Barcode ,
            @User1 AS FromUser ,
            @User2 AS ToUser ,
            @Time1 AS FromTime ,
            @Time2 AS ToTime ,
            @SystemDay + ' ' + @SystemDate + @TimeTitle + @SystemTime AS Sysdate ,
            @level11 AS FromGoodCodeLvele1 ,
            @level21 AS FromGoodCodeLvele2 ,
            @level12 AS ToGoodCodeLevel1 ,
            @level22 AS ToGoodCodeLevel2 ,
			dbo.ViewRepSellKind.NamePrn AS GoodName ,
            dbo.ViewRepSellKind.Level1Desc AS Level1Description ,
            dbo.ViewRepSellKind.Level2Desc AS Level2Description ,
            dbo.ViewRepSellKind.DiscountRate ,
            dbo.ViewRepSellKind.nvcBranchName ,
            @ServiceTotal AS ServiceTotal ,
            @DiscountTotal AS DiscountTotal ,
            @CarryFeeTotal AS CarryFeeTotal ,
            @PackingTotal AS PackingTotal ,
            @Status1 AS Status ,
            ViewRepSellKind.AccountYear ,
            ViewRepSellKind.Branch 
          --,dbo.ViewRepSellKind.ServePlace
	    , @taxTotal AS TaxTotal , @DutyTotal AS DutyTotal , ViewRepSellKind.serveplacedesc ,ViewRepSellKind.ServePlace 
		, @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2 , Branch 
     FROM    dbo.ViewRepSellKind
    WHERE   dbo.ViewRepSellKind.[date] >= @Date1
            AND dbo.ViewRepSellKind.[date] <= @Date2
	--AND dbo.ViewRepSellKind.[Supplier] IN (SELECT code FROM tSupplier WHERE code  >= @Sup1 AND code < =@Sup2)
            AND dbo.ViewRepSellKind.[User] >= @User1
                  AND dbo.ViewRepSellKind.[User] <= @User2

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
            dbo.ViewRepSellKind.Barcode ,
            dbo.ViewRepSellKind.DiscountRate ,
            ViewRepSellKind.AccountYear ,
            ViewRepSellKind.Branch ,
            ViewRepSellKind.nvcBranchName
	        ,ViewRepSellKind.ServePlace
	        ,ViewRepSellKind.[serveplacedesc]
    ORDER BY dbo.ViewRepSellKind.GoodCode





GO
