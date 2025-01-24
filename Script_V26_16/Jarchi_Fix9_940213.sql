
--For Jarchi
--V26_16_Fix9
--94/02/13


UPDATE tfacm SET transferAccounting = 0 WHERE AccountYear = 1394
GO

DELETE FROM dbo.tblAcc_DocumentHeader
Go

DELETE FROM dbo.tblAcc_Tafsili WHERE TafsiliId > 0
GO

UPDATE dbo.tPer SET Tafsili = NULL 
UPDATE dbo.tCust SET Tafsili = NULL
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER     PROCEDURE [dbo].[Get_SaleSummary]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
SELECT SUM(SumPrice)AS SumPriceTotal ,
        tvw.[Date] ,
        Branch ,
        Tafsili ,
        InventoryName

FROM 
(
SELECT DISTINCT dbo.tFacM.Branch  ,--NO ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacD.intRow ,
                    tfacd.Amount ,
                    tfacd.Feeunit ,
                    ( tfacd.Amount * tfacd.Feeunit ) AS SumPrice ,
                    ISNULL(dbo.tInventory.Tafsili ,0) AS Tafsili ,
                    dbo.tInventory.Description AS InventoryName
          FROM      dbo.tFacM
                    INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                                        AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
					INNER JOIN dbo.tInventory ON dbo.tInventory.InventoryNo = dbo.tFacD.intInventoryNo
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND (dbo.tCust.Tafsili = 0 OR dbo.tCust.Tafsili IS NULL) ))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch , tvw.Tafsili , InventoryName
 ORDER BY tvw.[Date] 
 
 
END

GO

UPDATE tblTotal_ItemReports_Details
SET ComboQuery = 'SELECT MembershipId ,( Family  + Name + WorkName) AS nvcName FROM tcust WHERE Code > 0 AND MembershipId > 0  ORDER BY nvcName ' 
	, ComboFieldCode = 'MembershipId' , ComboFieldDescr = 'nvcName' , ParameterType = 5 , parameterLengh = 4 , ObjectType = 1
WHERE intReportId = 39 AND Row = 2
GO

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER   view VwCustomerPurchaseDetails_Goods
as

SELECT     TOP 100 PERCENT dbo.tFacM.Customer AS CustCode, dbo.tFacM.[No], dbo.tFacM.SumPrice, dbo.tFacM.StationID, dbo.tFacM.[Date], dbo.tFacM.[Time],dbo.tCust.memberShipId, 
                      dbo.tFacD.Amount, dbo.tFacD.GoodCode, dbo.tFacD.FeeUnit, dbo.tGood.[Name], 
                      dbo.tFacM.[User], isnull(dbo.tFacM.Incharge, 0) AS CarrierPpno, CASE rtrim(ltrim(dbo.tCust.[Name] + dbo.tCust.Family)) WHEN NULL 
                      THEN isnull(dbo.tCust.WorkName, '') WHEN '' THEN isnull(dbo.tCust.WorkName, '') ELSE dbo.tCust.[Name] + ' ' + dbo.tCust.Family END AS FullName, 
                      CASE dbo.tCust.Sex WHEN 1 THEN N'ÂÞÇí' WHEN 0 THEN N'ÎÇäã' ELSE N'' END AS Gender
			, tFacm.Status , tfacm.Branch , tfacM.ServiceTotal , tfacM.TaxTotal , tfacM.DutyTotal
FROM         dbo.tFacM INNER JOIN
              dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND dbo.tFacM.Branch = dbo.tFacD.Branch INNER JOIN
              dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code INNER JOIN 
              dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer   LEFT OUTER JOIN --AND  dbo.tCust.Branch = dbo.tFacM.Branch
              dbo.tPer ON dbo.tFacM.InCharge = dbo.tPer.pPno --and  dbo.tFacM.Branch = dbo.tPer.Branch
WHERE     (dbo.tFacM.Customer <> - 1) 
--                   AND (dbo.tPer.job = 3 OR    dbo.tPer.job IS NULL) 
                   AND  (dbo.tFacM.Recursive = 0)    AND  (dbo.tFacM.Status = 2 or dbo.tFacM.Status = 5)-- AND  (dbo.tFacM.Branch = dbo.Get_Current_Branch())
ORDER BY dbo.tFacM.[Date], dbo.tFacM.[Time]





GO



ALTER  PROCEDURE dbo.GetCustPurchaseDetailsInfo_Goods
(
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(10),
      @Date2 NVARCHAR(10),
      @Customer1 NVARCHAR(50),
      @Customer2 NVARCHAR(50),
      @Branch1 INT,
      @Branch2 INT	
)
AS

    DECLARE @Tmp NVARCHAR(50)
    DECLARE @FromCustCode BIGINT      
    DECLARE @ToCustCode BIGINT      
      
	SET @FromCustCode = CAST(@Customer1 AS BIGINT)
	SET @ToCustCode = CAST(@Customer2 AS BIGINT)
    IF @Date2 < @Date1 
        BEGIN
            SET @Tmp = @Date2
            SET @Date2 = @Date1
            SET @Date1 = @Tmp
        END

SELECT  --dbo.VwCustomerPurchaseDetails_Goods.[No],
               -- dbo.VwCustomerPurchaseDetails_Goods.SumPrice,dbo.VwCustomerPurchaseDetails_Goods.StationId,
               -- dbo.VwCustomerPurchaseDetails_Goods.[Date],dbo.VwCustomerPurchaseDetails_Goods.[Time],
               dbo.VwCustomerPurchaseDetails_Goods.MemberShipId,--dbo.VwCustomerPurchaseDetails_Goods.Amount,
               dbo.VwCustomerPurchaseDetails_Goods.GoodCode,
		dbo.VwCustomerPurchaseDetails_Goods.FeeUnit,
               dbo.VwCustomerPurchaseDetails_Goods.[Name] as GoodName,dbo.VwCustomerPurchaseDetails_Goods.FullName,
			@Date1 AS DateBefore , @Date1 AS DateAfter , 
	       @FromCustCode As FromCustomer , @ToCustCode AS ToCustomer,
	 	@SystemDay + ' ' + @SystemDate +' '+N' ÓÇÚÊ : ' + @SystemTime AS Sysdate
                
               ,SUM(dbo.VwCustomerPurchaseDetails_Goods.Amount)AS SumAmount
                 
		  ,dbo.VwCustomerPurchaseDetails_Goods.FeeUnit * SUM(dbo.VwCustomerPurchaseDetails_Goods.Amount) AS PriceTotal
		            , VwCustomerPurchaseDetails_Goods.No
		            , VwCustomerPurchaseDetails_Goods.Date
					 , VwCustomerPurchaseDetails_Goods.ServiceTotal 
					 , (VwCustomerPurchaseDetails_Goods.TaxTotal + VwCustomerPurchaseDetails_Goods.DutyTotal) AS TaxTotal
					 , VwCustomerPurchaseDetails_Goods.SumPrice
        FROM 	dbo.VwCustomerPurchaseDetails_Goods

	WHERE   
		dbo.VwCustomerPurchaseDetails_Goods.[Date] >= @Date1  
		AND dbo.VwCustomerPurchaseDetails_Goods.[Date] <= @Date2 
		AND dbo.VwCustomerPurchaseDetails_Goods.memberShipId >= @FromCustCode 
		AND dbo.VwCustomerPurchaseDetails_Goods.memberShipId <= @ToCustCode

        GROUP BY [Branch], dbo.VwCustomerPurchaseDetails_Goods.[Name]
		-- ,dbo.VwCustomerPurchaseDetails_Goods.Amount
       		 ,dbo.VwCustomerPurchaseDetails_Goods.MemberShipId
		 ,dbo.VwCustomerPurchaseDetails_Goods.FeeUnit
		 ,dbo.VwCustomerPurchaseDetails_Goods.FullName
		 ,dbo.VwCustomerPurchaseDetails_Goods.GoodCode
            , VwCustomerPurchaseDetails_Goods.No
            , VwCustomerPurchaseDetails_Goods.Date
			 , VwCustomerPurchaseDetails_Goods.ServiceTotal 
			 , VwCustomerPurchaseDetails_Goods.TaxTotal 
			 , VwCustomerPurchaseDetails_Goods.DutyTotal
			 , VwCustomerPurchaseDetails_Goods.SumPrice
			 
       Order By	dbo.VwCustomerPurchaseDetails_Goods.memberShipId,
		dbo.VwCustomerPurchaseDetails_Goods.SumAmount desc ,
                 dbo.VwCustomerPurchaseDetails_Goods.GoodName

GO



