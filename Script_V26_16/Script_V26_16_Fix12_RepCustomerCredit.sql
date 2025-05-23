
--برای همه ورژن های رستورانی
--Script_V26_16_Fix12_RepCustomerCredit.SQL
--ReportName : RepCustomerPayRec.rpt
--94/03/30


--RepCustomerBuyDetails_Goods.rpt
--V26_16
--94/02/13


UPDATE tblTotal_ItemReports_Details
SET ComboQuery = 'SELECT MembershipId ,( Family  + Name + WorkName) AS nvcName FROM tcust WHERE Code > 0 AND MembershipId > 0  AND Credit > 0  ORDER BY MembershipId ' 
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
                      CASE dbo.tCust.Sex WHEN 1 THEN N'آقاي' WHEN 0 THEN N'خانم' ELSE N'' END AS Gender
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
	 	@SystemDay + ' ' + @SystemDate +' '+N' ساعت : ' + @SystemTime AS Sysdate
                
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


UPDATE dbo.tObjects
SET ObjectName = N'گزارش لیست مشترکین' 
WHERE  intobjectcode = 397
GO

 

INSERT INTO dbo.tbltotal_ItemReports
        ( intReportId ,
          intGroupReportId ,
          ReportName ,
          LatinReportName ,
          Refrence_Sp 
        )
VALUES  ( 45 , -- intReportId - int
          8 , -- intGroupReportId - int
          N'گزارش صورتحساب مشتری اعتباری' , -- ReportName - nvarchar(100)
          'RepCustomerPayRec' , -- LatinReportName - varchar(50)
          ''  -- Refrence_Sp - varchar(50)
        )
GO

INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 398 , -- intObjectCode - int
          N'RepCustomerPayRec' , -- ObjectId - nvarchar(50)
          N'گزارش صورتحساب مشتری اعتباری' , -- ObjectName - nvarchar(50)
          N'RepCustomerPayRec' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          108  -- ObjectParent - int
        )
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          398  -- intObjectCode - int
          )
          
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
		  45 ,
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
FROM dbo.tblTotal_ItemReports_Details WHERE intReportId = 39 

GO

UPDATE tblTotal_ItemReports_Details
SET ComboQuery = 'SELECT MembershipId ,( Family  + Name + WorkName) AS nvcName FROM tcust WHERE Code > 0 AND MembershipId > 0  AND Credit > 0  ORDER BY MembershipId ' 
	, ComboFieldCode = 'MembershipId' , ComboFieldDescr = 'nvcName' , ParameterType = 5 , parameterLengh = 4 , ObjectType = 1
WHERE intReportId = 45 AND Row = 2
GO

DELETE FROM tblTotal_ItemReports_Details WHERE intReportId = 45 AND Row = 3
UPDATE tblTotal_ItemReports_Details SET Row = 3 WHERE intReportId = 45 AND Row = 4
GO






if exists (select * from dbo.sysobjects where id = object_id(N'Fn_Customer_TurnOver') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Fn_Customer_TurnOver]
GO


CREATE   Function [dbo].[Fn_Customer_TurnOver]
(
      @Date1 NVARCHAR(10) ,
      @Date2 NVARCHAR(10) ,
      @Cust1 BIGINT ,
      @Cust2 BIGINT ,
      @Branch1 INT ,
      @Branch2 INT ,
      @AccountYear INT  


) 
RETURNS  @ReturnTable table
			(
            [MembershipId] BIGINT  ,
            [nvcName] NVARCHAR(70)  ,
            [Tel1] NVARCHAR(50),
            [Address] NVARCHAR(250),
            TotalRemainingAmount BIGINT ,
            SumBuy BIGINT,
            SumSale BIGINT,
            ReturnSumSale BIGINT,
            ReturnSumBuy BIGINT,
            RecievedAmount BIGINT,
            PaidAmount BIGINT,
            TotalCreditDebit BIGINT
			)
	
	
As

BEGIN

DECLARE  @Remain1 INT 
SET @Remain1 = 1


Insert into @ReturnTable( MembershipId ,nvcName , Tel1 ,Address,  TotalRemainingAmount ,SumSale,ReturnSumSale
			,RecievedAmount ,PaidAmount ,TotalCreditDebit  )

    SELECT  
            [tCust].[MembershipId] ,
            ([tCust].Name + ' ' + tcust.Family + tcust.WorkName) AS nvcName ,
            [tCust].[Tel1] ,
            [tCust].[Address] ,
            (h.[TotalRemainingAmount] ) AS TotalRemainingAmount ,
            h.SumSale ,
            h.ReturnSumSale ,
            h.CustomerDaryaft  AS RecievedAmount ,
            h.CustomerPardakht AS PaidAmount ,
            (-1 * h.SumSale - h.CustomerPardakht + h.CustomerDaryaft 
                + h.ReturnSumSale ) AS TotalCreditDebit
    FROM    (
              SELECT    
                        SUM(SumSale) AS SumSale ,
                        SUM(ReturnSumSale) AS ReturnSumSale ,
                        SUM(CustomerDaryaft) AS CustomerDaryaft ,
                        SUM(CustomerPardakht) AS CustomerPardakht ,
						SUM(PreRemain) AS TotalRemainingAmount ,
                        intcode
              FROM      (
--======================  Customer & Supplier  Daryaft - ReturnBuy Daryaft ==============================
                          SELECT    Code_Bes AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    SUM(Bestankar)  AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    0 AS PreRemain
                          FROM      (
                                      SELECT    Bestankar  ,
                                                Code_Bes 

										FROM    tblAcc_Recieved 
											INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tblAcc_Recieved.Code_Bes
										WHERE	RecieveType = 3 AND AccountYear = @AccountYear
											    AND [MembershipId] >= @Cust1
                                                AND MembershipId <= @Cust2
                                                AND tblAcc_Recieved.[Date] >= @Date1
                                                AND tblAcc_Recieved.[Date] <= @Date2
                                                AND tblAcc_Recieved.Branch = @Branch1 AND tblAcc_Recieved.Branch <= @Branch2
                                     
                                    ) AS a
                          GROUP BY  Code_Bes
                          UNION ALL
--======================  Customer & Supplier  Daryaft - ReturnBuy Daryaft ==============================
                          SELECT    Customer AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    SUM(CustomerDaryaft)  AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    0 AS PreRemain
                          FROM      (
                                      SELECT    intAmount AS CustomerDaryaft ,
                                                Customer

										FROM    dbo.tFacCash
											INNER JOIN dbo.tFacM ON dbo.tFacCash.Branch = dbo.tFacM.Branch AND dbo.tFacCash.intSerialNo = dbo.tFacM.intSerialNo
											INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer
										WHERE  Status = 2 AND AccountYear = @AccountYear
                                                AND    [MembershipId] >= @Cust1
                                                        AND MembershipId <= @Cust2
                                                AND tFacM.[Date] >= @Date1
                                                AND tFacM.[Date] <= @Date2
                                                AND Recursive <> 1
                                                AND tFacM.Branch = @Branch1 AND tFacM.Branch <= @Branch2
                                      
                                    ) AS b
                          GROUP BY  Customer
                          UNION ALL
--======================= Customer & Supplier Pardakht - SaleReturn Pardakht ===================
                          SELECT    Uid_Bede AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0 AS CustomerDaryaft ,
                                    SUM(CustomerPardakht) AS CustomerPardakht ,
                                    0 AS PreRemain
                          FROM      (
                                      SELECT    Bestankar AS CustomerPardakht ,
                                                Uid_Bede 
                                      FROM      dbo.tblAcc_Cash INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tblAcc_Cash.Uid_Bede
                                                WHERE PaymentType = 6 AND AccountYear = @AccountYear AND 
                                                   [MembershipId] >= @Cust1
                                                        AND MembershipId <= @Cust2
                                                AND tblAcc_Cash.[Date] >= @Date1
                                                AND tblAcc_Cash.[Date] <= @Date2
                                                AND tCust.Branch >= @Branch1 AND tCust.Branch <= @Branch2
                                      
                                    ) AS c
                          GROUP BY  Uid_Bede

                          UNION ALL

--======================= Sum Sale ======================================================

                          SELECT    [Customer] AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    SUM(CAST([SumPrice] AS BIGINT )) AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0 AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    0 AS PreRemain
                          FROM      tfacm INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer
                          WHERE     [MembershipId] >= @Cust1
                                    AND MembershipId <= @Cust2
                                    AND recursive <> 1 AND AccountYear = @AccountYear
                                    AND status = 2
                                    AND tfacm.[Date] >= @Date1
                                    AND tfacm.[Date] <= @Date2
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                          GROUP BY  [tFacM].[Customer]

                          UNION ALL

--======================= Sum SaleReturn ========================================================
                          SELECT    [Customer] AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    SUM(CAST([SumPrice] AS BIGINT )) AS ReturnSumSale ,
                                    0 AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    0 AS PreRemain
                          FROM      tfacm INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer
                          WHERE     [MembershipId] >= @Cust1
                                    AND MembershipId <= @Cust2
                                    AND recursive <> 1 AND AccountYear = @AccountYear
                                    AND status = 5
                                    AND tfacm.[Date] >= @Date1
                                    AND tfacm.[Date] <= @Date2
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                          GROUP BY  [tFacM].[Customer]


						UNION ALL

--======================  Customer & Supplier  Daryaft - ReturnBuy Daryaft ==============================
                          SELECT    Code_Bes AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0  AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    SUM(CustomerDaryaft) AS PreRemain
                          FROM      (
                                      SELECT    Bestankar AS CustomerDaryaft ,
                                                Code_Bes 
										FROM    tblAcc_Recieved 
											INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tblAcc_Recieved.Code_Bes
										WHERE	RecieveType = 3 AND AccountYear = @AccountYear AND 
											  [MembershipId] >= @Cust1
                                                AND MembershipId <= @Cust2
                                                AND tblAcc_Recieved.[Date] < @Date1
                                                AND tCust.Branch = @Branch1 AND tCust.Branch <= @Branch2
                                                AND @Remain1 = 1 
                                      
                                    ) AS a
                          GROUP BY  Code_Bes
                          UNION ALL
--======================  Customer & Supplier  Daryaft - ReturnBuy Daryaft ==============================
                          SELECT    Customer AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0  AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    SUM(CustomerDaryaft) AS PreRemain
                          FROM      (
                                      SELECT    intAmount AS CustomerDaryaft ,
                                                Customer 

										FROM    dbo.tFacCash
											INNER JOIN dbo.tFacM ON dbo.tFacCash.Branch = dbo.tFacM.Branch AND dbo.tFacCash.intSerialNo = dbo.tFacM.intSerialNo
											INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer
										WHERE  Status = 2 AND AccountYear = @AccountYear
                                                AND    [MembershipId] >= @Cust1
                                                        AND MembershipId <= @Cust2
                                                AND dbo.tFacM.[Date] < @Date1
                                                AND Recursive <> 1
                                                AND tFacM.Branch = @Branch1 AND tFacM.Branch <= @Branch2
                                                AND @Remain1 = 1 
                                      
                                    ) AS b
                          GROUP BY  Customer
                          UNION ALL
--======================= Customer & Supplier Pardakht - SaleReturn Pardakht ===================
                          SELECT    Uid_Bede AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0 AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    -1 * SUM(CustomerPardakht) AS PreRemain
                          FROM      (
                                      SELECT    Bestankar AS CustomerPardakht ,
                                                Uid_Bede 
                                      FROM      dbo.tblAcc_Cash INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tblAcc_Cash.Uid_Bede
                                                WHERE PaymentType = 6 AND AccountYear = @AccountYear AND 
                                                   [MembershipId] >= @Cust1
                                                        AND MembershipId <= @Cust2
                                                AND tblAcc_Cash.[Date] < @Date1
                                                AND tCust.Branch >= @Branch1 AND tCust.Branch <= @Branch2
                                                AND @Remain1 = 1
                                      
                                    ) AS c
                          GROUP BY  Uid_Bede

                          UNION ALL

--======================= Sum Sale =====================================================
                          SELECT    [Customer] AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0 AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    -1 * SUM(CAST([SumPrice] AS BIGINT )) AS PreRemain 
                          FROM      tfacm INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer
                          WHERE     [MembershipId] >= @Cust1
                                            AND MembershipId <= @Cust2
                                    AND recursive <> 1 AND AccountYear = @AccountYear
                                    AND status = 2
                                    AND dbo.tFacM.[Date] < @Date1
                                    AND tfacm.Branch = @Branch1  AND tfacm.Branch <= @Branch2 
                                    AND @Remain1 = 1
                          GROUP BY  [tFacM].[Customer]

                          UNION ALL

--======================= Sum SaleReturn =====================================================
                          SELECT    [Customer] AS intcode ,
                                    0 AS SumBuy ,
                                    0 AS ReturnSumBuy ,
                                    0 AS SumSale ,
                                    0 AS ReturnSumSale ,
                                    0 AS CustomerDaryaft ,
                                    0 AS CustomerPardakht ,
                                    SUM(CAST([SumPrice] AS BIGINT )) AS PreRemain
                          FROM      tfacm INNER JOIN dbo.tCust ON dbo.tCust.Code = dbo.tFacM.Customer
                          WHERE     [MembershipId] >= @Cust1
                                            AND MembershipId <= @Cust2
                                    AND recursive <> 1 AND AccountYear = @AccountYear
                                    AND status = 5
                                    AND dbo.tFacM.[Date] < @Date1
                                    AND tfacm.Branch = @Branch1  AND tfacm.Branch <= @Branch2 
                                    AND @Remain1 = 1
                          GROUP BY  [tFacM].[Customer]
			

                        ) t
              GROUP BY  [intcode]
            ) AS h
            INNER JOIN tcust ON h.intcode = [tCust].[Code]


Return


End



GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Customer_Credit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Customer_Credit]
GO


CREATE PROCEDURE [dbo].[Get_Customer_Credit]    
(    
 @SystemDate   NVARCHAR(50),    
 @SystemDay    NVARCHAR(50),    
 @SystemTime   NVARCHAR(50),    
 @Date1   NVARCHAR(10),    
 @Date2   NVARCHAR(10),    
 @Customer1   INT  ,    
 @Customer2   INT  ,
 @Branch1 INT ,
 @Branch2 INT 

    
)    

AS    
    
    DECLARE @AccountYear INT 
    SET @AccountYear = CAST('13' + LEFT(@Date2 ,2) AS int )
    
    SELECT  @SystemDay + ' ' + @SystemDate + ' ' + N' ساعت : ' + @SystemTime AS Sysdate ,
            [MembershipId] ,
            [nvcName] ,
            [Tel1] ,
            [Address] ,
            -1 * TotalRemainingAmount AS TotalRemainingAmount,
            SumSale ,
            ReturnSumSale ,
            RecievedAmount ,
            PaidAmount ,
            -1 * TotalCreditDebit AS TotalCreditDebit
            FROM dbo.[Fn_Customer_TurnOver](@Date1 , @Date2 , @Customer1 , @Customer2 , @Branch1 , @Branch2 , @AccountYear  )
            ORDER BY MembershipId 	



GO


