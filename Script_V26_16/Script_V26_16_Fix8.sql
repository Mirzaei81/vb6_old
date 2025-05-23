
--Script_V26_16_Fix8
--در نسخه های الماس
--گزارشات مصرف طی دوره 
--گزارش از ورود و خروج کالاها
--گزارش موجودی ریالی انبار در بازه زمانی
--گزارش موجودی تعدادی انبار در بازه زمانی
--فرم سود و زیان بازرگانی کالاها
--فرم سود و زیان کلی رستوران بدون داشتن حسابداری
--
-- 93/07/25

INSERT  INTO dbo.tblPub_Script2
        ( Version ,
          Script ,
          LastScriptNo ,
          [Date] ,
          FixNumber
        )
VALUES  ( 26 ,
          16 ,
          0 ,
          dbo.shamsi(GETDATE()) ,
          8
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
VALUES  ( 345 , -- intObjectCode - int
          N'frmTarazSoodZian' , -- ObjectId - nvarchar(50)
          N'سود و زیان کلی' , -- ObjectName - nvarchar(50)
          N'frmTarazSoodZian' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          102  -- ObjectParent - int
        )

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          345  -- intObjectCode - int
          )


go


IF COL_LENGTH('tGood','AvgBuyPrice') IS NULL
BEGIN
	ALTER TABLE tGood
	ADD AvgBuyPrice INT NOT NULL DEFAULT(0)
END

GO

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Update_tGood_By_AvgBuyPrice] 
(
	@NotSupportedGoodType int,
	@nvcFromDate NVARCHAR(10),
	@nvcToDate NVARCHAR(10)
) 
AS

DECLARE @AccountYear SMALLINT
SET @AccountYear = dbo.Get_AccountYear()


UPDATE dbo.tGood 
	SET dbo.tGood.AvgBuyPrice=(CASE ISNULL(T.AvgBuyPrice,0) when 0 then dbo.tGood.BuyPrice
					ELSE  ISNULL(T.AvgBuyPrice,0) END)
FROM

(
SELECT     FacdBuyPrice.GoodCode,CASE (FacdBuyPrice.SAmount + FacdBuyPrice.FirstMojodi)  WHEN 0 THEN 0
					ELSE CAST((FacdBuyPrice.SFeeUnit + FacdBuyPrice.FirstMojodiPrice)/(FacdBuyPrice.SAmount + FacdBuyPrice.FirstMojodi) AS int)END AS AvgBuyPrice
	FROM         dbo.vw_Good_Levels
			INNER JOIN
			(

			SELECT tInventory_Good.GoodCode , SUM(FirstMojodi * FirstPrice) AS FirstMojodiPrice , SUM(FirstMojodi) AS  FirstMojodi
				--,ISNULL(MAX(T2.SFeeUnit) ,0) AS SFeeUnit  , ISNULL(MAX(T2.SAmount),0) AS SAmount
				,MAX(ISNULL(T2.SFeeUnit,0)) AS SFeeUnit  , MAX(ISNULL(T2.SAmount,0)) AS SAmount
				FROM dbo.tInventory_Good
			LEFT OUTER JOIN 
			(
				SELECT     GoodCode, SUM(FeeUnit*Amount)AS SFeeUnit,SUM(Amount)AS SAmount
					FROM         dbo.tFacm
							INNER JOIN
							 dbo.tFacd ON dbo.tFacD.Branch = dbo.tFacM.Branch 
								AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
						
					WHERE  dbo.tFacM.Status=1
						AND (dbo.tFacM.[Date]>=@nvcFromDate AND dbo.tFacM.[Date]<=@nvcToDate) AND AccountYear = @AccountYear
					GROUP BY GoodCode
				)T2 ON tInventory_Good.GoodCode = T2.GoodCode AND AccountYear = @AccountYear
				GROUP BY tInventory_Good.GoodCode
			)
			AS FacdBuyPrice ON dbo.vw_Good_Levels.Code=FacdBuyPrice.GoodCode
			
	where (dbo.vw_Good_Levels.GoodType <> @NotSupportedGoodType OR @NotSupportedGoodType = -1)
	--ORDER BY FacdBuyPrice.GoodCode

)AS T


WHERE dbo.tGood.Code=T.GoodCode

GO


--exec Update_tGood_By_AvgBuyPrice 2, N'93/01/01  ', N'93/07/25  '
--GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Update_FinalPrice] 
AS
 Update tGood
   set FinalPrice = T.Finalprice
	FROM
		(
		Select x.GoodCode , (1 + ISNULL(cast(PercentOverFlow as float) ,0)/100 )  *    (Sum(AvgBuyPrice *(fltusedvalue+isnull(Pert,0)))+isnull(ChargeCooking,0)+isnull(ChargeServe,0))  As FinalPrice  
			From tGood
			inner join 
			(Select GoodCode , GoodFirstcode ,fltusedvalue,Pert  From tGood 
			 inner join tUsepercent On tGood.Code = tUsepercent.Goodcode 
			Where GoodType = 2 And tUsepercent.intserveplace = 1
			) X 
			On X.Goodfirstcode = tGood.code
			left outer join 
			dbo.tblTotal_ChargeGood on dbo.tblTotal_ChargeGood.GoodCode=x.Goodcode
			Group By x.Goodcode,ChargeCooking,ChargeServe,PercentOverFlow
	
		)AS T
	Where Code = T.Goodcode

UPDATE dbo.tGood
	SET FinalPrice = AvgBuyPrice WHERE FinalPrice = 0

GO



SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

-------------------------------------*******************************************
---------------------------------------------------اصلاح گزارش مواد مصرفی********


ALTER  PROC Rep_UseOfGood
    (
      @SystemDate NVARCHAR(20),
      @SystemDay NVARCHAR(20),
      @SystemTime NVARCHAR(20),
      @Date1 NVARCHAR(8),
      @Date2 NVARCHAR(8),
      @Status1 INT,
      @InventoryNo1 INT,
      @Branch1 INT
							
    )
AS 

    BEGIN
        SELECT  SUM(T.Amount) AS Amount,
                T.[Name],
             --AS FeeUnit,
                T.GoodCode,
                ( SELECT    dbo.tUnitGood.Description
                  FROM      dbo.tUnitGood
                  WHERE     Code = T.Unit
                ) AS Unit,
                T.Discount,
                T.Weight,
                T.NumberOfUnit,
               -- Rate,
                ( SELECT    dbo.tInventory.Description
                  FROM      dbo.tInventory
                  WHERE     dbo.tInventory.InventoryNo = intInventoryNo
                ) AS InventoryName,
                SellPrice,
                BuyPrice,
                @SystemDate AS SystemDate,
                @SystemDay AS SystemDay,
                @SystemTime AS SystemTime ,
                AvgBuyPrice
				FROM    
				( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                            --dbo.tFacD.Rate,
                            dbo.tFacD.intInventoryNo,
                            dbo.tFacD.DestInventoryNo,
                            ( dbo.tFacD.Amount * dbo.tUsePercent.fltUsedValue ) AS Amount,
                            ( SELECT    Name
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [Name],
                            dbo.tGood.Weight,
                            ( SELECT    Unit
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS Unit,
                            dbo.tGood.NumberOfUnit,
                            dbo.tFacD.Discount,
                            ( SELECT    SellPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [SellPrice],
                            ( SELECT    BuyPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [BuyPrice] ,
                            ( SELECT    AvgBuyPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS AvgBuyPrice
						FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                    AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                            JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
						WHERE     tfacm.Branch = @Branch1
                            AND tfacm.Recursive = 0
                            AND tfacm.status = @Status1
                            AND tfacm.Date >= @Date1
                            AND tfacm.Date <= @Date2
                            AND tfacd.intInventoryNo = @InventoryNo1
                  UNION ALL
                  SELECT    dbo.tFacD.GoodCode,
                            --dbo.tFacD.Rate,
                            dbo.tFacD.intInventoryNo,
                            dbo.tFacD.DestInventoryNo,
                            dbo.tFacD.Amount,
                            dbo.tGood.Name,
                            dbo.tGood.Weight,
                            dbo.tGood.Unit,
                            dbo.tGood.NumberOfUnit,
                            dbo.tFacD.Discount,
                            dbo.tGood.SellPrice,
                            dbo.tGood.BuyPrice ,
                            dbo.tGood.AvgBuyPrice
                  FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                  WHERE     tfacm.Branch = @Branch1
                            AND tfacm.Recursive = 0
                            AND tfacm.status = @Status1
                            AND tfacm.Date >= @Date1
                            AND tfacm.Date <= @Date2
                            AND tfacd.intInventoryNo = @InventoryNo1
                            --AND dbo.tFacD.GoodCode NOT IN (
                            --SELECT  dbo.tUsePercent.GoodCode
                            --FROM    dbo.tUsePercent
                            --WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                            AND dbo.tGood.GoodType = 3
                ) T
        GROUP BY T.GoodCode,
                T.Name,
                T.Weight,
                T.Unit,
                --Rate,
                intInventoryNo,
                DestInventoryNo,
                T.Discount,
                T.NumberOfUnit,
                T.SellPrice,
                T.BuyPrice ,
                T.AvgBuyPrice
	
    END


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Fn_SoodZian_Sale') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].Fn_SoodZian_Sale
GO


CREATE Function [dbo].Fn_SoodZian_Sale

(
  @DateBefore NVARCHAR(8)  ,
  @DateAfter NVARCHAR(8)  ,
  @AccountYear SMALLINT ,
  @Branch INT 
)

RETURNS  @ReturnTable TABLE(
 TotalSellAmount BIGINT ,
 TotalSellReturnAmount BIGINT ,
 TotalFirstPrice BIGINT ,
 TotalBuyAmount BIGINT ,
 TotalBuyReturnAmount BIGINT ,
 TotalSaleDiscount BIGINT ,
 TotalBuyDiscount BIGINT ,

 TotalCareeFee BIGINT ,
 TotalPacking BIGINT ,
 TotalService BIGINT ,
 TotalTax BIGINT ,

 TotalLosses BIGINT ,
 TotalHoghough BIGINT ,
 TotalHazineTolid BIGINT , 
 TotalHazineTax BIGINT 
)	
As

BEGIN


	DECLARE @TotalSellAmount BIGINT
	DECLARE @TotalSellReturnAmount BIGINT
	DECLARE @TotalFirstPrice BIGINT
	DECLARE @TotalBuyAmount BIGINT
	DECLARE @TotalBuyReturnAmount BIGINT
	DECLARE @TotalSaleDiscount BIGINT
	DECLARE @TotalBuyDiscount BIGINT

	DECLARE @TotalCareeFee BIGINT
	DECLARE @TotalPacking BIGINT
	DECLARE @TotalService BIGINT
	DECLARE @TotalTax BIGINT

	DECLARE @TotalLosses BIGINT
	DECLARE @TotalHoghough BIGINT
	DECLARE @TotalHazineTolid BIGINT
	DECLARE @TotalHazineTax BIGINT
	

		Select @TotalSellAmount = SUM(Sumprice) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
		
		
		Select @TotalSellReturnAmount = SUM(Sumprice) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 5 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalFirstPrice =  SUM(FirstMojodi * FirstPrice) FROM dbo.tInventory_Good
		WHERE AccountYear = @AccountYear AND Branch = @Branch 

		Select @TotalBuyAmount = SUM(Sumprice) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 1 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalBuyReturnAmount = SUM(Sumprice) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 4 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalSaleDiscount = SUM(DiscountTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalBuyDiscount = SUM(DiscountTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 1 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalCareeFee = SUM(CarryFeeTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalPacking = SUM(PackingTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @Totalservice = SUM(ServiceTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
		Select @TotalTax = SUM(TaxTotal) + SUM(DutyTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			

		Select @TotalLosses = SUM(dbo.tgood.FinalPrice * Amount) FROM dbo.tFacM 
		INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
		WHERE Recursive = 0 AND Status = 3 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND dbo.tFacM.Branch = @Branch
			
			
		SET @TotalHoghough = 0
		--Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		--From tblAcc_DocumentDetail
		--INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		--Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		--AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		----AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate


		Select @TotalHazineTolid = SUM(dbo.tgood.FinalPrice * Amount) FROM dbo.tFacM 
		INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND dbo.tFacM.Branch = @Branch
			
			
		Select @TotalHazineTax = SUM(TaxTotal) + SUM(DutyTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 1 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,  Totalservice ,
			 TotalTax ,TotalLosses  , TotalHoghough  , TotalHazineTolid , TotalHazineTax )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  , @Totalservice ,
			 @TotalTax ,@TotalLosses  , @TotalHoghough  , @TotalHazineTolid ,@TotalHazineTax)
		            
RETURN 


End



GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_TarazSoodZian_Sale]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_TarazSoodZian_Sale]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE   PROCEDURE [dbo].[Get_TarazSoodZian_Sale]
    (
      @DateBefore NVARCHAR(8)  ,
      @DateAfter NVARCHAR(8)  ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS 


SELECT ISNULL(TotalSellAmount , 0) AS TotalSellAmount ,
       ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
       ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
       ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
       ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
       ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
       ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
       ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
       ISNULL(TotalPacking , 0) AS TotalPacking ,
       ISNULL(TotalService , 0) AS TotalService ,
       ISNULL(TotalTax , 0) AS TotalTax ,
       ISNULL(TotalLosses , 0) AS TotalLosses ,
       ISNULL(TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(TotalHazineTolid , 0) AS TotalHazineTolid ,
       ISNULL(TotalHazineTax , 0) AS TotalHazineTax 
       
	FROM dbo.Fn_SoodZian_Sale(@DateBefore ,@DateAfter ,@AccountYear ,@Branch )
--===============================================

GO


--exec Get_TarazSoodZian_Sale N'93/01/01', N'93/07/25', 1393, 1








--Script_V26_16_Fix7_Inventory
--گزارشات مصرف طی دوره 
-- 93/07/23


--SELECT * FROM dbo.tbltotal_ItemReports
--GO
--SELECT * FROM dbo.tbltotal_ItemReports_Details where intreportid = 95
--GO

INSERT INTO dbo.tbltotal_ItemReports
        ( intReportId ,
          intGroupReportId ,
          ReportName ,
          LatinReportName ,
          Refrence_Sp
        )
VALUES  ( 94 , -- intReportId - int
          2 , -- intGroupReportId - int
          N'گزارش از ورود و خروج کالاها' , -- ReportName - nvarchar(100)
          'RepMojodiControlGoods' , -- LatinReportName - varchar(50)
          'Get_InventoryMojodiControl'  -- Refrence_Sp - varchar(50)
        )
GO

UPDATE tbltotal_ItemReports_Details SET Quantity = 1 WHERE intReportId = 8 AND Row = 4
GO


--گزارش موجودی ریالی انبار در بازه زمانی

INSERT INTO dbo.tbltotal_ItemReports
        ( intReportId ,
          intGroupReportId ,
          ReportName ,
          LatinReportName ,
          Refrence_Sp
        )
VALUES  ( 95 , -- intReportId - int
          2 , -- intGroupReportId - int
          N'گزارش موجودی ریالی انبار در بازه زمانی' , -- ReportName - nvarchar(100)
          'RepInventoryAtomicRials_Report_InDate' , -- LatinReportName - varchar(50)
          'Get_InventoryAtomicRials_Report_InDate'  -- Refrence_Sp - varchar(50)
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
VALUES  ( 448 , -- intObjectCode - int
          N'RepMojodiControlGoods' , -- ObjectId - nvarchar(50)
          N'گزارش از ورود و خروج کالاها' , -- ObjectName - nvarchar(50)
          N'RepMojodiControlGoods' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          108  -- ObjectParent - int
        )
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          448  -- intObjectCode - int
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
VALUES  ( 449 , -- intObjectCode - int
          N'RepInventoryAtomicRials_Report_InDate' , -- ObjectId - nvarchar(50)
          N'گزارش موجودی ریالی انبار در بازه زمانی' , -- ObjectName - nvarchar(50)
          N'RepInventoryAtomicRials_Report_InDate' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          108  -- ObjectParent - int
        )
GO


INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          449  -- intObjectCode - int
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
		  94 ,
          1 ,
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
FROM tblTotal_ItemReports_Details WHERE intReportId = 81 AND Row = 1 

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
		  94 ,
          2 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          1 ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 81 AND Row = 2 

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
		  94 ,
          3 ,
          N'انبارفروش' ,
          ' ' ,
          'intInventory' ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          1 ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 81 AND Row = 2 

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
		  94 ,
          4 ,
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
FROM tblTotal_ItemReports_Details WHERE intReportId = 84 AND Row = 5 

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
		  94 ,
          5 ,
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
FROM tblTotal_ItemReports_Details WHERE intReportId = 85 AND Row = 2 

GO

--DELETE FROM tblTotal_ItemReports_Details WHERE intReportId = 95


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
		  95 ,
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
FROM tblTotal_ItemReports_Details WHERE intReportId = 8

GO


--SELECT * FROM dbo.tblTotal_ItemReports_Details WHERE intReportId = 94
--GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FnUsedGoods_FirstDate]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FnUsedGoods_FirstDate]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE  FUNCTION [dbo].[FnUsedGoods_FirstDate]
    (
      @DateAfter VARCHAR(10),
      @UsedInventory INT,
      @SaleInventory INT,
      --@Branch INT,
      @AccountYear SMALLINT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT ,
      @GoodCode INT   
    )
RETURNS @ReturnTable TABLE
    (
      FirstGoodCode INT,
      FirstDateMojodi FLOAT ,
      FirstDateUsed FLOAT
    )
AS 
BEGIN 

DECLARE @UsedBranch INT 
SET @UsedBranch = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @UsedInventory)

DECLARE @SaleBranch INT 
SET @SaleBranch = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @SaleInventory)

IF @GoodLevel11 = 0
 SELECT @GoodLevel11 = MIN(Code) FROM dbo.tGoodLevel1
IF @GoodLevel12 = 0
 SELECT @GoodLevel12 = Max(Code) FROM dbo.tGoodLevel1

    INSERT  INTO @ReturnTable
            (
              FirstGoodCode,
              FirstDateMojodi ,
              FirstDateUsed                
            )


SELECT X1.FirstGoodCode AS FirstGoodCode ,
			CAST(X2.FirstDateMojodi AS DECIMAL(20,3)) AS FirstDateMojodi  ,
			X1.FirstDateUsed AS FirstDateUsed 
	 FROM 
(
        SELECT    
			ISNULL(k.GoodCode,dbo.tInventory_Good.GoodCode) AS FirstGoodCode ,
			CAST(ISNULL(k.Amount, 0) AS DECIMAL(20,3)) AS FirstDateUsed 
			
			FROM      (

					SELECT  SUM(T5.Amount) AS Amount,
							T5.Branch,T5.AccountYear ,
							T5.GoodCode
							FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
										( dbo.tFacD.Amount * dbo.tUsePercent.fltUsedValue ) AS Amount,
										 dbo.tFacM.Branch , dbo.tFacM.AccountYear
										FROM      dbo.tFacM
										JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
														  AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
										JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
																AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
										JOIN dbo.tGood ON dbo.tUsePercent.GoodCode = dbo.tGood.Code
							  WHERE     tfacm.Branch = @SaleBranch
										AND dbo.tFacM.AccountYear = @AccountYear
										AND tfacm.Recursive = 0
										AND tfacm.status = 2
										AND tfacm.Date < @DateAfter
										AND tfacd.intInventoryNo = @SaleInventory
										AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
									AND (tFacd.GoodCode = @GoodCode OR @GoodCode = 0)
								)T5
						GROUP BY T5.GoodCode ,T5.Branch,T5.AccountYear
			) k
			
			INNER JOIN tinventory_good ON k.goodcode = tinventory_good.goodcode
			--AND (tinventory_good.Mojodi <> 0 OR @ZeroMojodi = 0)
			AND tinventory_good.inventoryno = @UsedInventory
			AND tinventory_good.AccountYear = k.AccountYear
			AND tinventory_good.Branch = k.Branch
        

)X1

INNER JOIN 

(
         SELECT 
			ISNULL(dbo.tInventory_Good.FirstMojodi ,0)
			+ ISNULL(SUM(k.TotalBuy), 0)
			- ISNULL(SUM(k.TotalHavaleh), 0)
			+ ISNULL(SUM(k.TotalResid), 0)
			- ISNULL(SUM(k.TotalLoss), 0)
			- ISNULL(SUM(k.TotalBuyReturn), 0) AS FirstDateMojodi,
			k.GoodCode AS FirstGoodCode 

			FROM   
			   (
				SELECT   
				DISTINCT [D].[GoodCode] ,D.intinventoryNo,
				M.AccountYear,M.Branch ,
				--------------------------------------------------------------------------------------------مقدار موجودي در  حالتهاي مختلف ----------------------------------------
				CASE WHEN [M].Status = 1 THEN SUM( [D].[Amount])ELSE 0 END AS TotalBuy ,
				CASE WHEN [M].Status = 3 THEN SUM([D].[Amount]) ELSE 0 END AS TotalLoss ,
				CASE WHEN [M].Status = 4 THEN SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturn ,
				CASE WHEN [M].Status = 6 THEN SUM([D].[Amount]) ELSE 0 END AS TotalHavaleh ,
				CASE WHEN [M].Status = 7 THEN SUM([D].[Amount]) ELSE 0 END AS TotalResid 

				-----------------------------------------------------------------------------------------------------------------------------------------------------------------
				FROM      [dbo].[tFacM] M
						INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
								AND [M].[intSerialNo] = [D].[intSerialNo]
						INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode 
				  WHERE     M.Branch = @UsedBranch
							AND M.AccountYear = @AccountYear
							AND M.Recursive = 0
							AND M.Date < @DateAfter
							AND D.intInventoryNo = @UsedInventory
				GROUP BY  [D].[GoodCode] ,
						[M].[Status] ,
						D.intInventoryNo ,
						M.AccountYear ,
						M.Branch
			) K
			INNER JOIN dbo.tInventory_Good 
				 ON k.goodcode = tInventory_Good.goodcode
				 AND K.AccountYear = tInventory_Good.AccountYear
				 AND K.Branch = tInventory_Good.Branch
				 AND K.intInventoryNo = tInventory_Good.inventoryno
			GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    tInventory_Good.firstMojodi ,
                    K.GoodCode ,
                    K.Branch
)X2
ON X1.FirstGoodCode = X2.FirstGoodCode
	ORDER BY   X1.FirstGoodCode

    RETURN
   END
--==========================================
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_InventoryMojodiControl]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_InventoryMojodiControl]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

------------------------------------------------------------گزارش ورود و خروج کالا---------------------------------
CREATE     PROCEDURE [dbo].[Get_InventoryMojodiControl]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(10),
      @Date2 NVARCHAR(10),
      @Inventory1 INT ,
      @intInventory1 INT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT,
      @AccountYear1 SMALLINT   
     
    )
AS 

DECLARE @Branch1 INT 
SET @Branch1 = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @Inventory1)

DECLARE @Branch2 INT 
SET @Branch2 = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @intInventory1)

SELECT DISTINCT 
ISNULL(T1.GoodCode , T2.[GoodCode]) AS [GoodCode] ,
ISNULL(T1.FirstDateUsed ,0) AS FirstDateUsed ,
ISNULL(FirstMojodi ,0) AS FirstMojodi ,
ISNULL(FromStoreAmount ,0) AS FromStoreAmount,
ISNULL(toStoreAmount ,0) AS toStoreAmount ,
ISNULL(BuyAmount ,0) AS BuyAmount ,
ISNULL(T2.Amount ,0) AS SaleAmount2,
ISNULL(Mojodi ,0) AS Mojodi ,
ISNULL(T1.NAME , T2.NAME) AS Name
,(Select  Description FROM dbo.tInventory WHERE InventoryNo = @Inventory1) AS DESCRIPTION 
,@Inventory1 AS InventoryNo,@AccountYear1 AS  AccountYear
,N'امروز'+' '+@SystemDay+'  '+N'مورخ '+@SystemDate+'  '+N'ساعت '+@SystemTime AS TotalDate
FROM 
(		

         SELECT 
            T4.Name ,
			ISNULL(T4.FirstDateUsed ,0) AS FirstDateUsed ,
			ISNULL(T4.FirstDateMojodi ,0) AS FirstMojodi ,
			ISNULL(T4.FirstDateMojodi ,0)
			+ ISNULL(SUM(k.TotalBuy), 0)
			- ISNULL(SUM(k.TotalHavaleh), 0)
			+ ISNULL(SUM(k.TotalResid), 0)
			- ISNULL(SUM(k.TotalLoss), 0)
			- ISNULL(SUM(k.TotalBuyReturn), 0) AS Mojodi,
			ISNULL(k.GoodCode,T4.GoodCode) AS GoodCode ,
			
			ISNULL(SUM(k.TotalBuy), 0) AS BuyAmount ,
			ISNULL(SUM(k.TotalLoss), 0) AS LossAmount ,
			ISNULL(SUM(k.TotalBuyReturn), 0) AS BuyReturnAmount,
			ISNULL(SUM(k.TotalHavaleh), 0) AS FromStoreAmount,
			ISNULL(SUM(k.TotalResid), 0) AS toStoreAmount


			FROM   
			   (
				SELECT   
				DISTINCT [D].[GoodCode] ,D.intinventoryNo,
				M.AccountYear,
				--------------------------------------------------------------------------------------------مقدار موجودي در  حالتهاي مختلف ----------------------------------------
				CASE WHEN [M].Status = 1 THEN SUM( [D].[Amount])ELSE 0 END AS TotalBuy ,
				CASE WHEN [M].Status = 3 THEN SUM([D].[Amount]) ELSE 0 END AS TotalLoss ,
				CASE WHEN [M].Status = 4 THEN SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturn ,
				CASE WHEN [M].Status = 6 THEN SUM([D].[Amount]) ELSE 0 END AS TotalHavaleh ,
				CASE WHEN [M].Status = 7 THEN SUM([D].[Amount]) ELSE 0 END AS TotalResid 

				-----------------------------------------------------------------------------------------------------------------------------------------------------------------
				FROM      [dbo].[tFacM] M
						INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
								AND [M].[intSerialNo] = [D].[intSerialNo]
						INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode 
								AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
				WHERE     [M].[Date] >= @Date1
						AND [M].[Date] <= @Date2
						AND [M].[AccountYear] = @AccountYear1
						AND [M].[Recursive] = 0
						AND [D].[intInventoryNo] = @Inventory1
						AND [M].Branch = @Branch1
				GROUP BY  [D].[GoodCode] ,
						[M].[Status] ,
						D.intInventoryNo ,
						M.AccountYear
			) K

			RIGHT OUTER  JOIN 
			(SELECT  GoodCode, FirstDateMojodi , Name  , FirstDateUsed
			  FROM 
				(SELECT  FirstGoodCode AS GoodCode, FirstDateMojodi , FirstDateUsed 
                       FROM    dbo.FnUsedGoods_FirstDate(@Date1, @Inventory1, @intInventory1, --@Branch1
                                              @AccountYear1 , @GoodLevel11 , @GoodLevel12 , 0 )
                                              --@AccountYear1 , @GoodLevel11 , @GoodLevel12 , 0 )
                                              --WHERE  FirstDateMojodi <> 0
				)T 
				INNER JOIN tGood ON T.GoodCode = dbo.tGood.Code AND (tgood.GoodType = 1 OR tgood.GoodType = 3)
			 )T4
			 ON k.goodcode = T4.goodcode
			 
          GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    T4.FirstDateMojodi ,
                    T4.GoodCode ,
                    T4.Name ,
                    T4.FirstDateUsed
)T1

FULL OUTER JOIN  
(

        SELECT  CAST(SUM(T5.Amount) AS DECIMAL(20,3)) AS Amount,
                T5.[Name],
                T5.GoodCode
                FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                            ( dbo.tFacD.Amount * dbo.tUsePercent.fltUsedValue ) AS Amount,
                            dbo.tGood.Name
							FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                    AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                            JOIN dbo.tGood ON dbo.tUsePercent.GoodFirstCode = dbo.tGood.Code
                  WHERE     tfacm.Branch = @Branch2
                            AND tfacm.Recursive = 0
                            AND tfacm.status = 2
                            AND tfacm.Date >= @Date1
                            AND tfacm.Date <= @Date2
                            AND tfacd.intInventoryNo = @intInventory1
							AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
					)T5
			GROUP BY T5.GoodCode,
					T5.Name


)T2

	ON T1.[GoodCode] = T2.[GoodCode]
WHERE 	
FirstMojodi <> 0 OR FromStoreAmount <> 0 OR toStoreAmount <> 0 OR BuyAmount <> 0 OR T2.Amount <> 0

ORDER BY ISNULL(T1.Name , T2.Name)

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FnFirstDateMojodi]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FnFirstDateMojodi]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE    FUNCTION [dbo].[FnFirstDateMojodi]
    (
      @DateAfter VARCHAR(10),
      @inventory INT,
      @Branch INT,
      @AccountYear SMALLINT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT ,
      @GoodCode INT  ,
      @ZeroMojodi INT  
    )
RETURNS @ReturnTable TABLE
    (
      GoodCode INT,
      FirstDateMojodi FLOAT ,
      FirstDateMojodiPrice BIGINT  
    )
AS 
BEGIN 

--Maybe @GoodCode = 0
--Maybe @InventoryNo = 0

IF @GoodLevel11 = 0
 SELECT @GoodLevel11 = MIN(Code) FROM dbo.tGoodLevel1
IF @GoodLevel12 = 0
 SELECT @GoodLevel12 = Max(Code) FROM dbo.tGoodLevel1

    INSERT  INTO @ReturnTable
            (
              GoodCode,
              FirstDateMojodi ,
              FirstDateMojodiPrice                
            )

SELECT M2.GoodCode , M2.FirstDateMojodi ,				 
		CASE WHEN (ISNULL(M2.FirstMojodiTotal,0) + ISNULL(M2.BuyTotal,0) - ISNULL(M2.BuyReturnTotal,0)) = 0 THEN 0
			WHEN  M2.FirstDateMojodi = 0 THEN 0
		ELSE 
		CAST(
		(ISNULL(M2.FirstMojodiPriceTotal ,0) + ISNULL(M2.BuyPriceTotal ,0) - ISNULL(M2.BuyReturnPriceTotal ,0)) 
		/ (ISNULL(M2.FirstMojodiTotal,0) + ISNULL(M2.BuyTotal,0) - ISNULL(M2.BuyReturnTotal,0) )
		AS BIGINT )END AS FirstDateMojodiPrice 

 from 
(
	SELECT M.GoodCode , M.FirstDateMojodi
	 ,
				(SELECT ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) 
							 FROM [dbo].[tFacM] M1
								INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
										AND [M1].[intSerialNo] = [D1].[intSerialNo]
								WHERE  
									M1.[date] < @DateAfter
									AND M1.Status = 1
									AND D1.GoodCode = M.GoodCode
									AND M1.AccountYear = @AccountYear
									AND M1.Branch = @Branch
									AND Recursive = 0
				  	) AS BuyPriceTotal,
				(SELECT ISNULL(SUM(D2.FeeUnit * D2.Amount ) ,0)
							 FROM [dbo].[tFacM] M2
								INNER JOIN [dbo].[tFacD] D2 ON [M2].[Branch] = [D2].[Branch]
										AND [M2].[intSerialNo] = [D2].[intSerialNo]
								WHERE  
									M2.[date] < @DateAfter
									AND M2.Status = 4
									AND (D2.GoodCode = M.GoodCode)
									AND M2.AccountYear = @AccountYear
									AND M2.Branch = @Branch
									AND Recursive = 0
				  	) AS BuyReturnPriceTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									[tFacM].[date] < @DateAfter
									AND [tFacM].Status = 1
									AND [tFacD].GoodCode = M.GoodCode
									AND [tFacM].AccountYear = @AccountYear
									AND [tFacM].Branch = @Branch
									AND [tFacM].Recursive = 0
				  	) AS BuyTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									[tFacM].[date] < @DateAfter
									AND [tFacM].Status = 4
									AND [tFacD].GoodCode = M.GoodCode
									AND [tFacM].AccountYear = @AccountYear
									AND [tFacM].Branch = @Branch
									AND [tFacM].Recursive = 0
				  	) AS BuyReturnTotal ,
				(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi) ,0)
							 FROM dbo.tInventory_Good
								WHERE  
									 tInventory_Good.GoodCode = M.GoodCode
									AND tInventory_Good.AccountYear = @AccountYear
									AND tInventory_Good.Branch = @Branch
				  	) AS FirstMojodiTotal ,
				(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi * tGood.BuyPrice) ,0)
							 FROM dbo.tInventory_Good
								INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
								WHERE  
									 tInventory_Good.GoodCode = M.GoodCode
									AND tInventory_Good.AccountYear = @AccountYear
									AND tInventory_Good.Branch = @Branch
				  	) AS FirstMojodiPriceTotal 
FROM 
(
        SELECT    
			ISNULL(k.GoodCode,dbo.tInventory_Good.GoodCode) AS GoodCode ,
			CAST(tinventory_good.FirstMojodi
			+ ISNULL(SUM(k.SumBuyAmount), 0)
			- ISNULL(SUM(k.SumFromStoreAmount), 0)
			+ ISNULL(SUM(k.SumtoStoreAmount), 0)
			- ISNULL(SUM(k.SumLossesAmount), 0)
			- ISNULL(SUM(k.SumBuyReturnAmount), 0) AS DECIMAL(20,3)) AS FirstDateMojodi
			
			FROM      (

				SELECT 
				CASE M.Status
				  WHEN 1 THEN SUM(D.Amount)  ELSE 0 END AS SumBuyAmount,
				CASE M.Status
				  WHEN 4 THEN SUM(D.Amount)  ELSE 0 END AS SumBuyReturnAmount,
				CASE M.Status
				  WHEN 6 THEN SUM(D.Amount)  ELSE 0 END AS SumFromStoreAmount,
				CASE M.Status
				  WHEN 7 THEN SUM(D.Amount)  ELSE 0  END AS SumtoStoreAmount,
				CASE M.Status
				  WHEN 3 THEN SUM(D.Amount)  ELSE 0  END AS SumLossesAmount,

				--CASE M.Status 
				--  WHEN 1 THEN D.FeeUnit * SUM(D.Amount)  ELSE 0  END AS BuyPriceTotal,
				--CASE M.Status
				--  WHEN 4 THEN D.FeeUnit * SUM(D.Amount) ELSE 0  END AS BuyReturnPriceTotal,
				--CASE M.Status
				--  WHEN 6 THEN D.FeeUnit * SUM(D.Amount) ELSE 0 END AS FromStorePriceTotal,
				--CASE M.Status
				--  WHEN 7 THEN D.FeeUnit * SUM(D.Amount) ELSE 0 END AS ToStorePriceTotal,
				--CASE M.Status
				--  WHEN 3 THEN D.FeeUnit * SUM(D.Amount) ELSE 0 END AS LossePriceTotal,
				D.GoodCode,
				D.intInventoryNo,
				M.AccountYear,
				M.Branch
			 FROM [dbo].[tFacM] M
				INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
						AND [M].[intSerialNo] = [D].[intSerialNo]
				INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode
						AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
						AND (tGood.Code = @GoodCode OR @GoodCode = 0)
				WHERE  
					M.[date] < @DateAfter
					AND M.Status IN ( 1, 3, 4, 6, 7 )
					AND (D.intInventoryNo = @inventory OR @inventory = 0)
					AND M.AccountYear = @AccountYear
					AND M.Branch = @Branch
					AND Recursive = 0
				GROUP BY  D.GoodCode,
					D.FeeUnit,
					M.Status,
					D.intInventoryNo,
					M.AccountYear,
					M.Branch

			) k
			
			RIGHT OUTER  JOIN tinventory_good ON k.goodcode = tinventory_good.goodcode
			AND (tinventory_good.Mojodi <> 0 OR @ZeroMojodi = 0)
			AND tinventory_good.inventoryno = k.intInventoryNo 
			AND tinventory_good.AccountYear = k.AccountYear
			AND tinventory_good.Branch = k.Branch
			INNER JOIN dbo.tGood ON dbo.tGood.Code = tinventory_good.GoodCode 
				AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12 
				AND (tGood.Code = @GoodCode OR @GoodCode = 0)
		  WHERE     tinventory_good.AccountYear = @AccountYear
					AND tinventory_good.Branch = @Branch
					AND (tinventory_good.InventoryNo = @inventory OR @inventory = 0) 
          GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    k.Branch,
                    tinventory_good.FirstMojodi,
                    tinventory_good.firstPrice,
                    tinventory_good.GoodCode ,
                    tGood.BuyPrice ,
                    tinventory_good.AccountYear ,
                    tinventory_good.Branch
        
)M
)M2
	ORDER BY goodcode


    RETURN
   END
--==========================================

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_InventoryAtomicRials_Report_InDate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_InventoryAtomicRials_Report_InDate]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



CREATE  PROCEDURE [dbo].[Get_InventoryAtomicRials_Report_InDate]
    (
      @SystemDate NVARCHAR(10) ,
      @SystemDay NVARCHAR(10) ,
      @SystemTime NVARCHAR(10) ,
      @Date1 NVARCHAR(50) ,
      @Date2 NVARCHAR(50) ,
      @AccountYear1 SMALLINT ,
      @Branch1 INT ,
      @Inventory1 INT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT 
    )
AS 

 BEGIN 

	SELECT DISTINCT X1.* ,
			CASE WHEN (X1.FirstMojodiTotal + X1.BuyTotal - X1.BuyReturnTotal) = 0 THEN X1.BuyPrice  
				WHEN  X1.Mojodi = 0 THEN 0
			ELSE 
			CAST(
			((X1.FirstMojodiTotal * X1.BuyPrice) + X1.BuyPriceTotal - X1.BuyReturnPriceTotal)  
			/ (X1.FirstMojodiTotal + X1.BuyTotal - X1.BuyReturnTotal) 
			AS BIGINT )END AS AvrageBuyPrice ,
		  @SystemDay + N' ' + @SystemDate + N' در ساعت' + @SystemTime AS SysDate ,
		(select [Description] FROM dbo.tInventory WHERE InventoryNo = @Inventory1) AS Inventory_Description ,
 		@Date1 AS DateBefore ,
		@Date2 AS DateAfter 
From (
  SELECT
		 X.* , tGood.Name , tGood.Code , tGood.BuyPrice ,
				(SELECT ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) 
							 FROM [dbo].[tFacM] M1
								INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
										AND [M1].[intSerialNo] = [D1].[intSerialNo]
								WHERE  
									    M1.[Date] <= @Date2
									AND M1.Status = 1
									AND D1.GoodCode = X.GoodCode
									AND M1.AccountYear = @AccountYear1
									AND M1.Branch = @Branch1
									AND Recursive = 0
				  	) AS BuyPriceTotal,
				(SELECT ISNULL(SUM(D2.FeeUnit * D2.Amount ) ,0)
							 FROM [dbo].[tFacM] M2
								INNER JOIN [dbo].[tFacD] D2 ON [M2].[Branch] = [D2].[Branch]
										AND [M2].[intSerialNo] = [D2].[intSerialNo]
								WHERE  
									    M2.[Date] <= @Date2
									AND M2.Status = 4
									AND (D2.GoodCode = X.GoodCode)
									AND M2.AccountYear = @AccountYear1
									AND M2.Branch = @Branch1
									AND Recursive = 0
				  	) AS BuyReturnPriceTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									    [tFacM].[date] <= @Date2
									AND [tFacM].Status = 1
									AND [tFacD].GoodCode = X.GoodCode
									AND [tFacM].AccountYear = @AccountYear1
									AND [tFacM].Branch = @Branch1
									AND [tFacM].Recursive = 0
				  	) AS BuyTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									    [tFacM].[date] <= @Date2
									AND [tFacM].Status = 4
									AND [tFacD].GoodCode = X.GoodCode
									AND [tFacM].AccountYear = @AccountYear1
									AND [tFacM].Branch = @Branch1
									AND [tFacM].Recursive = 0
				  	) AS BuyReturnTotal ,
				(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi) ,0)
							 FROM dbo.tInventory_Good
								WHERE  
									 tInventory_Good.GoodCode = X.GoodCode
									AND tInventory_Good.AccountYear = @AccountYear1
									AND tInventory_Good.Branch = @Branch1
				  	) AS FirstMojodiTotal --,
				--(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi * tGood.BuyPrice) ,0)
				--			 FROM dbo.tInventory_Good
				--				INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
				--  					AND tInventory_Good.GoodCode = X.GoodCode
				--					AND tInventory_Good.AccountYear = @AccountYear1
				--					AND tInventory_Good.Branch = @Branch1
				--  	) AS FirstMojodiPriceTotal 
		
 FROM (
         SELECT    
			T.FirstDateMojodi AS FirstMojodi ,
			T.FirstDateMojodiPrice AS FirstDateMojodiPrice ,
			CASE WHEN T.FirstDateMojodi = 0 THEN 0 
			ELSE CAST(T.FirstDateMojodiPrice * T.FirstDateMojodi  AS BIGINT )END  AS FirstDatePrice ,
			T.FirstDateMojodi
			+ ISNULL(SUM(k.TotalBuy), 0)
			- ISNULL(SUM(k.TotalHavaleh), 0)
			+ ISNULL(SUM(k.TotalResid), 0)
			- ISNULL(SUM(k.TotalLoss), 0)
			- ISNULL(SUM(k.TotalBuyReturn), 0) AS Mojodi,
			--CASE WHEN (T.FirstDateMojodi + ISNULL(SUM(k.TotalBuy), 0) - ISNULL(SUM(k.TotalBuyReturn), 0)) = 0 THEN K.BuyPrice  
			--ELSE 
			--CAST((
			--(T.FirstDateMojodiPrice * T.FirstDateMojodi)
			--+ ISNULL(SUM(k.TotalBuyPrice), 0) - ISNULL(SUM(k.TotalBuyReturnPrice), 0))  
			--/ (T.FirstDateMojodi + ISNULL(SUM(k.TotalBuy), 0)  - ISNULL(SUM(k.TotalBuyReturn), 0) ) 
			--AS BIGINT )END AS AvrageBuyPrice ,
			ISNULL(k.GoodCode,T.GoodCode) AS GoodCode ,
			
			ISNULL(SUM(k.TotalBuy), 0) AS TotalBuy ,
			ISNULL(SUM(k.TotalSell), 0) AS TotalSell,
			ISNULL(SUM(k.TotalLoss), 0) AS TotalLoss ,
			ISNULL(SUM(k.TotalBuyReturn), 0) AS TotalBuyReturn,
			ISNULL(SUM(k.TotalSellReturn), 0) AS TotalSellReturn,
			ISNULL(SUM(k.TotalHavaleh), 0) AS TotalHavaleh,
			ISNULL(SUM(k.TotalResid), 0) AS TotalResid,

			ISNULL(SUM(k.TotalBuyPrice), 0) AS TotalBuyPrice,
			ISNULL(SUM(k.TotalSellPrice), 0)AS TotalSellPrice ,
			ISNULL(SUM(k.TotalLossPrice), 0) AS TotalLossPrice,
			ISNULL(SUM(k.TotalBuyReturnPrice), 0) AS TotalBuyReturnPrice,
			ISNULL(SUM(k.TotalSellReturnPrice), 0) AS TotalSellReturnPrice,
			ISNULL(SUM(k.TotalHavalehPrice), 0) AS TotalHavalehPrice ,
			ISNULL(SUM(k.TotalResidPrice), 0) AS TotalResidPrice 

			FROM   
			   (
                 
				SELECT   
				DISTINCT [D].[GoodCode] ,D.intinventoryNo,
				M.AccountYear,
				M.Branch ,
				--------------------------------------------------------------------------------------------مقدار موجودي در  حالتهاي مختلف ----------------------------------------
				CASE WHEN [M].Status = 1 THEN SUM( [D].[Amount])ELSE 0 END AS TotalBuy ,
				CASE WHEN [M].Status = 2 THEN SUM( [D].[Amount])ELSE 0 END AS TotalSell ,
				CASE WHEN [M].Status = 3 THEN SUM([D].[Amount]) ELSE 0 END AS TotalLoss ,
				CASE WHEN [M].Status = 4 THEN SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturn ,
				CASE WHEN [M].Status = 5 THEN SUM([D].[Amount]) ELSE 0 END AS TotalSellReturn ,
				CASE WHEN [M].Status = 6 THEN SUM([D].[Amount]) ELSE 0 END AS TotalHavaleh ,
				CASE WHEN [M].Status = 7 THEN SUM([D].[Amount]) ELSE 0 END AS TotalResid ,

				CASE WHEN [M].Status = 1 THEN D.FeeUnit * SUM( [D].[Amount])ELSE 0 END AS TotalBuyPrice ,
				CASE WHEN [M].Status = 2 THEN D.FeeUnit * SUM( [D].[Amount])ELSE 0 END AS TotalSellPrice ,
				CASE WHEN [M].Status = 3 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalLossPrice ,
				CASE WHEN [M].Status = 4 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturnPrice ,
				CASE WHEN [M].Status = 5 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalSellReturnPrice ,
				CASE WHEN [M].Status = 6 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalHavalehPrice ,
				CASE WHEN [M].Status = 7 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalResidPrice ,

				tGood.BuyPrice
				-----------------------------------------------------------------------------------------------------------------------------------------------------------------
				FROM      [dbo].[tFacM] M
						INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
								AND [M].[intSerialNo] = [D].[intSerialNo]
						INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode 
								AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
				WHERE     [M].[Date] >= @Date1
						AND [M].[Date] <= @Date2
						AND [M].[AccountYear] = @AccountYear1
						AND [M].[Branch] = @Branch1
						AND [M].[Recursive] = 0
						AND [D].[intInventoryNo] = @Inventory1
				GROUP BY  [D].[GoodCode] ,
						[M].[Status] ,
						D.intInventoryNo ,
						D.FeeUnit ,
						M.AccountYear,
						M.Branch ,
						tGood.BuyPrice
			) K

			RIGHT OUTER  JOIN 
			(SELECT  GoodCode, FirstDateMojodi ,FirstDateMojodiPrice
                FROM    dbo.FnFirstDateMojodi(@Date1, @Inventory1, @Branch1,
                                              @AccountYear1 , @GoodLevel11 , @GoodLevel12 , 0 , 0) 
			)T 
			 ON k.goodcode = T.goodcode
          GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    k.Branch ,
                    T.FirstDateMojodi ,
                    T.FirstDateMojodiPrice ,
                    T.GoodCode ,
                    K.BuyPrice
    )X
    INNER JOIN tgood ON X.GoodCode = tGood.Code  AND   X.mojodi <> 0

)X1

	ORDER BY Name -- goodcode

    END


GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER   VIEW vw_PreviousFactors
AS
SELECT     dbo.tRepFacEditM.*, dbo.tShift.Description, dbo.tShift.LatinDescription, dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName
			,dbo.tServePlace.Description AS ServeDescription , dbo.tServePlace.LatinDescription  AS ServeLatinDescription , dbo.tTable.Name AS TableName
FROM         dbo.tShift RIGHT OUTER JOIN
                      dbo.tRepFacEditM INNER JOIN
                      dbo.tPer INNER JOIN
                      dbo.tUser 
		ON dbo.tPer.pPno = dbo.tUser.pPno --and dbo.tPer.Branch = dbo.tUser.Branch 
		ON dbo.tRepFacEditM.[User] = dbo.tUser.UID --and  dbo.tRepFacEditM.Branch = dbo.tUser.Branch
		ON dbo.tShift.Code = dbo.tRepFacEditM.ShiftNo --and dbo.tShift.Branch = dbo.tRepFacEditM.Branch
		INNER JOIN dbo.tServePlace ON dbo.tServePlace.intServePlace = dbo.tRepFacEditM.ServePlace  LEFT JOIN
		dbo.tTable ON dbo.tTable.Branch = dbo.tRepFacEditM.Branch AND dbo.tTable.No = dbo.tRepFacEditM.TableNo 
--Where		dbo.tRepFacEditM.Branch = dbo.Get_Current_Branch()
GO

