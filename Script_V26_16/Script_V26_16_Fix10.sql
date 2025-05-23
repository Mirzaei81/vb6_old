

--Script_V26_16_Fix10
--اضافه شدن مرکز هزینه به فرم سود وزیان حسابداری
--تولید سند حسابداری یکپارچه از داخل اسکریپت
--اضافه شدن بدهی فروش به مشتریان (اگر سند تولید نشده)در تولید سند حسابداری
--اضافه شدن فیلد تفضیلی به پارتیشن ها برای محاسبه عوارض و مالیات و سایر افزایش ها

-- 93/10/19

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
          10
        )
GO
-- Date 930920

IF COL_LENGTH('tPartitions','Tafsili') IS NULL
BEGIN
	ALTER TABLE dbo.tPartitions
	ADD Tafsili INT NULL
END

GO
IF COL_LENGTH('tStations','PartitionId') IS NULL
BEGIN
	ALTER TABLE dbo.tStations
	ADD PartitionId INT NOT NULL DEFAULT(1)
END

GO


UPDATE dbo.tStations SET PartitionId = 1 WHERE PartitionId IS NULL 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_SaleSummaryCustom]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_SaleSummaryCustom]
GO

CREATE PROCEDURE [dbo].[Get_SaleSummaryCustom]
(
@Branch INT ,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT = 0
)

 AS
BEGIN

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت صندوق' + '  ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date] AS [Name] ,  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
 INNER JOIN dbo.tUser TU ON TU.UID = TF.[User] AND TU.Branch = TF.Branch
 INNER JOIN dbo.tPer TP ON TP.pPno = TU.pPno AND TP.Branch = TU.Branch  
 INNER JOIN dbo.tFacCash TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 7
GROUP BY TF.[User] , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت کارت' + N' بانک ' + MIN(TPP.nvcBankName) + N' شماره ' + MIN(TPP.NvcPosNo) + N' درتاریخ ' +  TF.[Date]  AS [Name],  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TPP.AccountId) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
 INNER JOIN dbo.tFacCard TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
INNER JOIN dbo.tblPub_Pos TPP on TPP.PosId = TFC.PosId
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 22
GROUP BY TFC.PosId , TF.[Date]


UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت پیک' + ' ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date]  AS [Name],  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)
     		        AND InCharge > 0 
     		        AND Balance = 0 and FacPayment = 0) TF
 INNER JOIN dbo.tPer TP ON TP.pPno = TF.InCharge AND TP.Branch = TF.Branch  and Job = 3
 INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 6
GROUP BY TP.pPno , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت گارسون' + ' ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date]  AS [Name],  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)
     		        AND InCharge > 0 
     		        AND Balance = 0 and FacPayment = 0) TF
 INNER JOIN dbo.tPer TP ON TP.pPno = TF.InCharge AND TP.Branch = TF.Branch  and Job = 9
 INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 6
GROUP BY TP.pPno , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' بدهکاری مشتری' + ' ' + MIN(TC.Name + '' + TC.Family) + N' در تاریخ ' + TF.[Date]  AS [Name] ,  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TC.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
					AND (InCharge IS NULL OR (InCharge > 0 AND FacPayment = 1)) 
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer > 0 AND Credit > 0)
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 7
GROUP BY TC.Code , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' تخفیفات فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , SUM(TF.DiscountTotal) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  2
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'  فروش ' + ' ' + MIN(Ts.[Description]) + N' در تاریخ  ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(Tf.Amount * Tf.FeeUnit) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TS.Tafsili) AS Tafsili FROM 
(SELECT tFacM.* , Amount , FeeUnit , intInventoryNo FROM dbo.tFacM INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tInventory TS ON TS.InventoryNo = TF.intInventoryNo
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  1
GROUP BY TS.InventoryNo , TF.[Date]

--SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' فروش ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice + TF.DiscountTotal - TF.CarryFeeTotal - PackingTotal - ServiceTotal - TaxTotal - DutyTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
--(SELECT * FROM dbo.tFacM
--                    where [Date] >= @DateBefore
--                    AND [Date] <= @DateAfter
--                    AND Recursive = 0
--                    AND Status = 2
--                    AND transferAccounting = 0
--                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
--INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
--INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
--INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  1
--GROUP BY TP.PartitionID , TF.[Date]

--UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' بستانکاری مشتری' + ' ' + MIN(TC.Name + '' + TC.Family) + N' در تاریخ ' + TF.[Date]  AS [Name] ,  0 AS SumBedehKar , SUM(ISNULL(TFC.intAmount,0)+ISNULL(TFCA.intAmount,0)) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TC.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 2
----                    AND transferAccounting = 0
----     		        AND (InCharge = NULL OR (InCharge > 0 AND FacPayment = 1)) 
----     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
----INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer > 0 AND Credit > 0)
----LEFT JOIN dbo.tFacCash TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
----LEFT JOIN dbo.tFacCard TFCA ON TFCA.Branch = TF.Branch AND TFCA.intSerialNo = TF.intSerialNo
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TC.Code , TF.[Date]
----HAVING SUM(ISNULL(TFC.intAmount,0)+ISNULL(TFCA.intAmount,0)) > 0

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'موجودي مواد و کالا' AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM INNER JOIN 
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 16
----GROUP BY TF.[Date]


----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , MIN(TS.Name + ' ' + TS.Family + ' ' + TS.WorkName) AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , TS.Code AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TS.Code , TF.[Date]


----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'برگشت از خريد' AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 4
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 17
----GROUP BY TF.[Date]

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , MIN(TS.Name + ' ' + TS.Family + ' ' + TS.WorkName) AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , TS.Code AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TS.Code , TF.[Date]

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'برگشت از فروش' AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 5
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 18
----GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' عوارض فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.DutyTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  24
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' مالیات فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(TF.TaxTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  26
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' درآمد سرویس   ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.ServiceTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  38
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'  درآمد بسته بندی   ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.PackingTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  3
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' کرایه حمل فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.CarryFeeTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  4
GROUP BY TF.[Date]


END

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER    PROCEDURE [dbo].[Get_SaleSummary]
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
                    dbo.tInventory.Tafsili ,
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



SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO


ALTER  Function [dbo].Fn_SoodZian

(
  @DateBefore INT  ,
  @DateAfter INT  ,
  @AccountYear SMALLINT ,
  @Branch INT ,
  @MarkazHazineh INT 
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

 TotalLosses BIGINT ,
 TotalHoghough BIGINT ,
 TotalHazine BIGINT ,
 TotalHazineMali BIGINT ,
 TotalHazineTozie BIGINT 
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

	DECLARE @TotalLosses BIGINT
	DECLARE @TotalHoghough BIGINT
	DECLARE @TotalHazine BIGINT
	DECLARE @TotalHazineMali BIGINT
	DECLARE @TotalHazineTozie BIGINT
	


		Select @TotalSellAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalSellReturnAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalFirstPrice = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalBuyAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalBuyReturnAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalSaleDiscount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalBuyDiscount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29)
		AND TafsiliId = @MarkazHazineh

		Select @TotalLosses = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHazine = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHazineMali = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 36  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHazineTozie = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 37  )
		AND MoeinId <> (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32) --Losses  moein code calculated in totallosses
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalCareeFee = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalPacking = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3)
		AND TafsiliId = @MarkazHazineh
		
		Select @Totalservice = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 38  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 38)
		AND TafsiliId = @MarkazHazineh
		
		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,  Totalservice ,
			 TotalLosses  , TotalHoghough  , TotalHazine , TotalHazineMali , TotalHazineTozie )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  , @Totalservice ,
			 @TotalLosses  , @TotalHoghough  , @TotalHazine , @TotalHazineMali , @TotalHazineTozie)
		            


RETURN 


End

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Get_TarazSoodZian]
    (
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT ,
      @MarkazHazineh INT 
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
       ISNULL(TotalLosses , 0) AS TotalLosses ,
       ISNULL(TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(TotalHazine , 0) AS TotalHazine ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
       
	FROM DBO.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch , @MarkazHazineh )
--===============================================


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Rep_TarazSoodZian]
    (
      @SystemDate NVARCHAR(20) ,
      @SystemDay NVARCHAR(20) ,
      @SystemTime NVARCHAR(20) ,
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT ,
      @MarkazHazineh INT ,
      @MojodiPrice BIGINT 
    )
AS 

    DECLARE @TimeTitle NVARCHAR(10)      
    SET @TimeTitle = N' ساعت : '   

SELECT @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS SysDay  ,
		SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,1 ,4) + N'/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,5,2) + N'/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,7,2) AS FromDate ,
		SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,1 ,4) + N'/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,5,2) + N'/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,7,2) AS ToDate ,
		@MojodiPrice AS MojodiPrice ,
		ISNULL(TotalSellAmount , 0) AS TotalSellAmount ,
		ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
		ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
		ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
		ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
		ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
		ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
		ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
		ISNULL(TotalPacking , 0) AS TotalPacking ,
		ISNULL(TotalService , 0) AS TotalService ,
		ISNULL(TotalLosses , 0) AS TotalLosses ,
		ISNULL(TotalHoghough , 0) AS TotalHoghough ,
		ISNULL(TotalHazine , 0) AS TotalHazine  ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
	FROM dbo.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch , @MarkazHazineh)
--===============================================

GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER    PROCEDURE [dbo].[Get_SaleSummary_Added]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
 
 SELECT 
        tvw.[Date] ,
        Branch ,
        SUM(DiscountTotal) AS Sumdiscount ,
        SUM(PackingTotal) AS SumPacking ,
        SUM(CarryFeeTotal) AS SumCarryFee ,
        SUM(ServiceTotal) AS SumService ,
        SUM(DutyTotal) AS DutyTotal ,
        SUM(TaxTotal) AS TaxTotal ,
        0 AS tafsili
 FROM   ( SELECT DISTINCT dbo.tFacM.Branch ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacM.[Time] ,
                    dbo.tFacM.[User] ,
                    CarryFeeTotal ,
                    DiscountTotal ,
                    StationID ,
                    ServiceTotal ,
                    PackingTotal ,
                    TaxTotal ,
                    DutyTotal ,
                    FacPayment ,
                    Balance ,
                    --( tfacd.Amount * tfacd.Feeunit ) AS SumPrice
                    dbo.tFacM.SumPrice
          FROM      dbo.tFacM
                    --INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                    --                    AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch
 ORDER BY tvw.[Date] 
 
 
END

GO


