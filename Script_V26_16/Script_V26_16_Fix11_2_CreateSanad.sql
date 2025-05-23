
--حذف شماره پوز بانکی از شرح سند
--اضافه کردن شماره حساب به شرح سند

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_SaleSummaryCustom]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_SaleSummaryCustom]
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE  PROCEDURE [dbo].[Get_SaleSummaryCustom]
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

--SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت کارت' + N' بانک ' + MIN(TAB.nvcBankName) + N' شماره ' + MIN(TPP.nvcAccountNo) + N' درتاریخ ' +  TF.[Date]  AS [Name],  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TPP.AccountId) AS Tafsili FROM 
SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت کارت' + N' درتاریخ ' +  TF.[Date]  AS [Name],  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein ,0 AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tFacCard TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
--INNER JOIN dbo.tblPub_Pos TPP on TPP.PosId = TFC.PosId AND Tf.StationID = TPP.StationId
--INNER JOIN dbo.tblAcc_Bank TAB ON TAB.tintBank = TPP.intBank
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
					--AND (InCharge = NULL OR (InCharge > 0 AND FacPayment = 1)) 
     		        AND Balance = 0
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
