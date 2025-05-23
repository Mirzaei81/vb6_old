
--Script_V26_16_Fix14_TarazSoodZian_Sale.sql
--محاسبه تراز کلی حاصل از فروش
--94/08/27


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    PROCEDURE [dbo].[Get_TarazSoodZian_Sale]
    (
      @DateBefore NVARCHAR(8)  ,
      @DateAfter NVARCHAR(8)  ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS 


SELECT ISNULL(TotalSellAmount , 0) - ISNULL(TotalCareeFee , 0) - ISNULL(TotalService , 0)
		- ISNULL(TotalPacking , 0) - ISNULL(TotalTax , 0) AS TotalSellAmount ,
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
