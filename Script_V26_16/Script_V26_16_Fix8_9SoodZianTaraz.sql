
--Script_V26_16_Fix8_9_SoodZianTaraz.sql
--سود و زیان حسابداری از سود و زیان فروش جدا شده
--

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fn_SoodZian]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Fn_SoodZian]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fn_SoodZian_Sale]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Fn_SoodZian_Sale]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE  Function [dbo].Fn_SoodZian

(
  @DateBefore INT  ,
  @DateAfter INT  ,
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

		Select @TotalSellReturnAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17)

		Select @TotalFirstPrice = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35)

		Select @TotalBuyAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16)

		Select @TotalBuyReturnAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18)

		Select @TotalSaleDiscount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2)

		Select @TotalBuyDiscount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29)

		Select @TotalLosses = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32)

		Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate

		Select @TotalHazine = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate

		Select @TotalHazineMali = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 36  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate

		Select @TotalHazineTozie = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 37  )
		AND MoeinId <> (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32) --Losses  moein code calculated in totallosses

		Select @TotalCareeFee = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4)

		Select @TotalPacking = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3)

		Select @Totalservice = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 38  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 38)

		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,  Totalservice ,
			 TotalLosses  , TotalHoghough  , TotalHazine , TotalHazineMali , TotalHazineTozie )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  , @Totalservice ,
			 @TotalLosses  , @TotalHoghough  , @TotalHazine , @TotalHazineMali , @TotalHazineTozie)
		            


RETURN 


End




GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
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



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_TarazSoodZian]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_TarazSoodZian]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROCEDURE [dbo].[Get_TarazSoodZian]
    (
      @DateBefore INT  ,
      @DateAfter INT  ,
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
       ISNULL(TotalLosses , 0) AS TotalLosses ,
       ISNULL(TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(TotalHazine , 0) AS TotalHazine ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
       
	FROM DBO.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch )
--===============================================


GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_TarazSoodZian_Sale]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_TarazSoodZian_Sale]
GO

SET QUOTED_IDENTIFIER ON 
GO
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
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

