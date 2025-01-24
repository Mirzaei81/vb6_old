
--ScriptV26_16_Fix16_AveragealculateBuyPrice.sql
--95/05/16

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[AverageCalculateBuyPrice] 
(@GoodCode as int ,  @DateAfter Nvarchar(10) , @DateBefore Nvarchar(10) , @Flag INT) 
 AS
DECLARE @AccountYear SMALLINT
SET @AccountYear = CAST('13' + SUBSTRING(@DateAfter ,1,2) AS SMALLINT)
--PRINT @AccountYear
DECLARE @Branch INT 
SET @Branch = dbo.Get_Current_Branch()
--PRINT @Branch
Declare @FeeUnit INT

IF @Flag = 0 
BEGIN

DECLARE @BuyTotal FLOAT
DECLARE @BuyPriceTotal BIGINT
DECLARE @BuyReturnTotal FLOAT
DECLARE @BuyReturnPriceTotal BIGINT
DECLARE @FirstMojodiTotal FLOAT 
 
	--SELECT @BuyTotal = CASE WHEN M1.Status = 1 THEN  ISNULL(SUM(D1.Amount) ,0) ELSE 0 END ,
	--		@BuyPriceTotal = CASE WHEN M1.Status = 1 THEN  ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) ELSE 0 END,
	--		@BuyReturnTotal = CASE WHEN M1.Status = 4 THEN  ISNULL(SUM(D1.Amount) ,0) ELSE 0 END,
	--		@BuyReturnPriceTotal = CASE WHEN M1.Status = 4 THEN  ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) ELSE 0 END
	--			 FROM [dbo].[tFacM] M1
	--				INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
	--						AND [M1].[intSerialNo] = [D1].[intSerialNo]
	--				WHERE  
	--					    M1.[Date] >= @DateAfter
	--					AND M1.[Date] <= @DateBefore
	--					AND (M1.Status = 1 OR M1.Status = 4)
	--					AND D1.GoodCode = @GoodCode
	--					AND M1.AccountYear = @AccountYear
	--					AND M1.Branch = @Branch
	--					AND Recursive = 0
	--					GROUP BY Status , GoodCode

	SELECT @BuyTotal = ISNULL(SUM(D1.Amount) ,0) ,
			@BuyPriceTotal = ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0)
				 FROM [dbo].[tFacM] M1
					INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
							AND [M1].[intSerialNo] = [D1].[intSerialNo]
					WHERE  
						    M1.[Date] >= @DateAfter
						AND M1.[Date] <= @DateAfter
						AND (M1.Status = 1 )
						AND D1.GoodCode = @GoodCode
						AND M1.AccountYear = @AccountYear
						AND M1.Branch = @Branch
						AND Recursive = 0
						GROUP BY Status , GoodCode
	SELECT 	@BuyReturnTotal = ISNULL(SUM(D1.Amount) ,0) ,
			@BuyReturnPriceTotal = ISNULL(SUM(D1.FeeUnit * D1.Amount ) , 0)
				 FROM [dbo].[tFacM] M1
					INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
							AND [M1].[intSerialNo] = [D1].[intSerialNo]
					WHERE  
						    M1.[Date] >= @DateAfter
						AND M1.[Date] <= @DateAfter
						AND (M1.Status = 4)
						AND D1.GoodCode = @GoodCode
						AND M1.AccountYear = @AccountYear
						AND M1.Branch = @Branch
						AND Recursive = 0
						GROUP BY Status , GoodCode

	SELECT @FirstMojodiTotal = ISNULL(SUM(tInventory_Good.FirstMojodi) ,0) 
				 FROM dbo.tInventory_Good
					WHERE  
						 tInventory_Good.GoodCode = @GoodCode
						AND tInventory_Good.AccountYear = @AccountYear
						AND tInventory_Good.Branch = @Branch



--PRINT @BuyTotal 
--PRINT @BuyPriceTotal 
--PRINT @BuyReturnTotal 
--PRINT @BuyReturnPriceTotal 
--PRINT @FirstMojodiTotal  


DECLARE @BuyPrice INT 
SET @BuyPrice = (Select IsNull(BuyPrice,1) From tGood 	Where  tGood.Code = @Goodcode )
SET @FeeUnit = CASE WHEN (ISNULL(@FirstMojodiTotal,0) + ISNULL(@BuyTotal,0) - ISNULL(@BuyReturnTotal ,0) ) = 0 THEN @BuyPrice  
			ELSE 
			CAST(
			((ISNULL(@FirstMojodiTotal,0) * @BuyPrice) + ISNULL(@BuyPriceTotal,0) - ISNULL(@BuyReturnPriceTotal ,0))  
			/ (ISNULL(@FirstMojodiTotal,0) + ISNULL(@BuyTotal,0) - ISNULL(@BuyReturnTotal ,0) ) 
			AS BIGINT ) END

		--UPDATE dbo.tInventory_Good
		--SET BuyPriceAverage = @FeeUnit
		--WHERE AccountYear = @AccountYear
		--AND Branch = @Branch AND GoodCode = @GoodCode

	--SET @FeeUnit = (Select (IsNull(Sum(FeeUnit * Amount) ,0)/ISNULL(Sum(Amount),1)) From tFacM inner join tfacd On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch
	--			Where tfacm.Status = 1 and Recursive = 0 And Date >= @DateAfter And Date <= @DateBefore And tfacd.GoodCode = @Goodcode ) 
	--IF @FeeUnit = 0 
	--	SET @FeeUnit = (Select IsNull(BuyPrice,1) From tGood 	Where  tGood.Code = @Goodcode ) 
END
ELSE
	SET @FeeUnit = (Select IsNull(BuyPrice,1) From tGood 	Where  tGood.Code = @Goodcode ) 


Select @FeeUnit As AverageBuyPrice


GO
