
--ScriptV26_16_Fix17_950815_Get_tblTotal_tGood_By_Prams.sql
--950815


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER    proc Get_tblTotal_tGood_By_Prams 
(
	@Level1 int ,
	@strSelectedLevels nvarchar(4000) ,
	@Type		INT, 
	@InventoryNo int ,
	@Branch int ,
	@AccountYear SMALLINT,
	@CheckNotZeroMojodi INT,
	@CheckFirstMojodi	INT,
	@CheckOrder		INT ,
	@Flag	INT =Null ,
	@SortItem	INT = NULL
)
	
as

IF @Flag IS NULL
	SET @Flag = 0

IF @SortItem IS NULL
	SET @SortItem = 1

If @Flag = 0 
bEGIN


	SELECT vw_Good.* , tInventory_Good.* ,vw_Good.AvgBuyPrice AS AverageBuyPrice,vw_Good.SellPrice AS LastSellPrice
	
	FROM 
		[dbo].[vw_Good] 
		Inner Join  
		dbo.tInventory_Good ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode 
		
	WHERE 
		(LEVEL1 = @Level1 OR @Level1=-1)
		And InventoryNo = @InventoryNo 
		And Branch = @Branch 
		And AccountYear = @AccountYear 
		AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels='')
		AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
		AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
		AND (GoodType=@Type OR @Type=-1)
		AND ((dbo.tInventory_Good.OrderPoint >= dbo.tInventory_Good.Mojodi) OR (-1=CASE @CheckOrder WHEN 1 THEN 0 ELSE -1 END))
	Order By
		Case @SortItem  When 1 Then  GoodCode
		     		When 2 Then Barcode
		     		When 3 Then [Name]
		     		When 4 Then Unit
		     		When 5 Then Mojodi
		     		When 6 Then Sellprice
		     		When 7 Then BuyPrice
		     		when 8 then Counting1
	
			End

 			
END

ELSE

BEGIN
	Select Y.* , IsNull(LastSellPrice , Y.SellPrice) As LastSellPrice
	From (
	Select t.* ,  Cast(ISNULL(AverageBuyPrice ,t.BuyPrice) AS int) As AverageBuyPrice
	From (
	SELECT vw_Good.* , tInventory_Good.* 
	
	FROM 
		[dbo].[vw_Good] 
		Inner Join  
		dbo.tInventory_Good ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode 
		
	WHERE 
		(LEVEL1 = @Level1 OR @Level1=-1)
		And (InventoryNo = @InventoryNo OR @InventoryNo=-1)
		And (Branch = @Branch OR @Branch=-1)
		And (AccountYear = @AccountYear OR @AccountYear=-1)
		AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels='')
		AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
		AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
		AND (GoodType=1 OR GoodType = 3) --OR @Type=-1
		AND ((dbo.tInventory_Good.OrderPoint >= dbo.tInventory_Good.Mojodi) OR (-1=CASE @CheckOrder WHEN 1 THEN 0 ELSE -1 END))
		)t
	
        Left Outer Join  
	(Select IsNull(Sum(FeeUnit * Amount) ,0)/ISNULL(Sum(Amount),1) As AverageBuyPrice , tFacd.GoodCode
	From tFacD
	inner join tfacM On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch
 
	Where tfacm.Status = 1 and Recursive = 0 And  tfacM.AccountYear = @AccountYear
	Group By GoodCode  )X
	On X.GoodCode = t.Code
 	)Y
       Left Outer Join  
	(Select Top 1 FeeUnit As LastSellPrice , tFacd.GoodCode
	From tFacD
	inner join tfacM On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch
 
	Where (tfacm.Status = 2 )  and Recursive = 0 And  tfacM.AccountYear = @AccountYear
	Order By [Date] Desc	)W
	On W.GoodCode = Y.Code

	Order By
			Case @SortItem  When 1 Then  Y.GoodCode
			     		When 2 Then Barcode
			     		When 3 Then [Name]
			     		When 4 Then Unit
			     		When 5 Then Mojodi
			     		When 6 Then Sellprice
			     		When 7 Then BuyPrice
			     		when 8 then Counting1
		
				End


			
END

GO
