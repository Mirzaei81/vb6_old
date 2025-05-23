

--حذف اضافه کردن کالاها به ایستگاهها در انتقال و انبار گردانی
-- 94/01/22

ALTER  proc Transport_tblTotal_tGood_By_Prams 
(
	@Level1 int ,
	@strSelectedLevels nvarchar(4000) , 
	@InventoryNo int ,
	@Branch int ,
	@AccountYear SMALLINT,
	@CheckNotZeroMojodi BIT,
	@CheckFirstMojodi	BIT,
	@CountingNo	INT,
	@ToOtherAccountYear SMALLINT
)
	
as
BEGIN TRAN

	DELETE tInventory_Good
	FROM
	(
		SELECT vw_Good.* , tInventory_Good.* 
		
		FROM 
			[dbo].[vw_Good] 
			Inner Join  
			dbo.tInventory_Good ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode 
		WHERE 
			(LEVEL1 = @Level1 OR @Level1=-1)
			And (InventoryNo = @InventoryNo OR @InventoryNo=-1)
			And (Branch = @Branch OR @Branch=-1)
			And (AccountYear = @ToOtherAccountYear OR @ToOtherAccountYear=-1)
			AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels=N'')
			AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
			AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
	)AS T
	WHERE T.GoodCode=tInventory_Good.GoodCode 
		AND T.InventoryNo=tInventory_Good.InventoryNo
		AND T.Branch=tInventory_Good.Branch
		AND T.AccountYear=tInventory_Good.AccountYear

if @@Error <> 0 
	Goto ErrHandler

	INSERT INTO tInventory_Good
		(
		      InventoryNo, Branch, GoodCode, FirstMojodi, Mojodi, MojodiControl, OrderPoint, MinValue, MaxValue, [Date], [Time], BuyAmount, SaleAmount, 
                      LossAmount, BuyReturnAmount, SaleReturnAmount, FromStoreAmount, toStoreAmount, AccountYear, Counting1, 
                      Counting2, Counting3, CountDifference
		)

		SELECT     InventoryNo, Branch, GoodCode,CASE @CountingNo 
								WHEN 0 THEN CAST(ISNULL(T.Mojodi  ,0) AS DECIMAL(20,3))
								WHEN 1 THEN ISNULL(T.Counting1,0)
								WHEN 2 THEN ISNULL(T.Counting2,0)
								WHEN 3 THEN ISNULL(T.Counting3,0)
								ELSE ISNULL(T.Mojodi,0)
								END
							AS FirstMojodi, 0, 0, 0, MinValue, MaxValue, [Date], [Time], 0, 0, 
	                      0, 0, 0, 0, 0, @ToOtherAccountYear, 0, 
	                      0, 0, 0
		FROM
		(
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
		)AS T
	
if @@Error <> 0 
	Goto ErrHandler

--	Insert into dbo.tStation_Inventory_Good ( branch ,InventoryNo, AccountYear ,StationID,  GoodCode , Active)
	
--	select Branch ,InventoryNo ,@ToOtherAccountYear ,StationID ,GoodCode ,Active 
--		From tStation_Inventory_Good 
--	        Where   inventoryno = @inventoryno and Branch = @Branch and AccountYear = @AccountYear

--if @@Error <> 0 
--	Goto ErrHandler


Commit Tran
Return

ErrHandler:
RollBack Tran
Return



GO


