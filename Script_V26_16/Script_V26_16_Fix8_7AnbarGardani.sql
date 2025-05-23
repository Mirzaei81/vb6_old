
--فقط در ورژن های الماس 
--Script_V26_16_Fix8_AnbarGardani.sql
--انبارگردانی و صدور سند کسری و اضافی انبار به صورت اتوماتیک
--امکان صدور حواله بدون مشخص کردن مقصد . مانند صدور حواله مصرف و یا ضایعات که نیازی به رسید انبار ندارد
--امکان صدور رسید به انبار مستقل و نه از طریق حواله . مانند سند اضافات انبار
--93/08/08


ALTER   VIEW dbo.vw_Good
AS
SELECT DISTINCT 
		dbo.tSupplier.WorkName AS CompDes, dbo.tGood.Code, dbo.tGood.Level1, dbo.tGood.Level2, dbo.tGood.Name, 

		dbo.tGood.LatinName, 
		dbo.tGood.NamePrn, dbo.tGood.LatinNamePrn, dbo.tGood.BarCode, dbo.tGood.Unit, dbo.tGood.Model, dbo.tGood.Weight, 

		dbo.tGood.NumberOfUnit, 
		dbo.tGood.ProductCompany, dbo.tGood.SellPrice, dbo.tGood.BuyPrice, dbo.tGood.BtnAscDefault, dbo.tUnitGood.Description AS 

		UnitDescription, 
		dbo.tGoodModel.Description, dbo.tGood.GoodType, dbo.tGood.BtnTz1No, dbo.tGood.TechnicalNo, dbo.tGoodType.Description AS 

		TypeDescription, 
		dbo.tGoodLevel1.Description AS Level1Description, dbo.tGoodLevel1.LatinDescription AS Level1LatinDescription, 
		dbo.tGoodLevel2.Description AS Level2Description, dbo.tGoodLevel2.LatinDescription AS Level2LatinDescription, 

		dbo.tGood.SellPrice2, 
		dbo.tGood.SellPrice3, dbo.tGood.Discount, dbo.tGood.MainType ,dbo.tGood.SellPrice4,dbo.tGood.SellPrice5 ,dbo.tGood.SellPrice6 , 

	  dbo.tGood.PicturePath , dbo.tGood.DutyBuy , dbo.tGood.DutySale , dbo.tGood.TaxBuy , dbo.tGood.TaxSale , dbo.tGood.CategoryShow , dbo.tGood.nvcDescription
	, GoodNamePrn2 , GoodNamePrn3 , AvgBuyPrice
FROM         dbo.tGood INNER JOIN
              dbo.tGoodModel ON dbo.tGood.Model = dbo.tGoodModel.Code INNER JOIN
              dbo.tSupplier ON dbo.tGood.ProductCompany = dbo.tSupplier.Code  INNER JOIN
              dbo.tUnitGood ON dbo.tGood.Unit = dbo.tUnitGood.Code INNER JOIN
              dbo.tGoodType ON dbo.tGood.GoodType = dbo.tGoodType.Code INNER JOIN
              dbo.tGoodLevel1 ON dbo.tGood.Level1 = dbo.tGoodLevel1.Code INNER JOIN
              dbo.tGoodLevel2 ON dbo.tGood.Level2 = dbo.tGoodLevel2.Code AND dbo.tGoodLevel1.Code = dbo.tGoodLevel2.Level1Code




GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO



ALTER   proc Get_tblTotal_tGood_By_Prams 
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
		And (InventoryNo = @InventoryNo OR @InventoryNo=-1)
		And (Branch = @Branch OR @Branch=-1)
		And (AccountYear = @AccountYear OR @AccountYear = -1)
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



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER PROCEDURE [dbo].[InsertFactorDetail]  (
	 @DetailsString Nvarchar(4000) ,
	 @intSerialNo bigint ,
	 @intserialNo2 bigint ,
	 @Customer Bigint ,
	 @Branch int = Null
	
) 
As


if @Branch is null
    select @Branch = branch from tInventory where inventoryNo=(SELECT Top 1  intInventoryNo FROM Split(@DetailsString))



Declare @Status Int 

Set @Status = (Select Status from tfacm Where intserialno = @intSerialNo and Branch = @Branch)


     INSERT INTO tFacD
	(
	    
		intRow,
		Amount ,
		GoodCode  ,
		FeeUnit ,
		Discount ,
		Rate ,
		ChairName ,
		[ExpireDate] ,
		intInventoryNo ,
		DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
		ServePlace ,
		DifferencesCodes , 
		DifferencesDescription ,
		intSerialNo , 
		Branch 
	)
	     SELECT
		
		tmpTable.Row ,
		tmpTable.Amount ,
		tmpTable.GoodCode ,
		tmpTable.FeeUnit ,
		tmpTable.Discount ,
		tmpTable.Rate ,
		tmpTable.ChairName ,
		tmpTable.[ExpireDate],
		tmpTable.intInventoryNo ,
		tmpTable.DestInventoryNo ,--Because Has a Relation and Can not insert for  Another Branch
		tmpTable.ServePlace ,
		tmpTable.DifferencesCode ,
		tmpTable.DifferencesDescription ,
		@intSerialNo , 
		@Branch 	
	
	FROM (SELECT * FROM Split(@DetailsString )) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode

	DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

	If @Status = 6 AND @DestinventoryNo > 0
	Begin
	
	declare @destbranch INT
	select @destbranch=@Branch --branch from tInventory where inventoryNo=(SELECT Top 1  DestInventoryNo FROM Split(@DetailsString))
	  	   begin
			 INSERT INTO tFacD
			(
			    
				intRow,
				Amount ,
				GoodCode  ,
				FeeUnit ,
				Discount ,
				Rate ,
				ChairName ,
				[ExpireDate] ,
				intInventoryNo ,
				DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
				ServePlace ,
				DifferencesCodes , 
				DifferencesDescription ,
				intSerialNo , 
				Branch
			)
				 SELECT
				
				tmpTable.Row ,
				tmpTable.Amount ,
				tmpTable.GoodCode ,
				tmpTable.FeeUnit ,
				tmpTable.Discount ,
				tmpTable.Rate ,
				tmpTable.ChairName ,
				tmpTable.[ExpireDate],
				tmpTable.DestInventoryNo ,
				tmpTable.intInventoryNo ,--Because Has a Relation and Can not insert for  Another Branch
				tmpTable.ServePlace ,
				tmpTable.DifferencesCode ,
				tmpTable.DifferencesDescription ,
				@intSerialNo2 , 
				@DestBranch --dbo.Get_Current_Branch()
		
		
			FROM (SELECT * FROM Split(@DetailsString )) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode
	
		   end
	end
	

Update tFacD Set Amount = 1 where amount = 0 and intserialno = @intSerialNo and Branch = @Branch
--Update tFacD Set DestInventoryNo = Null Where intserialno = @intSerialNo and Branch = dbo.Get_Current_Branch()

GO


--Script_V26_16_Fix5_Added2
--DestBranch is CurrentBranch For Replication
--اصلاح گردش کالا در انبار براي شعبات
--اصلاح بروزرساني در انبار شعبات

ALTER PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

            @Status INT ,      
            @Owner INT ,      
            @Customer INT ,      
            @DiscountTotal FLOAT ,      
            @CarryFeeTotal FLOAT ,      
            @Recursive INT ,      
            @InCharge INT ,      
            @FacPayment BIT ,      
            @OrderType INT ,      
            @StationId INT ,      
            @ServiceTotal FLOAT ,      
            @PackingTotal FLOAT ,      
            @TableNo INT ,      
            @User INT ,      
            @Date NVARCHAR(50) ,      
            @DetailsString nText,      
            @ds nText = '',      
            @Balance BIT ,      
            @AccountYear smallint = null  ,       
            @NvcDescription Nvarchar(150) = Null ,      
            @HavaleNo int = Null  ,      
            @TempAddress Nvarchar(255) = '',  
			@GuestNo INT,    
            @lastFacMNo INT OUT  ,
		    @Person INT = NULL     
             )      

AS      

Declare @intserialNo int      
Declare @intserialNo2 int      
--Declare @intserialNo3 Bigint    

SET @intserialNo = 0        
SET @intserialNo2   = 0      
--SET @intserialNo3   = 0      

DECLARE @No1  INT     
DECLARE @No2  INT     
--DECLARE @No3  INT     

DECLARE @SumPrice  float      
Set @SumPrice = 0      

DECLARE @proper_time nvarchar(5)      

IF @AccountYear IS  Null
    Set @AccountYear = dbo.get_AccountYear() 
    
IF  @Owner = 0      
    SET @Owner = NULL      

IF  @TableNo < 1      
    SET @TableNo = NULL      

IF  @Incharge < 1      
    SET @Incharge = NULL      

IF  @Customer=0      
    SET @Customer = NULL      

BEGIN TRAN      

    DECLARE @MasterServePlace INT      
    DECLARE @newtime nvarchar(5)      
    select @newtime=dbo.setTimeFormat(getdate())      
    SELECT @MasterServePlace = SUM(tmpTable.SServePlace)      
    FROM (  SELECT DISTINCT ServePlace As SServePlace      
         FROM Split(@DetailsString)      
           ) tmpTable      

----------------------------------------Date From Server-----------------------------------------------------------------      
If @Status = 2 And dbo.Get_DateFromServer() = 1      
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      

------Start New Line For Avoid Repeat in tFacm------
DECLARE @RepeatNo INT

declare @d1 as datetime

set @d1 = CONVERT(datetime ,@NewTime)
set @d1 = DATEADD(MINUTE,-1,@d1)

--select  CONVERT(VARCHAR(5),@d1,108)

SELECT @RepeatNo = COUNT(*) FROM dbo.tFacM WHERE Status = 2 AND [Date] = @Date 
  --AND [Time] <= @NewTime AND [Time] >= CONVERT(VARCHAR(5),@d1,108) 
  AND TableNo = @TableNo AND Recursive = 0 AND Balance = 0

IF @RepeatNo > 0 
    GOTO EventHandler

----End New Line -----------------------------------------------------------------------------------------------      

 Declare @intBranch  int      
 Declare @ShiftNo int      
 DECLARE @TempNo INT 

 select @intBranch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@DetailsString))      
 IF @intBranch = 0 OR @intBranch IS NULL     SET @intBranch = dbo.Get_Current_Branch()

    DECLARE @IdentityNo INT
    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
    FROM    tFacm
    WHERE   Branch = @intBranch 

    IF @IdentityNo < ( @intBranch * 10000000 ) 
        SET @IdentityNo = ( @intBranch * 10000000 )

 SET @NO1 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND AccountYear = @AccountYear)      

 SET @ShiftNo= dbo.Get_Shift(GETDATE())      
 SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      


     INSERT INTO tFacM (   
		intSerialNo ,   
		[No] ,      
		[Date] ,      
		RegDate ,      
		Status ,      
		Customer ,      
		SumPrice ,      
		OrderType ,      
		ServePlace ,      
		StationId ,      
		ServiceTotal ,      
		Recursive ,      
		CarryFeeTotal ,      
		PackingTotal ,      
		DiscountTotal ,      
		[Time] ,      
		[User] ,      
		TableNo ,      
		shiftNo ,      
		incharge,      
		owner ,      
		FacPayment ,       
		Balance ,       
		Branch,      
		AccountYear ,      
		NvcDescription,      
		TempAddress ,
		GuestNo ,
		TempNo    
		
 )      
     Values       

(			    @IdentityNo ,  
                @NO1 ,      
                @Date ,      
                dbo.Shamsi(GETDATE()) ,      
                @Status,      
                @Customer ,      
                @SumPrice ,      
                @OrderType ,      
                @MasterServePlace ,      
                @StationId ,      
                @ServiceTotal ,      
                @Recursive ,      
                @CarryFeeTotal ,      
                @PackingTotal ,      
                @DiscountTotal ,      
                @newtime,      
                @User ,      
                @TableNo,      
                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
                @Incharge ,      
                @owner ,      
                @FacPayment ,      
                @Balance ,      
		@intBranch , --dbo.Get_Current_Branch(),      
		@AccountYear ,      
		@NvcDescription,      
		@TempAddress,
		@GuestNo,
		@TempNo  
 )      
    SET @intserialNo = @IdentityNo
     IF @@ERROR <>0      
        GoTo EventHandler       



declare @destbranch  INT 
SET @destbranch = 0
DECLARE @TempNo2 INT 
DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      
 
If @Status = 6 AND @DestinventoryNo > 0  -- And (@destbranch= @intBranch Or dbo.AutoResid() = 1)    
	Begin      
	select @destbranch=  @intBranch --   branch from tInventory where inventoryNo=(SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

          SET @NO2 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @destbranch AND AccountYear = @AccountYear)      
		  --SET @TempNo2 = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      

     INSERT INTO tFacM ( 
				intSerialNo ,     
                [No] ,      
                [Date] ,      
                RegDate ,      
                Status ,      
                Customer ,      
                SumPrice ,      
                OrderType ,      
                ServePlace ,      
                StationId ,      
                ServiceTotal ,      
                Recursive ,      
                CarryFeeTotal ,      
                PackingTotal ,      
                DiscountTotal ,      
                TaxTotal ,
                DutyTotal ,     
                [Time] ,      
                [User] ,      
                TableNo ,      
                shiftNo ,      
                incharge,      
                owner ,      
                FacPayment ,       
                Balance ,       
                Branch,      
			  AccountYear ,      
			  NvcDescription,      
			  TempAddress,
			  GuestNo ,
			  TempNO     

 )      
     Values      
(				@IdentityNo + 1 ,     
                @NO2 ,      
                @Date ,      
                dbo.Shamsi(GETDATE()) ,      
                7,      
                @Customer ,      
                @SumPrice ,      
                @OrderType ,      
                @MasterServePlace ,      
                @StationId ,      
                @ServiceTotal ,      
                @Recursive ,      
                @CarryFeeTotal ,      
                @PackingTotal ,      
                @DiscountTotal ,      
                0 ,
                0 ,      
                @newtime,      
                @User ,      
                @TableNo,      
                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
                @Incharge ,      
                @owner ,      
                @FacPayment ,      
                @Balance ,      
				@DestBranch ,     
				@AccountYear ,      
				@NvcDescription,      
				@TempAddress,
				@GuestNo ,
				NULL --@TempNo2    
		
 )      
		SET @intserialNo2 = @IdentityNo + 1      
		 IF @@ERROR <>0      
			GoTo EventHandler      

            UPDATE  tfacm
            SET     NvcDescription = @NvcDescription + N' رسيد -   '
                    + CAST(@No2 AS NVARCHAR(8))
            WHERE   intSerialNo = @intserialNo
            UPDATE  tfacm
            SET     RefrenceHavale = @intserialNo2
            WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch

end      


----------------------------------Fill Details Factor  --------------------------------------------------------------      
If @Status = 6 AND @DestinventoryNo > 0 -- AND (@destbranch= @intBranch  Or dbo.AutoResid() = 1)        
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch      
Else       
 exec InsertFactorDetail @DetailsString , @intserialNo , 0, @Customer , @intBranch      

     IF @@ERROR <>0      
        GoTo EventHandler      
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------      

----------------------------------Total SumPrice Calculate  --------------------------------------------------------------      
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * discount/100 ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * (1 - discount/100) ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

     IF @@ERROR <>0      
        GoTo EventHandler      
DECLARE @TaxTotal FLOAT  
SET @TaxTotal = 0
DECLARE @ValueGoodsTax FLOAT
SET @ValueGoodsTax = 0
IF @status = 2 OR @status = 5  OR @status = 10
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxBuy = 1
	            )  
SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
IF @@ERROR <> 0 
GOTO EventHandler

DECLARE @DutyTotal INT 
SET @DutyTotal = 0
DECLARE @ValueGoodsDuty FLOAT
SET @ValueGoodsDuty = 0
IF @status = 2 OR @status = 5  OR @status = 10
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutyBuy = 1
	            )  
SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
IF @@ERROR <> 0 
GOTO EventHandler

Declare @SumPrice2 Bigint      
Set @SumPrice2 = (Select Cast(Sum(Amount * FeeUnit) as Bigint) From tFacd Where intSerialNo = @intserialNo2 And Branch = @DestBranch )        
     IF @@ERROR <>0      
        GoTo EventHandler      
----------------------------------ServiceRate Calculate  --------------------------------------------------------------      
Declare @ReserveServiceRate Int      
Set @ReserveServiceRate = 0      

If  @TableNo >0      
Begin      
	Declare @Reserve Bit      
	Set @Reserve = (Select Reserve From tTable Where tTable.[No] = @TableNo)      
	If @Reserve = 1      
		Set @ReserveServiceRate = (Select t.ReserveServiceRate From(Select ReserveServiceRate From tPartitions  inner join tTable        
		On tPartitions.PartitionId = tTable.PartitionId and  tTable.[No] = @TableNo)t )      

        Update dbo.tTable      
           Set   dbo.tTable.Empty  = 0      
                Where dbo.tTable.[No] = @TableNo AND  @Balance = 0    
	If dbo.Get_TableMonitoring() = 1   ---Table Monitoring      
	 	Begin      
		DECLARE @intTableUsedNo INT      
--		SELECT @intTableUsedNo=intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
--		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@intBranch      
		DECLARE @nvcString NVARCHAR(100)      
		SET @nvcString=','+CAST(@TableNo AS NVARCHAR(5))+'/'      
		--IF @intTableUsedNo is NULL      
		EXEC insert_tblSamar_TableUsage @nvcString,1      
--		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.nvcStartTime=  @newtime      
--		FROM    ( SELECT     dbo.vwSamar_TableUsage_BusyTable.intTableUsedNo, dbo.vwSamar_TableUsage_BusyTable.nvcStartTime,       
--		                 dbo.vwSamar_TableUsage_BusyTable.intBranch, dbo.tTable.[No]      
--				FROM         dbo.tTable LEFT OUTER JOIN      
--		                 dbo.vwSamar_TableUsage_BusyTable ON dbo.vwSamar_TableUsage_BusyTable.intTableNo = dbo.tTable.[No] AND       
--		                 dbo.vwSamar_TableUsage_BusyTable.intBranch = dbo.tTable.Branch)t      
--		WHERE  tblSamar_TableUsage.intTableNo=t.[No] and tblSamar_TableUsage.intBranch=t.intBranch      
--		and tblSamar_TableUsage.intTableNo=@TableNo and tblSamar_TableUsage.intBranch= @intBranch     
		END        
End      


If @ReserveServiceRate > 0       
 Set @ServiceTotal = @ReserveServiceRate      


 If @ServiceTotal <> 0      
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)      
     IF @@ERROR <>0      
        GoTo EventHandler       
----------------------------------Round Sumprice  --------------------------------------------------------------      
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5  OR @status = 10
 BEGIN 
  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
 END 
    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal     

    Declare @Remain INT
    SET @Remain = 0  
    IF @Status = 2 OR @Status = 10
    BEGIN   
    Set @Remain = dbo.RoundSumPrice(@SumPrice )         
    Set @SumPrice = @SumPrice - @Remain      
    Set @DiscountTotal = @DiscountTotal + @Remain    
    END  
---select @Remain as remain      
----------------------------------Calculate Packing---------------------------------------------------------------      
If dbo.Get_AutoPacking() = 1      
Begin      
    Declare @UserPacking INT      
    SET @UserPacking = ISNULL((Select ISNULL(Sum(Amount),0)  From tGood inner join tFacD On tFacD.GoodCode = tGood.Code       
        where intSerialNo = @intSerialNo and MainType =1 and Branch = @intBranch Group By MainType) ,0)      
    Set @UserPacking = @UserPacking * [dbo].[Get_User_Packing]()      
    Set @SumPrice = @SumPrice + @UserPacking      
    Update tFacm Set PackingTotal = @UserPacking  Where intSerialNo = @intserialNo And Branch = @intBranch       
End      
----------------------------------Net Price Update  --------------------------------------------------------------      

Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch      
Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DiscountTotal = @DiscountTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

If @Status = 6 AND @DestinventoryNo > 0-- AND (@destbranch= @intBranch )  -- Or dbo.AutoResid() = 1   
	Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch       
      IF @@ERROR <>0       

        GoTo EventHandler           
-------------------------------------Fill Detail Cash , ..........--------------------------------------------------      
IF (@Status =  1 OR @Status = 2 )      
 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds ,  @intBranch  , @Remain    

     IF @@ERROR <>0      
   GoTo EventHandler      
-------------------------------------Monitoring---------------------------------------------------------------------      
--Declare  @Monitor1 int      
--Declare  @Monitor2 int       

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  @intBranch)      
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  @intBranch)      


--IF @Monitor1 > 0       
--   exec Notify_to_Clients      

--Else If @Monitor2 > 0       
--   exec Notify_to_Clients      

----------------------------History---------------------------      

Exec InsertHistory  @No1, @Status , @User , 1 , @AccountYear , @intBranch      
IF @Status = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 7 , @User , 1 , @AccountYear , @destbranch      


----------------------------Cash ---------------------------      

------------------------Mojodi Control Online--------------------------------------------      

Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @intBranch      
IF @@ERROR <>0      
 GoTo EventHandler      
IF @Status = 6 AND @DestinventoryNo > 0--AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1      
 BEGIN      
 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch      
 IF @@ERROR <>0      
 GoTo EventHandler      
 Exec InsertMojodiCalculate  7,  @intserialNo2 , @AccountYear , @destbranch      
 IF @@ERROR <>0      
 GoTo EventHandler      
 END       

------------------------------------------Update Balance After Recived----------------------------
Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch


COMMIT TRAN

--DECLARE @TemporaryNo BIT 
--SELECT @TemporaryNo = TemporaryNo FROM dbo.tStations WHERE StationID = @StationId AND Branch = @intBranch
--IF @TemporaryNo = 0 set @lastFacMNo = @No1
--ELSE set @lastFacMNo = @TempNo

set @lastFacMNo = @intserialNo


---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @lastFacMNo , 1

--------------------------------------------------------------------------------------------------------------------------------------


Return @lastFacMNo      

EventHandler:      

    ROLLBACK TRAN      
    SET @LastFacMNo = -1      

    RETURN @lastFacMNo
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--تغییر تاریخ فاکتور خرید و عدم تغییر تاریخ فاکتور فروش
ALTER    PROCEDURE [dbo].[EditFactorMasterDetails]  (  


	@No       INT,  
	@Status  INT ,  
	@Owner  INT ,  
	@Customer  INT ,  
	@DiscountTotal Float ,  
	@CarryFeeTotal Float ,  
	@Recursive  INT ,  
	@InCharge  INT ,  
	@FacPayment  BIT ,  
	@OrderType  INT ,  
	@StationId  INT ,  
	@ServiceTotal  Float ,  
	@PackingTotal  Float ,  
	@TableNo  INT ,  
	@User INT ,  
	@Date   Nvarchar(50) =NULL,  
	@DetailsString  nText,  
	@ds nText = '',  
	@Balance Bit,  
	@AccountYear Smallint = Null ,  
	@NvcDescription Nvarchar(150) = Null ,  
	@TempAddress Nvarchar(255) = '', 
	@GuestNo INT,     
	@LastFacMNo  INT OUT  ,
	@Person INT = NULL 
  )  


AS  
DECLARE @SumPrice BIGINT  
DECLARE @SumPrice2 BIGINT  
DECLARE @intSerialNo BIGINT  
DECLARE @intSerialNo2 BIGINT  
--DECLARE @intSerialNo3 BIGINT  
DECLARE @OldRegDate Nvarchar(50)  
DECLARE  @FactorSerial BIGINT  
Declare @Result int  

SET @Sumprice = 0  
SET @Sumprice2= 0  
SET @intSerialNo = 0  
SET @intSerialNo2 = 0  
--SET @intserialNo3 = 0  


 Declare @intBranch  int  
 Declare @ShiftNo int  

 Declare @DestBranch INT  
 SET @DestBranch = 0

IF @AccountYear IS  Null
    Set @AccountYear = dbo.get_AccountYear() 

 select @intBranch = branch from tInventory where inventoryNo=(SELECT TOP 1  IntInventoryNo FROM Split(@DetailsString))  
 SET @ShiftNo= dbo.Get_Shift(GETDATE())  

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  

--Control is difficult
--If No received then Bypass received
--DECLARE @DestinventoryNo INT 
--select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

IF @Status = 6  
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  


if @status=10   
set @OldRegDate = (SELECT tFacM.regdate FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  
else set @OldRegDate=dbo.Shamsi(GETDATE())  
-------------No Change StationId , If this Fich Is For Pocket Pc---------------------------------------  
DECLARE @OldStationId INT  
 SET @OldStationId = (Select StationId From tFacm Where intserialNo = @intSerialNo and Branch =  dbo.Get_Current_Branch())  

DECLARE @StationType INT  
 SET @StationType = (Select StationType From tStations Where StationId = @OldStationId and Branch =  dbo.Get_Current_Branch())  
If  @StationType = 8  
 SET @StationId = @OldStationId  
----------------------------------------------------------------------------------------------------------  
IF  @Owner = 0  
    SET @Owner = NULL  

IF  @TableNo < 1  
    SET @TableNo = NULL  

Declare @OldTableNo   int  

SET  @OldTableNo =  IsNull((SELECT tFacM.TableNo FROM tFacM WHERE intSerialNo = @intSerialNo and Branch = dbo.Get_Current_Branch()) , 0)  

IF  @Incharge < 1  
    SET @Incharge = NULL  

IF  @Customer=0  
    SET @Customer = NULL  
IF @Date IS NULL  
 SET @Date=Rtrim(LTRIM(dbo.Shamsi(GETDATE())))  

BEGIN TRANSACTION  

If IsNull(@TableNo , 0) <> @OldTableNo  
BEGIN  
 IF @OldTableNo > 0   
 Update ttable SET Empty = 1 where No = @OldTableNo  
END  

    DECLARE @MasterServePlace INT  

 SELECT @MasterServePlace = SUM(tmpTable.SServePlace)  
 FROM   
 (  SELECT DISTINCT ServePlace As SServePlace  FROM Split(@DetailsString)) tmpTable  


 if @Status = 2  
 begin  
       INSERT INTO tRepFacEditM (Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance , OrderType,  
                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate, AccountYear , TaxTotal , DutyTotal  )  
          SELECT Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance, OrderType,  
                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate , AccountYear , TaxTotal , DutyTotal    
   FROM tFacM WHERE tFacM.intSerialNo = @intSerialNo and Branch = @intBranch  

      IF @@ERROR <>0  
          GoTo EventHandler  

      INSERT INTO tFacD2(Code , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate], intInventoryNo )   
    SELECT @@identity , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate],intInventoryNo  
                 From tFacD  
                 WHERE intSerialNo = @intSerialNo  And Branch = @intBranch

      IF @@ERROR <>0  
          GoTo EventHandler  

 end  



If @status = 6 AND @intSerialNo2 > 0 --And (@destbranch = dbo.Get_Current_Branch()  )--or dbo.AutoResid() = 1   
 select @destbranch=branch from tInventory where inventoryNo=(SELECT TOP 1 DestInventoryNo FROM Split(@DetailsString))  

---------------------------------------Mojodi Control Online---------------------------------------------------------  
Exec DeleteMojodiCalculate @Status , @intserialNo  ,  1 , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = @intBranch )  --or dbo.AutoResid() = 1 
 Exec DeleteMojodiCalculate 7 , @intserialNo2  , 1 , @AccountYear , @DestBranch  
----------------------------------------Delete Old Details -----------------------------------------------------------  
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = @intBranch or dbo.AutoResid() = 1 )   
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo2 AND Branch =  @DestBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
------------------------------------------------------------    
  Exec DeleteFactorChildren @intSerialNo , @intBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
 If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
  Exec DeleteFactorChildren @intSerialNo2 , @DestBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
----------------------------------------Date From Server-----------------------------------------------------------------  
If @Status = 2 And dbo.Get_DateFromServer() = 1  
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())  
----------------------------------------Update Master-----------------------------------------------------------------  

    Update tFacM  
        SET Owner       = @Owner,  
        Customer        = @Customer,  
        DiscountTotal   = @DiscountTotal,  
        CarryFeeTotal   = @CarryFeeTotal,  
        SumPrice        = @SumPrice,  
        Recursive       = @Recursive,   
        InCharge        = @InCharge,  
        FacPayment      = @FacPayment,  
        Balance         = @Balance,  
        OrderType       = @OrderType,  
        ServePlace      = @MasterServePlace,  
        StationId       = @StationId,  
        ServiceTotal    = @ServiceTotal,  
        PackingTotal    = @PackingTotal,  
        ShiftNo         = @ShiftNo , --dbo.Get_Shift(GETDATE()),  
        TableNo         = @TableNo,  
        [Date]          =  CASE WHEN @Status = 2 THEN  [Date] WHEN @Status = 5 THEN [Date] ELSE @Date END,  
        [Time]          = dbo.SetTimeFormat(GETDATE()),  
        [User]          = @User,  
        RegDate 	= @OldRegDate,---dbo.Shamsi(GETDATE()),  
        NvcDescription  = @NvcDescription ,  
 		TempAddress     = @TempAddress,
		GuestNo		= @GuestNo ,
		TempNo = CASE WHEN @Status = 2 THEN  TempNo WHEN @Status = 5 THEN TempNo ELSE NULL END      
    WHERE tFacM.intSerialNo = @intSerialNo  AND Branch =  @intBranch  

    IF @@ERROR <>0  
        GoTo EventHandler  

DECLARE @No2 INT
SET @No2 = 0

If @Status = 6 AND @intSerialNo2 > 0 --And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
Begin  
    SET @NO2 = (SELECT [NO] FROM tFacM WHERE intserialNo = @intSerialNo2  And Branch =  @DestBranch )      


    Update tFacM  
        SET Owner       = @Owner,  
        Customer        = @Customer,  
        DiscountTotal   = @DiscountTotal,  
        CarryFeeTotal   = @CarryFeeTotal,  
        SumPrice        = @SumPrice,  
        Recursive       = @Recursive,   
        InCharge        = @InCharge,  
        FacPayment      = @FacPayment,  
        Balance         = @Balance,  
        OrderType       = @OrderType,  
        ServePlace      = @MasterServePlace,  
        StationId       = @StationId,  
        ServiceTotal    = @ServiceTotal,  
        PackingTotal    = @PackingTotal,  
        ShiftNo         = @ShiftNo ,--dbo.Get_Shift(GETDATE()),  
        TableNo         = @TableNo,  
        [Date]          = @Date,  
        [Time]          =dbo.SetTimeFormat(GETDATE()),  
        [User]          = @User,  
        RegDate 	= dbo.Shamsi(GETDATE()),  
        NvcDescription  = @NvcDescription,  
 		TempAddress     = @TempAddress ,
		GuestNo		= @GuestNo  
    WHERE tFacM.intSerialNo = @intSerialNo2  AND Branch =  @DestBranch  

END  

----------------------------------Fill Details Factor ----------------------------------------------------------------------  
If @Status = 6  AND @intSerialNo2 > 0--AND (@destbranch= @intBranch Or dbo.AutoResid() = 1 )    
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch  
Else         
 exec InsertFactorDetail @DetailsString , @intserialNo , 0 , @Customer , @intBranch        

     IF @@ERROR <>0  
        GoTo EventHandler  
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------  


----------------------------------Total SumPrice Calculate  --------------------------------------------------------------  
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * discount/100 ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * (1 - discount/100) ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

     IF @@ERROR <>0      
        GoTo EventHandler      
DECLARE @TaxTotal FLOAT  
SET @TaxTotal = 0
DECLARE @ValueGoodsTax FLOAT
SET @ValueGoodsTax = 0
IF @status = 2 OR @status = 5 OR @status = 10
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxBuy = 1
	            )  
SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
IF @@ERROR <> 0 
GOTO EventHandler

DECLARE @DutyTotal INT 
SET @DutyTotal = 0
DECLARE @ValueGoodsDuty FLOAT
SET @ValueGoodsDuty = 0
IF @status = 2 OR @status = 5 OR @status = 10
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutyBuy = 1
	            )  
SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
IF @@ERROR <> 0 
GOTO EventHandler

If @Status = 6 AND @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1 )  
 Set @SumPrice2 = (Select Cast (Sum(Amount * FeeUnit) as Bigint)   From tFacd Where intSerialNo = @intSerialNo2 And Branch = @DestBranch )    
IF @@ERROR <>0  
        GoTo EventHandler  
----------------------------------ServiceRate Calculate  --------------------------------------------------------------  
Declare @ReserveServiceRate Int  
Set @ReserveServiceRate = 0  
If  @TableNo >0  
Begin  
	Declare @Reserve Bit  
	Set @Reserve = (Select Reserve From tTable Where tTable.[No] = @TableNo)  
	If @Reserve = 1  
		Set @ReserveServiceRate = (Select t.ReserveServiceRate From(Select ReserveServiceRate From tPartitions  inner join tTable    
		On tPartitions.PartitionId = tTable.PartitionId and  tTable.[No] = @TableNo)t )  


	If   @Recursive = 0  
	 Update dbo.tTable  
	    Set   dbo.tTable.Empty  = 0  
	        Where dbo.tTable.[No] = @TableNo  AND @Balance = 0
	
	if  @Recursive = 1  
         Update dbo.tTable  
            Set   dbo.tTable.Empty  = 1  
                Where dbo.tTable.[No] = @TableNo  

	If dbo.Get_TableMonitoring() = 1 AND IsNull(@TableNo , 0) <> @OldTableNo   ---Table Monitoring      
	 	Begin      
		DECLARE @intTableUsedNo INT      
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@OldTableNo and vwSamar_TableUsage_BusyTable.intBranch=@intBranch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.intTableNo = @TableNo      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
		END        

End  

If @ReserveServiceRate > 0   
 Set @ServiceTotal = @ReserveServiceRate  


 If @ServiceTotal <> 0  
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)  

     IF @@ERROR <>0  
        GoTo EventHandler   
----------------------------------Round Sumprice  --------------------------------------------------------------  
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5 OR @status = 10
 BEGIN 
  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
 END 
    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal   

    Declare @Remain INT  
    SET @Remain = 0
    IF @Status = 2 OR @status = 10
    BEGIN
    Set @Remain = dbo.RoundSumPrice(@SumPrice )     
    Set @SumPrice = @SumPrice - @Remain  
    Set @DiscountTotal = @DiscountTotal + @Remain  
    END
----------------------------------Calculate Packing---------------------------------------------------------------  
IF dbo.Get_AutoPacking() = 1  
Begin  
    Declare @UserPacking INT  
    SET @UserPacking = ISNULL((Select ISNULL(Sum(Amount),0)  From tGood inner join tFacD On tFacD.GoodCode = tGood.Code   
 where intSerialNo = @intSerialNo and MainType =1 and Branch = @intBranch Group By MainType) ,0)  
    Set @UserPacking = @UserPacking * [dbo].[Get_User_Packing]()  
   Set @SumPrice = @SumPrice + @UserPacking  
   Update tFacm Set PackingTotal = @UserPacking  Where intSerialNo = @intserialNo And Branch = @intBranch   
End  
----------------------------------Net Price Update  --------------------------------------------------------------  

    Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch   
 IF @@ERROR <>0  
         GoTo EventHandler  
If @Status = 6 AND @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1)   

    Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch  
 IF @@ERROR <>0  
         GoTo EventHandler  

Update tFacm Set DiscountTotal = @DiscountTotal Where intSerialNo = @intserialNo  And Branch = @intBranch   
Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch   
Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch   
Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

-----------------------------------------Fill Detail Cash ,....---------------------------------------------------  
If (@Status = 2 OR @Status = 1)  
 
 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds  , @intBranch  , @Remain  
 IF @@ERROR <>0  
        GoTo EventHandler  
-----------------------------------------Monitoring  --------------------------------------------------------------  

--Declare  @Monitor1 int  
--Declare  @Monitor2 int  

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  


--If @Monitor1 > 0   
--  exec Notify_to_Clients  
--Else If @Monitor2 > 0   
--  exec Notify_to_Clients  

-- IF @@ERROR <>0  
--        GoTo EventHandler  

-----------------------------------------History  --------------------------------------------------------------  

Exec InsertHistory  @No, @Status , @User , 2 ,@AccountYear  , @intBranch
 IF @@ERROR <>0  
        GoTo EventHandler  

-----------------------------------------Cash  --------------------------------------------------------------  

------------------------------------------Mojodi Control Online-----------------------------------------------------  

Exec InsertMojodiCalculate @Status , @intserialNo , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
IF @STATUs = 6 AND @intSerialNo2 > 0--AND (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
 BEGIN  
 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch  
 IF @@ERROR <>0  
 GoTo EventHandler  
 Exec InsertMojodiCalculate  7,  @intserialNo2 , @AccountYear , @destbranch  
 IF @@ERROR <>0  
 GoTo EventHandler  
 END   
------------------------------------------Update Balance After Recived----------------------------
Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch


COMMIT TRANSACTION  

---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @intSerialNo , 2

--------------------------------------------------------------------------------------------------------------------------------------
Set @LastFacMNo = @No  
Return @LastFacMNo  


EventHandler:  
    ROLLBACK TRAN  
    SET @LastFacMNo = -1   

    RETURN @LastFacMNo
GO




ALTER    PROCEDURE Update_tFacM_Recursive
(
@No  Bigint,
@Status int,
@Recursive int,
@Uid int,
@Balance Bit,
@FacPayment Bit ,
@AccountYear Smallint = NULL ,
@Branch INT 
)

AS
Declare @TableNo int
DECLARE @intTableUsedNo INT      
IF @AccountYear Is Null 
	SET @AccountYear = dbo.Get_AccounYear()

DECLARE @intSerialNo BIGINT

--DECLARE @Branch INT
--	SET @Branch = dbo.Get_Current_Branch()

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch and AccountYear = @AccountYear)

UPDATE tFacM
     SET Recursive= @Recursive
         WHERE tFacM.intSerialNo = @intserialNo And  Branch = @Branch 

If @Status = 6
BEGIN 
	DECLARE @intserialNo2 BIGINT
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear)  
	IF @intSerialNo2 > 0
		UPDATE tFacM
			 SET Recursive= @Recursive
				 WHERE tFacM.intSerialNo = @intserialNo2 And  Branch = @Branch 
END 

If @Recursive = 1 
Begin

UPDATE tFacM
     SET FacPayment = 0 , Balance = 0
         WHERE tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear

SET @TableNo = ISNULL((Select TableNo From tfacm   Where  tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear) , 0)
  UPDATE tTable
       SET Empty = 1 
           WHERE dbo.tTable.[No] = @TableNo
	If dbo.Get_TableMonitoring() = 1 AND @TableNo > 0		---Table Monitoring
	Begin
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.bitIsValid = 0      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	


Exec DeleteFactorChildren @intSerialNo , @Branch

UPDATE dbo.tblAcc_Recieved SET Bestankar = 0 WHERE intSerialNo = @intSerialNo And  Branch = @Branch  

End

If @Recursive = 0

Begin
   Update tFacm 
       SET FacPayment = @FacPayment , Balance = @Balance
         WHERE tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear

	SET @TableNo = ISNULL((Select TableNo From tfacm   Where  tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear) , 0)
    UPDATE tTable
       SET Empty = 0 
           WHERE dbo.tTable.[No] = @TableNo
	If dbo.Get_TableMonitoring() = 1 AND @TableNo > 0		---Table Monitoring
	Begin
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.bitIsValid = 1      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	

	IF @Balance = 1
	BEGIN 
	DELETE FROM tFacCash WHERE intSerialNo = @intSerialNo AND [Branch] = @Branch
	INSERT INTO tFacCash (intSerialNo, intAmount ,branch)
		SELECT @intSerialNo AS
	 intSerialNo, Sumprice,@Branch From tFacM  WHERE tFacM.[No]=@No   AND Status = 2 And  Branch = @Branch and AccountYear = @AccountYear

	END 
End

--Declare @Monitor1 Bit
--Declare @Monitor2 Bit

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())


--If @Monitor1 > 0 
--  exec Notify_to_Clients
--Else If @Monitor2 > 0 
--  exec Notify_to_Clients

If @Recursive = 0
   Exec InsertHistory  @No, @Status , @Uid , 8 ,@AccountYear , @Branch
Else if @Recursive = 1
   Exec InsertHistory  @No, @Status , @Uid , 3 ,@AccountYear , @Branch 

---------------------------------------Mojodi Control Online---------------------------------------------------------

Exec DeleteMojodiCalculate @Status , @intserialNo , @Recursive ,@AccountYear , @Branch
If @Status = 6 AND @intserialNo2 > 0
	EXEC DeleteMojodiCalculate 7, @intSerialNo2 , @Recursive, @AccountYear , @Branch

--------------------------------------------------------------------------------------------------------------------------------------

---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @intSerialNo , 3

--------------------------------------------------------------------------------------------------------------------------------------
GO



--Kitchen Monitoring
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fn_KM_GetParentStations]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[fn_KM_GetParentStations]
GO

CREATE  function [fn_KM_GetParentStations](@StationID INT)
Returns @Results Table (Items int)
As
Begin


Declare @Index int
Declare @Slice nvarchar(4000)
Declare @Delimiter char(1)
Declare @PStations nvarchar(4000)

SELECT @PStations = ParentStations FROM dbo.tStations WHERE StationID = @StationID

SET @Delimiter = ','
Select @Index = 1
If @PStations Is NULL Return

While @Index != 0
Begin
Select @Index = CharIndex(@Delimiter, @PStations)
If @Index <> 0

Select @Slice = left(@PStations, @Index - 1)

else

Select @Slice = @PStations
IF ISNUMERIC(@Slice) = 1
	 Insert into @Results(Items) Values (CAST(@Slice AS INT))

SELECT @PStations = right(@PStations, Len(@PStations) - @Index)

If Len(@PStations) = 0 break

End
Return
END



GO
