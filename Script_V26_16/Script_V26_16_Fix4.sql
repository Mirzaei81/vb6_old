
--Script_V26_16_Fix4
--93/03/13

--امکان تغییر کد کالا  و زیر گروهها و گروههای اصلی

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
          4
        )
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tInventory_Level1_tGoodLevel1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].tInventory_Level1 DROP CONSTRAINT FK_tInventory_Level1_tGoodLevel1
GO
 
ALTER TABLE dbo.tInventory_Level1 WITH NOCHECK ADD CONSTRAINT
	FK_tInventory_Level1_tGoodLevel1 FOREIGN KEY
	(
	Level1
	) REFERENCES dbo.tGoodLevel1
	(
	Code
	) ON UPDATE CASCADE
	
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTotal_ChargeGood_tGood]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTotal_ChargeGood] DROP CONSTRAINT FK_tblTotal_ChargeGood_tGood
GO
 
ALTER TABLE dbo.tblTotal_ChargeGood WITH NOCHECK ADD CONSTRAINT
	FK_tblTotal_ChargeGood_tGood FOREIGN KEY
	(
	GoodCode
	) REFERENCES dbo.tGood
	(
	Code
	) ON UPDATE CASCADE
	
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTotal_Inventory_Good_CycleStock_tGood]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].tblTotal_Inventory_Good_CycleStock DROP CONSTRAINT FK_tblTotal_Inventory_Good_CycleStock_tGood
GO

ALTER TABLE dbo.tblTotal_Inventory_Good_CycleStock WITH NOCHECK ADD CONSTRAINT
	FK_tblTotal_Inventory_Good_CycleStock_tGood FOREIGN KEY
	(
	GoodCode
	) REFERENCES dbo.tGood
	(
	Code
	) ON UPDATE CASCADE
	
GO


--if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tUsePercent_tGood]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
--ALTER TABLE [dbo].[tUsePercent] DROP CONSTRAINT [FK_tUsePercent_tGood]
--GO


--ALTER TABLE [dbo].[tUsePercent]  WITH NOCHECK ADD  CONSTRAINT [FK_tUsePercent_tGood] FOREIGN KEY([GoodCode])
--REFERENCES [dbo].[tGood] ([Code])
--ON UPDATE CASCADE
--GO

--ALTER TABLE [dbo].[tUsePercent] CHECK CONSTRAINT [FK_tUsePercent_tGood]
--GO


ALTER  PROCEDURE dbo.UpdatetGood
(
	@intLanguage	INT,
	@Goodname	NVARCHAR(50),
	@GoodNamePrn	NVARCHAR(50),
	@SellPrice	FLOAT,
	@BuyPrice	FLOAT,
	@Unit		INT,
	@GoodType	INT,
	@Barcode	NVARCHAR(50),
	@Code		INT,
	@Weight	Float,
	@NumberOfUnit 	INT,
	@SellPrice2 Float,
	@SellPrice3 Float ,
	@MainType Bit ,
	@Supplier Int ,
	@Level1 Int ,
	@Level2 Int ,
	@SellPrice4 Float,
	@SellPrice5 Float,
	@SellPrice6 Float,
	@CategoryShow INT ,
	@PicturePath NVARCHAR(100) ,
	@nvcDescription NVARCHAR(100) ,
	@Picture IMAGE ,
	@GoodNamePrn2	NVARCHAR(100),
	@GoodNamePrn3	NVARCHAR(100),
	@RealNewCode INT ,
	@Result Int Out
)

AS

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tUsePercent_tGood1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
		ALTER TABLE [dbo].[tUsePercent] DROP CONSTRAINT [FK_tUsePercent_tGood1]


Declare  @NewCode INT
SET @NewCode = @RealNewCode
IF @RealNewCode = 0 
BEGIN 
	Set @NewCode = @Code
	Declare @Level2Code	INT
	Set @Level2Code = (Select Level2 From tGood Where Code = @Code)

	Begin Tran

	If @Level2 <>  @Level2Code
	Begin
	--	Set @NewCode =  (SELECT  ISNULL(MAX(RIGHT(RTRIM(LTRIM(STR(code))),LEN(RTRIM(LTRIM(STR(Code))))-4)),0) +1   
		Set @NewCode =  (SELECT  ISNULL(MAX(code),0) + 1   
		FROM dbo.tgood 
		WHERE Level2 = @Level2 )
		If Len(@NewCode) = 1 

		Set @NewCode = (@Level2 * 10000) + @NewCode 

	End

END 

IF @intLanguage = 0 
Begin		
		UPDATE dbo.tGood

		SET [Name]    = dbo.Get_ArabicToFarsiString(@GoodName) ,
		    NamePrn   = dbo.Get_ArabicToFarsiString(@GoodNamePrn) ,
		    SellPrice = @SellPrice ,
		    BuyPrice  = @BuyPrice ,
		    Unit      = @Unit ,
		    GoodType  = @GoodType ,
		    Barcode = @Barcode,
	                 Weight = @Weight,
		    NumberOfUnit=@NumberOfUnit,
		    SellPrice2 = @SellPrice2,
		    SellPrice3 = @SellPrice3 ,	    	
		    SellPrice4 = @SellPrice4 ,	    	
		    SellPrice5 = @SellPrice5 ,	    	
		    SellPrice6 = @SellPrice6 ,	    	
		    MainType = @MainType  ,
		    ProductCompany = @Supplier ,
		   Level1 = @Level1 ,
		   Level2 = @Level2 ,
		 Code = @NewCode ,
		 CategoryShow = @CategoryShow ,
		 PicturePath = @PicturePath ,
		 nvcDescription = @nvcDescription ,
	    GoodNamePrn2   = dbo.Get_ArabicToFarsiString(@GoodNamePrn2) ,
	    GoodNamePrn3   = dbo.Get_ArabicToFarsiString(@GoodNamePrn3) 
	WHERE Code = @Code		
        IF @@ERROR <>0
	        GoTo EventHandler

End
ELSE IF @intLanguage = 1 
Begin
		UPDATE dbo.tGood

		SET LatinName     = @GoodName ,
		    LatinNamePrn  = @GoodNamePrn ,
		    SellPrice     = @SellPrice ,
		    BuyPrice      = @BuyPrice ,
		    Unit          = @Unit ,
		    GoodType      = @GoodType,
		    Barcode = @Barcode,
		    Weight = @Weight,
		    NumberOfUnit=@NumberOfUnit,
		    SellPrice2 = @SellPrice2,
		    SellPrice3 = @SellPrice3 ,
		    SellPrice4 = @SellPrice4 ,	    	
		    SellPrice5 = @SellPrice5 ,	    	
		    SellPrice6 = @SellPrice6 ,	    	
		    MainType = @MainType ,
	 	    ProductCompany = @Supplier ,
		   Level1 = @Level1 ,
		   Level2 = @Level2 ,
		 Code = @NewCode ,
		 CategoryShow = @CategoryShow ,
		 PicturePath = @PicturePath ,
		 nvcDescription = @nvcDescription ,
	    GoodNamePrn2   = @GoodNamePrn2 ,
	    GoodNamePrn3   = @GoodNamePrn3
		WHERE Code = @Code

        IF @@ERROR <>0
	        GoTo EventHandler

End
Set @Result = 1
	update  [dbo].[tGood]  set [name] = latinname where ([Name] is null or [Name] = ''  ) And Code = @NewCode

	update  [dbo].[tGood] set latinname = [name] where ([latinName] is null or latinname = '') And Code = @NewCode

	update  [dbo].[tGood] set [nameprn]=[latinnameprn] where ([Nameprn] is null or [Nameprn] = '') And Code = @NewCode 

	update  [dbo].[tGood] set [GoodNamePrn2]=[nameprn] where ([GoodNamePrn2] is null or [GoodNamePrn2] = '' ) And Code = @NewCode

	update  [dbo].[tGood] set [GoodNamePrn3]=[nameprn] where ([GoodNamePrn3] is null or [GoodNamePrn3] = '' ) And Code = @NewCode

	update  [dbo].[tGood] set [latinnameprn] = [nameprn] where ([latinNameprn] is null or latinnameprn = '') And Code = @NewCode

	UPDATE dbo.tGood SET Picture = @Picture WHERE Code = @Code

	ALTER TABLE [dbo].[tUsePercent]  WITH CHECK ADD  CONSTRAINT [FK_tUsePercent_tGood1] FOREIGN KEY([GoodFirstCode])
		REFERENCES [dbo].[tGood] ([Code])

	ALTER TABLE [dbo].[tUsePercent] CHECK CONSTRAINT [FK_tUsePercent_tGood1]

COMMIT TRANSACTION


Return @Result


EventHandler:
    ROLLBACK TRAN
    Set @Result = 0
    RETURN @Result



GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ChangeLevel1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].ChangeLevel1
GO


CREATE PROCEDURE ChangeLevel1
@OldLevel1 INT ,
@NewLevel1 INT ,
@Replace BIT ,
@Update INT OUT

AS
BEGIN TRAN

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tGoodLevel2_tGoodLevel1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
	ALTER TABLE [dbo].tGoodLevel2 DROP CONSTRAINT FK_tGoodLevel2_tGoodLevel1


SET @Update = 0
IF @Replace = 0 
	BEGIN
	UPDATE dbo.tGoodLevel1 
		SET Code = @NewLevel1 WHERE Code = @OldLevel1
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGood 
		SET Level1 = @NewLevel1 WHERE Level1 = @OldLevel1
	UPDATE dbo.tGoodLevel2 
		SET Level1Code = @NewLevel1 WHERE Level1Code = @OldLevel1
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	IF @@ERROR > 0 
		GOTO ErrorHandler
	
	END
ELSE
	BEGIN

	UPDATE dbo.tGoodLevel1 
		SET Code = Code + 1 WHERE Code >= @NewLevel1
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGood 
		SET Level1 = Level1 + 1 WHERE Level1 >= @NewLevel1
	IF @@ERROR > 0 
		GOTO ErrorHandler
	UPDATE dbo.tGoodLevel1 
		SET Code = @NewLevel1 WHERE Code = @OldLevel1
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGood 
		SET Level1 = @NewLevel1 WHERE Level1 = @OldLevel1
	IF @@ERROR > 0 
		GOTO ErrorHandler


	END


	ALTER TABLE dbo.tGoodLevel2 ADD CONSTRAINT
		FK_tGoodLevel2_tGoodLevel1 FOREIGN KEY
		(
		Level1Code
		) REFERENCES dbo.tGoodLevel1
		(
		Code
		)

COMMIT TRAN
SET @Update = 1
	
RETURN @Update

ErrorHandler:
ROLLBACK TRAN
RETURN 0

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ChangeLevel2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].ChangeLevel2
GO

CREATE  PROCEDURE ChangeLevel2
@OldLevel2 INT ,
@NewLevel2 INT ,
@Level1 INT ,
@Replace BIT ,
@Update INT OUT

AS
BEGIN TRAN
SET @Update = 0

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tGoodLevel2_tGoodLevel1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
	ALTER TABLE [dbo].tGoodLevel2 DROP CONSTRAINT FK_tGoodLevel2_tGoodLevel1

IF @Replace = 0
	BEGIN
	UPDATE dbo.tGoodLevel2 
		SET Level1Code = @Level1 WHERE Code = @OldLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGoodLevel2 
		SET Code = @NewLevel2 WHERE Code = @OldLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGood 
		SET Level1 = @Level1 WHERE Level2 = @NewLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	END	
ELSE
	BEGIN
	UPDATE dbo.tGoodLevel2 
		SET Code = Code + 1 WHERE Code >= @NewLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGoodLevel2 
		SET Level1Code = @Level1 WHERE Code = @OldLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGoodLevel2 
		SET Code = @NewLevel2 WHERE Code = @OldLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGood 
		SET Level1 = @Level1 WHERE Level2 = @NewLevel2
	IF @@ERROR <> 0 
		GOTO ErrorHandler	END	

	ALTER TABLE dbo.tGoodLevel2 ADD CONSTRAINT
		FK_tGoodLevel2_tGoodLevel1 FOREIGN KEY
		(
		Level1Code
		) REFERENCES dbo.tGoodLevel1
		(
		Code
		)
COMMIT TRAN
SET @Update = 1
RETURN @Update

ErrorHandler:
ROLLBACK TRAN
RETURN @Update


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
	@LastFacMNo  INT OUT  
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
IF @Status = 6
	SET @intSerialNo2 = @intSerialNo + 1


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



If @status = 6 --And (@destbranch = dbo.Get_Current_Branch()  )--or dbo.AutoResid() = 1   
 select @destbranch=branch from tInventory where inventoryNo=(SELECT TOP 1 DestInventoryNo FROM Split(@DetailsString))  

---------------------------------------Mojodi Control Online---------------------------------------------------------  
Exec DeleteMojodiCalculate @Status , @intserialNo  ,  1 , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
If @status = 6 --And (@destbranch = @intBranch )  --or dbo.AutoResid() = 1 
 Exec DeleteMojodiCalculate 7 , @intserialNo2  , 1 , @AccountYear , @DestBranch  
----------------------------------------Delete Old Details -----------------------------------------------------------  
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
If @status = 6 --And (@destbranch = @intBranch or dbo.AutoResid() = 1 )   
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo2 AND Branch =  @DestBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
------------------------------------------------------------    
  Exec DeleteFactorChildren @intSerialNo , @intBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
 If @status = 6 --And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
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

If @Status = 6 --And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
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
If @Status = 6  --AND (@destbranch= @intBranch Or dbo.AutoResid() = 1 )    
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

If @Status = 6 --And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1 )  
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
If @Status = 6 --And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1)   

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

Declare  @Monitor1 int  
Declare  @Monitor2 int  

Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  
Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  


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
IF @STATUs = 6 --AND (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
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

Set @LastFacMNo = @No  
Return @LastFacMNo  


EventHandler:  
    ROLLBACK TRAN  
    SET @LastFacMNo = -1   

    RETURN @LastFacMNo
GO

