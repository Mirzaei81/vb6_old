

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


-------------------------------------*******************************************
--اصلاح نحوه جستجو در فرم جستجو کالا فاکتور خرید , فروش , کاردکس کالا و غیره********


ALTER   PROC [dbo].[Get_Good_Name]
    (
      @Name1 NVARCHAR(20),
      @NotSupportedGoodType INT
    )
AS 

Set @Name1 = Replace(  @Name1  , N'ك', N'ک'  ) 
Set @Name1 = Replace(  @Name1  , N'ي' , N'ی' )

    SELECT  *
    FROM    [dbo].[vw_Good]
    WHERE   CHARINDEX(@Name1, [Name]) > 0
            AND [dbo].[vw_Good].[GoodType] NOT IN ( @NotSupportedGoodType, 4 )
    ORDER BY [Name]

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   proc Update_Good_btnAscDefault 
AS

Update tGood SET btnAscDefault = ASCII(left(ltrim([Name] COLLATE Arabic_CI_AI),1))

Update tGood
Set [Name] = Replace(  [Name]   , N'ك', N'ک' ) 

Update tGood
Set [NamePrn] = Replace(  [NamePrn]  , N'ك', N'ک'  ) 

Update tPocketPC_Good
Set [NameDisplay] = Replace(  [NameDisplay]  , N'ك', N'ک'  ) 

Update tGood
Set [Name] = Replace(  [Name], N'ي' , N'ی'  ) 

Update tGood
Set [NamePrn] = Replace(  [NamePrn] , N'ي', N'ی'   ) 

Update tPocketPC_Good
Set [NameDisplay] = Replace(  [NameDisplay]  , N'ي', N'ی'  ) 

Update [tCust]
Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

	Update tper 
	Set nvcFirstName = Replace(  nvcFirstName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcFirstName = Replace(  nvcFirstName , N'ي' , N'ی' ) 

	Update tper 
	Set nvcSurName = Replace(  nvcSurName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcSurName = Replace(  nvcSurName , N'ي' , N'ی' ) 



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  PROCEDURE dbo.InserttGood
(
	@intLanguage	INT,
	@Code		INT,
	@GoodName	NVARCHAR(50),
	@GoodNamePrn	NVARCHAR(50),
	@SellPrice	FLOAT,
	@BuyPrice	FLOAT,
	@Barcode	NVARCHAR(50),
	@Level1		INT,
	@Level2		INT,
	@Model		INT,
	@Supplier	INT,
	@Unit		INT,
	@GoodType	INT,
	@Weight	Float,
	@NumberOfUnit 	INT,
	@SellPrice2 Float,
	@SellPrice3 Float ,
	@MainType Bit ,
	@SellPrice4 Float,
	@SellPrice5 Float,
	@SellPrice6 FLOAT ,
	@CategoryShow INT ,
	@PicturePath NVARCHAR(100) ,
	@nvcDescription NVARCHAR(100) ,
	@Picture IMAGE ,
	@GoodNamePrn2	NVARCHAR(100),
	@GoodNamePrn3	NVARCHAR(100),
	@Result INT OUT 
	


)

AS
Begin tran

	IF @intLanguage = 0 

		INSERT INTO dbo.tGood (Code,Name,NamePrn,SellPrice,BuyPrice,Barcode,Level1,Level2,Model,ProductCompany,Unit,GoodType,Weight,NumberOfUnit,SellPrice2 ,SellPrice3 , MainType , SellPrice4 ,SellPrice5 ,SellPrice6 , CategoryShow , PicturePath , nvcDescription , GoodNamePrn2 , GoodNamePrn3 )
		VALUES (@Code,dbo.Get_ArabicToFarsiString(@GoodName),dbo.Get_ArabicToFarsiString(@GoodNamePrn),@SellPrice,@BuyPrice,@Barcode,@Level1,@Level2,@Model,@Supplier,@Unit,@GoodType,@Weight,@NumberOfUnit, @SellPrice2 , @SellPrice3 , @MainType , @SellPrice4, @SellPrice5, @SellPrice6 , @CategoryShow , @PicturePath , @nvcDescription , dbo.Get_ArabicToFarsiString(@GoodNamePrn2) , dbo.Get_ArabicToFarsiString(@GoodNamePrn3)  )

	ELSE IF @intLanguage = 1 

		INSERT INTO dbo.tGood (Code,LatinName,LatinNamePrn,SellPrice,BuyPrice,Barcode,Level1,Level2,Model,ProductCompany,Unit,GoodType,Weight,NumberOfUnit, SellPrice2, SellPrice3 , MainType , SellPrice4 ,SellPrice5 ,SellPrice6 , CategoryShow ,PicturePath ,nvcDescription , GoodNamePrn2 , GoodNamePrn3)
		VALUES (@Code,@GoodName,@GoodNamePrn,@SellPrice,@BuyPrice,@Barcode,@Level1,@Level2,@Model,@Supplier,@Unit,@GoodType,@Weight,@NumberOfUnit , @SellPrice2 , @SellPrice3 , @MainType , @SellPrice4, @SellPrice5, @SellPrice6 ,@CategoryShow ,@PicturePath ,@nvcDescription , dbo.Get_ArabicToFarsiString(@GoodNamePrn2) , dbo.Get_ArabicToFarsiString(@GoodNamePrn3)  )

     IF @@ERROR <>0
        GoTo EventHandler

 --         if  @GoodType = 3 
   --          Begin
     --               INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
	--	VALUES (@Code,@Code,1,1)
               --     INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
		--VALUES (@Code,@Code,2,1)
                  --  INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
--	--	VALUES (@Code,@Code,4,1)
              --      INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
	--	VALUES (@Code,@Code,8,1)
             --      INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
	--	VALUES (@Code,@Code,16,1)
         --   End

	update  [dbo].[tGood]  set [name] = latinname where ([Name] is null or [Name] = '' ) And Code = @Code

	update  [dbo].[tGood] set latinname = [name] where ([latinName] is null or latinname = '') And Code = @Code

	update  [dbo].[tGood] set [nameprn]=[latinnameprn] where ([Nameprn] is null or [Nameprn] = '' ) And Code = @Code

	update  [dbo].[tGood] set [GoodNamePrn2]=[nameprn] where ([GoodNamePrn2] is null or [GoodNamePrn2] = '' ) And Code = @Code

	update  [dbo].[tGood] set [GoodNamePrn3]=[nameprn] where ([GoodNamePrn3] is null or [GoodNamePrn3] = '' ) And Code = @Code

	update  [dbo].[tGood] set [latinnameprn] = [nameprn] where ([latinNameprn] is null or latinnameprn = '') And Code = @Code

	UPDATE dbo.tGood SET Picture = @Picture WHERE Code = @Code
	
insert into .dbo.[tInventory_Good](GoodCode,InventoryNo,Branch , AccountYear )
  select @Code,tInventory.InventoryNo,tInventory.Branch , dbo.Get_AccountYear()  from dbo.tInventory 
		inner join tInventory_Level1 On tInventory_Level1.Branch = tInventory.Branch  and tInventory_Level1.InventoryNo  = tInventory.InventoryNo 
	        Where tInventory_Level1.Level1 = @Level1 

     IF @@ERROR <>0
        GoTo EventHandler

Insert into dbo.tStation_Inventory_Good ( branch ,InventoryNo, StationID,  AccountYear , GoodCode , Active)
select tInventory.Branch , tInventory.InventoryNo,  t.stationid , dbo.Get_AccountYear() , @Code, 1 as active  from dbo.tInventory -- t.stationid
		inner join tInventory_Level1 On tInventory_Level1.Branch = tInventory.Branch  and tInventory_Level1.InventoryNo  = tInventory.InventoryNo 
         	Inner join (select  Distinct(StationId), InventoryNo , Branch , AccountYear from tStation_Inventory_Good )t On  t.InventoryNo = tInventory_Level1.InventoryNo AND t.Branch =  tInventory_Level1.Branch and t.AccountYear = dbo.Get_AccountYear()

	        Where tInventory_Level1.Level1 = @Level1 

     IF @@ERROR <>0
        GoTo EventHandler
	UPDATE dbo.tGood SET nvcDate = dbo.shamsi(GETDATE()) WHERE Code = @Code
	Update tGood SET btnAscDefault = ASCII(left(ltrim([Name] COLLATE Arabic_CI_AI),1)) where code =  @Code


	Update tGood
	Set [Name] = Replace(  [Name]  , N'ك', N'ک'  ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn]  , N'ك', N'ک'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'ك', N'ک'  ) 

	Update tGood
	Set [Name] = Replace(  [Name] , N'ي' , N'ی' ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn] , N'ي' , N'ی'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'ي', N'ی'  ) 


Commit Tran

SET @Result = 1 
RETURN @Result

EventHandler:
	RollBack Tran
SET @Result = -1 
RETURN @Result



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER     PROCEDURE dbo.UpdatetGood
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

	UPDATE dbo.tGood SET Picture = @Picture WHERE Code = @NewCode
	
	UPDATE [tUsePercent] SET GoodFirstCode = @NewCode WHERE GoodFirstCode = @Code
	
	ALTER TABLE [dbo].[tUsePercent]  WITH CHECK ADD  CONSTRAINT [FK_tUsePercent_tGood1] FOREIGN KEY([GoodFirstCode])
		REFERENCES [dbo].[tGood] ([Code])

	ALTER TABLE [dbo].[tUsePercent] CHECK CONSTRAINT [FK_tUsePercent_tGood1]

	update  [dbo].[tGood] SET nvcDate = dbo.shamsi(GETDATE()) WHERE Code = @Code

	Update tGood
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک'  ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn]  , N'ك', N'ک'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'ك', N'ک'  ) 

	Update tGood
	Set [Name] = Replace(  [Name] , N'ي' , N'ی' ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn] , N'ي' , N'ی'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay] , N'ي', N'ی'   ) 

COMMIT TRANSACTION


Return @Result


EventHandler:
    ROLLBACK TRAN
    Set @Result = 0
    RETURN @Result



GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  PROCEDURE [dbo].[InsertPersonel]( 
	@PersonnelNumber nvarchar(50),
	@nvcFirstName nvarchar(50),
	@nvcSurName nvarchar(50),
	@Gender bit,
	@IdNumber nvarchar(50),
	@Job int,
	@InsuranceNo nvarchar(50) ,
	@Address nvarchar(300),
	@Tel nvarchar(30),
	@User int , 
	@UserName nvarchar(50) ,
	@Password nvarchar(50) ,
	@intAccessLevel int ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno int out

	)
 AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

set @Time = dbo.SetTimeFormat(getdate())

select @pPno = isnull(max(Ppno),0) + 1 from tper Where Branch = @Branch 
If @pPno < (@Branch * 1000 ) Set @pPno = (@Branch * 1000 )


begin Tran
insert into dbo.tper (
	pPno ,
	PersonnelNumber,
	nvcFirstName,
	nvcSurName,
	Gender ,
	IdNumber,
	Job ,
	InsuranceNo  ,
	Address ,
	Tel ,
	[Date] ,
	[Time] ,			
	[User] ,
	Branch,
	MaxCredit,
	ActDeAct

)
values(
	@pPno ,
	@PersonnelNumber,
	@nvcFirstName,
	@nvcSurName ,
	@Gender ,
	@IdNumber ,
	@Job ,
	@InsuranceNo ,
	@Address ,
	@Tel ,
	@Date,
	@Time ,
	@User ,
	@Branch,
	@MaxCredit,
	@ActDeAct
)
if @@Error <> 0 
		GOTO EventHandler	


--set @pPno=@@identity
DECLARE @UID INT
if @intAccessLevel<>0 and @UserName <> '' and @Password<>''

BEGIN

	select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch 
	If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )

	insert into dbo.tUser 
	(
		[Uid] ,
	 UserName ,
	 [Password] ,
	 intAccessLevel ,
	 pPno ,
	 addUser , 
	 Branch, 
	 CountRePrint, 
	 CountInvoicePrint,
	 CountInvoiceEditable,
	 CountInvoiceRefferable
	)
 values (
	@UID ,					
	@UserName  ,
	@Password  ,
	@intAccessLevel ,
	@pPno , 
	@User ,
	@Branch,
	@CountRePrint,
	@CountInvoicePrint,
	@CountInvoiceEditable,
	@CountInvoiceRefferable
	)

if @@Error <> 0 
		GOTO EventHandler	
--SET @UID = @@IDENTITY		
END	


	Update tper 
	Set nvcFirstName = Replace(  nvcFirstName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcFirstName = Replace(  nvcFirstName , N'ي' , N'ی' ) 

	Update tper 
	Set nvcSurName = Replace(  nvcSurName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcSurName = Replace(  nvcSurName , N'ي' , N'ی' ) 


commit Tran



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1




GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER     PROCEDURE [dbo].[UpdatePersonel]( 
	@CurrentPPNO 		INT,
	@PersonnelNumber 	NVARCHAR(50),
	@nvcFirstName 		NVARCHAR(50),
	@nvcSurName	 	NVARCHAR(50),
	@Gender 		BIT,
	@IdNumber 		NVARCHAR(50),
	@Job 			INT,
	@InsuranceNo 		NVARCHAR(50) ,
	@Address 		NVARCHAR(300),
	@Tel 			NVARCHAR(30),
	@User 			INT , 
	@UID 			INT ,
	@UserName 		NVARCHAR(50) ,
	@Password 		NVARCHAR(50) ,
	@intAccessLevel 	INT ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno 			INT OUT
	       )
AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

SET @Time= dbo.SetTimeFormat(getdate())

BEGIN TRANSACTION

	UPDATE tPer
		SET PersonnelNumber 	= @PersonnelNumber,
		    nvcFirstName    	= @nvcFirstName,
		    nvcSurName	    	= @nvcSurName,
		    Gender	    	= @Gender,
		    IdNumber       	= @IdNumber,
		    Job		    	= @Job,
		    InsuranceNo     	= @InsuranceNo,
		    Address	    	= @Address,
		    Tel   	    	= @Tel,
		    [Date]	    	= @Date,
		    [Time]	    	= @Time,
		    [User]	    	= @User,
		    MaxCredit		=@MaxCredit,
		    ActDeAct 		=@ActDeAct ,
		    Branch			= @Branch
	WHERE       pPNO = @CurrentPPNO  


	IF @@ERROR <> 0 
		GOTO EventHandler	

	set @pPno = @CurrentPPNO

	IF @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID<>0
		UPDATE tUser
			SET 		UserName       	= @UserName,
	        	   		 	[Password]     	= @Password,
			    		intAccessLevel 	= @intAccessLevel,
			    		pPno           	= @pPno,
			    		addUser        	= @User,
					 CountRePrint		=@CountRePrint,
		  			 CountInvoicePrint	=@CountInvoicePrint,
					 CountInvoiceEditable		=@CountInvoiceEditable,
		  			 CountInvoiceRefferable	=@CountInvoiceRefferable ,
		  			 Branch					= @Branch
			WHERE   UID = @UID    
	else 
		if @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID=0
		BEGIN 
			select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch   
			If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )  
			insert into dbo.tUser (
						UID ,
						UserName , 
						[Password] , 
						intAccessLevel , 
						pPno , 
						addUser , 
						Branch,
						CountRePrint,
						CountInvoicePrint,
						CountInvoiceEditable,
		  			 	CountInvoiceRefferable
			) values (	
						@UID ,				
						@UserName  ,
						@Password  ,
						@intAccessLevel ,
						@pPno , 
						@User , 
						@Branch,
						@CountRePrint,
						@CountInvoicePrint,
						@CountInvoiceEditable,
		  			 	@CountInvoiceRefferable
			)			
			END 
	IF @@ERROR <> 0 
		GOTO EventHandler	


	Update tper 
	Set nvcFirstName = Replace(  nvcFirstName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcFirstName = Replace(  nvcFirstName , N'ي' , N'ی' ) 

	Update tper 
	Set nvcSurName = Replace(  nvcSurName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcSurName = Replace(  nvcSurName , N'ي' , N'ی' ) 


COMMIT TRANSACTION



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER    Procedure dbo.Insert_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,  
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,
	@TotalRemainingAmount INT ,   
	@Code Bigint out  
)  

as  

Begin Tran  

DECLARE @Branch INT 
 Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode ) --AND (Branch = @Branch ) )  
    end   
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId) <> 0 -- AND Branch = @Branch ) <> 0   
  Goto ErrHandler   

--Set @Code = (Select  isnull(Max(Code),0) + 1 from tSupplier where code > 0)  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

insert Into dbo.tSupplier  
(   
	--Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	State ,  
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	Discount,   
	[Description],   
	[Date],   
	[Time],   
	[User], 
	TotalRemainingAmount , 
	Branch  
)  
values  
(   
	--@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@State,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@Discount,   
	@Description,   
	@Date,   
	@Time,   
	@User , 
	@TotalRemainingAmount , 
	@Branch  
)  
if @@Error <> 0   
 goto ErrHandler  

Set @Code = @@Identity  
 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] )   

update tSupplier set address=dbo.addressedit(address) where code=@code  -- AND Branch = @Branch
   

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

Commit Tran
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code  
--Select @Code


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   Procedure dbo.Update_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint , 
	@TotalRemainingAmount  INT , 
	@Updated Bigint out  

)  

as  

Begin Tran  
--IF @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode )--  AND (Branch = @Branch ) )  
    end  
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code ) <> 0 --  AND (Branch = @Branch ) ) <> 0    
  Goto ErrHandler   
 else  

  Update dbo.tSupplier     
   Set MembershipId = @MembershipId   

  Where MasterCode = @Code   --AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tSupplier  

 Set MembershipId = @MembershipId ,  
 MasterCode  = @MasterCode ,    
 Owner = @Owner ,  
 Name = @Name ,  
 Family = @Family ,  
 Sex = @Sex ,  
 WorkName = @WorkName ,   
 InternalNo = @InternalNo ,  
 Unit = @Unit ,  
 State = @State ,  
 City = @City ,  
 ActKind = @ActKind ,  
 ActDeAct = @ActDeAct ,  
 Prefix = @Prefix ,  
 Address = @Address ,  
 PostalCode = @PostalCode ,  
 Tel1 = @Tel1 ,  
 Tel2 = @Tel2 ,  
 Tel3 = @Tel3 ,  
 Tel4 = @Tel4 ,  
 Mobile = @Mobile ,  
 Fax = @Fax ,  
 Email = @Email ,  
 Flour = @Flour ,  
 Discount = @Discount ,  
 [Description] = @Description ,  
 [Date] = @Date ,  
 [Time] = @Time ,  
 [User] = @User ,
 TotalRemainingAmount = @TotalRemainingAmount 
Where Code = @Code  -- AND (Branch = @Branch )  


if @@Error <> 0   
 goto ErrHandler  

 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] or dbo.tSupplier.[Branch] is Null)  

update tSupplier set address=dbo.addressedit(address) where code=@code  --AND (Branch = @Branch )

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 


Commit Tran   
Set @Updated = @Code  
return @Updated  

ErrHandler:  
RollBack Tran  
Set @Updated = 0  
return @Updated



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER     Procedure dbo.Insert_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int, 
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Code Bigint out 

)  

as  

Begin Tran  

--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  
if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  )  --AND (Branch = @Branch )
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode   )--AND (Branch = @Branch )
 end   
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId ) <> 0 --AND Branch = @Branch)   
  Goto ErrHandler   

Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  Branch = @Branch )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

if @nvcRFID = N''  
  SET @nvcRFID=N'-999'  

insert Into dbo.tCust  
(   
	Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Assansor,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	CarryFee,   
	PaykFee,   
	Distance,   
	Credit,   
	Discount,   
	BuyState,   
	[Description],   
	[Date],   
	[Time],   
	[User],  
	FamilyNo ,  
	Member ,  
	State ,  
	Central ,  
	Branch,  
	nvcRFID,  
	sellprice ,
	EconomicCode ,
	nvcBirthDate ,
	TotalRemainingAmount
	
)  
values  
(   
	@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Assansor,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@CarryFee,   
	@PaykFee,   
	@Distance,   
	@Credit,   
	@Discount,   
	@BuyState,   
	@Description,   
	@Date,   
	@Time,   
	@User ,  
	@FamilyNo ,  
	@Member ,  
	@State ,  
	@Central ,  
	@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate ,
	@TotalRemainingAmount
	
)  
if @@Error <> 0   
 goto ErrHandler  

--Set @Code = @@Identity  
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
  and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address)  
 , nvcRFID=CAST(Branch AS NVARCHAR(1))+CAST(Code AS NVARCHAR(8))  
  where code=@code  AND Branch = @Branch 


Update [tCust]
Set [Name] = Replace(  [Name] , N'ك' , N'ک'  ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي' , N'ی' ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ي' , N'ی' ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ي' , N'ی' ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ي' , N'ی' ) 


Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   Procedure dbo.Update_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint ,  
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Updated Bigint out  

)  

as  

Begin Tran  
--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  ) --AND (Branch = @Branch ) )  
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode )   --AND (Branch = @Branch ) )  
 end  
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code  ) <> 0  -- AND (Branch = @Branch )    
  Goto ErrHandler   
 else  

  Update dbo.tCust     
   Set MembershipId = @MembershipId   

   Where MasterCode = @Code   AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tCust  

 Set MembershipId = @MembershipId ,  
	MasterCode  = @MasterCode ,    
	Owner = @Owner ,  
	Name = @Name ,  
	Family = @Family ,  
	Sex = @Sex ,  
	WorkName = @WorkName ,   
	InternalNo = @InternalNo ,  
	Unit = @Unit ,  
	City = @City ,  
	ActKind = @ActKind ,  
	ActDeAct = @ActDeAct ,  
	Prefix = @Prefix ,  
	Assansor = @Assansor ,  
	Address = @Address ,  
	PostalCode = @PostalCode ,  
	Tel1 = @Tel1 ,  
	Tel2 = @Tel2 ,  
	Tel3 = @Tel3 ,  
	Tel4 = @Tel4 ,  
	Mobile = @Mobile ,  
	Fax = @Fax ,  
	Email = @Email ,  
	Flour = @Flour ,  
	CarryFee = @CarryFee ,  
	PaykFee = @PaykFee ,  
	Distance = @Distance ,  
	Credit = @Credit ,  
	Discount = @Discount ,  
	BuyState = @BuyState ,  
	[Description] = @Description ,  
	[Date] = @Date ,  
	[Time] = @Time ,  
	[User] = @User ,  
	FamilyNo = @FamilyNo ,  
	Member = @Member ,  
	State = @State ,  
	Central = @Central,  
	Sellprice=@Sellprice  ,
	EconomicCode = @EconomicCode ,
	nvcRFID = @nvcRFID ,
	nvcBirthDate = @nvcBirthDate ,
	TotalRemainingAmount = @TotalRemainingAmount
	
Where Code = @Code   AND (Branch = @Branch )   

if @@Error <> 0   
 goto ErrHandler  


Set @Updated = @Code   
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
 and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address) where code=@code  AND Branch = @Branch 
 
Update [tCust]
Set [Name] = Replace(  [Name] , N'ك' , N'ک'  ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي' , N'ی' ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ي' , N'ی' ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ي' , N'ی' ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ي' , N'ی' ) 


Commit Tran  
return @Updated  

ErrHandler:  
RollBack Tran  
return -1



GO
SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

--Script_V26_16_Fix10_9_FarsiCodePage
--اصلاح ی و ک فارسی در دیتا بیس برای کالاها و مشتریان و تامین کنندگان و پرسنل
--در برنامه فروش در قسمت جستجوی کالا اگر اصلاح ک و ی فارسی را بزنیم همه دیتاهای قبلی درست می شود
--93/10/29


ALTER   Proc Get_Customer_Name
@ActDeact int ,
@Name Nvarchar(50)     
as    

Set @Name = Replace(  @Name  , N'ك', N'ک'  ) 
Set @Name = Replace(  @Name  , N'ي' , N'ی' )

Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where  CHARINDEX ( @Name , [Name] ) > 0 and actdeact <> @ActDeact  -- AND Branch = @Branch
AND vw_Get_Cust.Code <> -1
Order By [Name]



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  Proc Get_Supplier_Name
@ActDeact int ,
@Name Nvarchar(50) 
as

Set @Name = Replace(  @Name  , N'ك', N'ک'  ) 
Set @Name = Replace(  @Name  , N'ي' , N'ی' )

Select * from dbo.vw_Get_Supplier where  CHARINDEX ( @Name , [Name] ) > 0 and actdeact <> @ActDeact  Order By [Name]



GO

