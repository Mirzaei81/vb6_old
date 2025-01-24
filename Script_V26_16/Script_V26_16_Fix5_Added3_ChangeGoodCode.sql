
--Script_V26_16_Fix5_Added3_ChangeGoodCode.sql
--93/04/23

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER    PROCEDURE dbo.UpdatetGood
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
	Set [Name] = Replace(  [Name]  , N'˜' , N'ß' ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn]  , N'˜' , N'ß' ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'˜' , N'ß' ) 


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


ALTER  PROCEDURE ChangeLevel1
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
	IF @@ERROR <> 0 
		GOTO ErrorHandler
	UPDATE dbo.tGoodLevel2 
		SET Level1Code = @NewLevel1 WHERE Level1Code = @OldLevel1
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

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   PROCEDURE ChangeLevel2
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
	IF @@ERROR <> 0 
		GOTO ErrorHandler	

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tUsePercent_tGood1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
		ALTER TABLE [dbo].[tUsePercent] DROP CONSTRAINT [FK_tUsePercent_tGood1]
	IF @@ERROR <> 0 
		GOTO ErrorHandler	

	UPDATE [tUsePercent] SET GoodFirstCode = CAST(CAST(@NewLevel2 AS VARCHAR(4)) + SUBSTRING(CAST(GoodFirstCode AS VARCHAR(8)),5 ,4) AS INT) 
	WHERE GoodFirstCode >= @OldLevel2 * 10000 AND GoodFirstCode < (@OldLevel2+1) * 10000
	IF @@ERROR <> 0 
		GOTO ErrorHandler

	UPDATE tGood SET code = CAST(CAST(Level2 AS VARCHAR(4)) + SUBSTRING(CAST(Code AS VARCHAR(8)),5 ,4) AS INT)
	WHERE Level1 = @Level1 AND Level2 = @NewLevel2 
	IF @@ERROR <> 0 
		GOTO ErrorHandler
			
	ALTER TABLE [dbo].[tUsePercent]  WITH CHECK ADD  CONSTRAINT [FK_tUsePercent_tGood1] FOREIGN KEY([GoodFirstCode])
		REFERENCES [dbo].[tGood] ([Code])
	IF @@ERROR <> 0 
		GOTO ErrorHandler	


COMMIT TRAN
SET @Update = 1
RETURN @Update

ErrorHandler:
ROLLBACK TRAN
RETURN @Update


GO


