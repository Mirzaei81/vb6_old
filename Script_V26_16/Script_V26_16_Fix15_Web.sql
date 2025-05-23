
--Script_V26_16_Fix15_Web
--  اضافه شدن حالت فروش اینترنتی به فرم های پیک و انبار
--در حالت فروش اینترنتی
-- StationId = -1 ایستگاه وب
--InventoryNo = -1  انباراینترنتی
--Ppno = -1 , UserId = -1  کاربروب
-- امکان کارکرد با اینترفیس فروش آنلاین آریا 
--95/02/30

IF NOT EXISTS(SELECT * FROM tblPub_Script2 WHERE [Version] = 26 AND Script = 16 AND FixNumber = 15 )

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
			  15
			)
GO


--Script_V26_16_Web
--فروش اینترنتی 

--92/10/21

-- ----------------------------
-- Add Column BitWebTransfer for tCust
-- ----------------------------
IF COL_LENGTH('tCust','BitWebTransfer') IS NULL
BEGIN

	ALTER TABLE dbo.tCust
	ADD BitWebTransfer [bit] NULL DEFAULT(0)
END

GO

-- ----------------------------
-- Add Column BitWebTransfer for tGood
-- ----------------------------
IF COL_LENGTH('tGood','BitWebTransfer') IS NULL
BEGIN

	ALTER TABLE dbo.tGood
	ADD BitWebTransfer [bit] NULL DEFAULT(0)
END
GO

IF COL_LENGTH('tGood','BitWebShow') IS NULL
BEGIN

	ALTER TABLE dbo.tGood 	ADD BitWebShow [bit] NOT NULL DEFAULT(0)
END
GO
	UPDATE tGood SET BitWebShow = 1 WHERE GoodType = 2 OR GoodType = 3 

GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_WebtFacD_WebtFacM]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tbl_Web_tFacD] DROP CONSTRAINT [FK_WebtFacD_WebtFacM]
GO

-- --------------------------------------
-- Create Table tbl_Web_tFacM And tbl_Web_tFacD
-- --------------------------------------
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Web_tFacM]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE [dbo].[tbl_Web_tFacM]
Go

CREATE TABLE [dbo].[tbl_Web_tFacM](
			[intSerialNo] [bigint] IDENTITY(1,1) NOT NULL,
			[Branch] [int] NOT NULL,
			[SaveInArya] [bit] NOT NULL,
			[DiscountTotal] [money] NOT NULL,
			[FreigtTotal] [money] NOT NULL,
			[NvcDescription] [nvarchar](150) NULL,
			[intCode] [int] NOT NULL,
			[Tempaddress] [nvarchar](255) NULL,
			[nvcName] [nvarchar](50) NOT NULL,
			[TelNo] [varchar](10) NULL,
			[MobileNO] [varchar](10) NULL,
			[nvcAddress] [nvarchar](255) NULL,
			[SumPrice] [money] NOT NULL,
			[nvcDate] [nvarchar](50) NULL,
			[nvcTime] [nvarchar](50) NULL,
		 CONSTRAINT [PK_tbl_Web_tFacM] PRIMARY KEY CLUSTERED 
		(
			[intSerialNo] ASC,
			[Branch] ASC
		)--WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
		) ON [PRIMARY]
		GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbl_Web_tFacD]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
DROP TABLE [dbo].[tbl_Web_tFacD]
GO

CREATE TABLE [dbo].[tbl_Web_tFacD](
			[intSerialNo] [bigint] NOT NULL,
			[Branch] [int] NOT NULL,
			[intRow] [int] NOT NULL,
			[amount] [float] NOT NULL,
			[GoodCode] [int] NOT NULL,
			[Feeunit] [float] NOT NULL,
		 CONSTRAINT [PK_tbl_Web_tFacD] PRIMARY KEY CLUSTERED 
		(
			[intSerialNo] ASC,
			[Branch] ASC,
			[intRow] ASC,
			[GoodCode] ASC
		)--WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
		) ON [PRIMARY]
		GO

		Alter TABLE [dbo].[tbl_Web_tFacD]  WITH NOCHECK ADD  CONSTRAINT [FK_WebtFacD_WebtFacM] FOREIGN KEY([intSerialNo], [Branch])
		REFERENCES [dbo].[tbl_Web_tFacM] ([intSerialNo], [Branch])
		ON UPDATE CASCADE
		ON DELETE CASCADE
		GO
		
		ALTER TABLE [dbo].[tbl_Web_tFacD] CHECK CONSTRAINT [FK_WebtFacD_webtFacM]
		GO
-- -------------------------
-- Create Procedure
-- -------------------------


ALTER    Procedure dbo.Update_Cust  
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

UPDATE dbo.tCust SET BitWebTransfer = 0 WHERE Code = @Code 

Commit Tran  
return @Updated  

ErrHandler:  
RollBack Tran  
return -1



GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_GetCustomer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE dbo.sp_Web_GetCustomer
GO

Create PROCEDURE dbo.sp_Web_GetCustomer
AS

SELECT tCust.* , ISNULL(NULLIF(WorkName ,''), Name + ' ' + Family) AS CustomerName FROM dbo.tCust WHERE Code > 0 AND BitWebTransfer is null or BitWebTransfer = 0
GO
--


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_GetGood]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE dbo.sp_Web_GetGood
GO

Create PROCEDURE dbo.sp_Web_GetGood
AS
SELECT * FROM dbo.tGood WHERE (GoodType = 2 OR GoodType = 3) AND  BitWebTransfer is null or BitWebTransfer = 0
GO
--

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_GetInvoiceInfoFromWebtfacM]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE dbo.sp_Web_GetInvoiceInfoFromWebtfacM
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE  PROCEDURE sp_Web_GetInvoiceInfoFromWebtfacM
AS

DECLARE @nvcDate NVARCHAR(8)
SET @nvcDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())

BEGIN
	SELECT [intSerialNo]
      ,[Branch]
      ,[SaveInArya]
      ,[DiscountTotal]
      ,[FreigtTotal]
      ,[NvcDescription]
      ,[intCode]
      ,[Tempaddress]
      ,[nvcName]
      ,[TelNo]
      ,[MobileNO]
      ,[nvcAddress]
      ,[SumPrice]
      ,[nvcDate]
      ,[nvcTime]
  FROM [tbl_Web_tFacM] 
	WHERE nvcDate = @nvcDate

END


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_SetCustomerBitTransfer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE dbo.[sp_Web_SetCustomerBitTransfer]

GO

CREATE PROCEDURE [dbo].[sp_Web_SetCustomerBitTransfer] ( @CustCode INT )
AS
    UPDATE  dbo.tCust
    SET     BitWebTransfer = 1
    WHERE   Code = @CustCode
GO
--


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_SetGoodBitTransfer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE dbo.[sp_Web_SetGoodBitTransfer]

GO

CREATE PROCEDURE [dbo].[sp_Web_SetGoodBitTransfer] ( @GoodCode INT )
AS
    UPDATE  dbo.tGood
    SET     BitWebTransfer = 1
    WHERE   Code = @GoodCode
GO
--


if  exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetFactorDetail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [GetFactorDetail]
go 

CREATE   PROCEDURE [dbo].[GetFactorDetail]
    (
      @intSerialNo BIGINT ,
      @Branch INT = 0
    )
AS
    IF @Branch = 0
        SET @Branch = dbo.Get_Current_Branch()
 
    SELECT  tbl_Web_tFacD.Amount ,
            tGood.Name ,
            tbl_Web_tFacD.Feeunit ,
            GoodCode
    FROM    tGood
            INNER JOIN tbl_Web_tFacD ON tGood.Code = tbl_Web_tFacD.GoodCode
    WHERE   intSerialNo = @intSerialNo
            AND Branch = @Branch
 GO
 --
 
 
 if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_InsertFactorMasterDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
 DROP PROCEDURE [dbo].[sp_Web_InsertFactorMasterDetails]
 GO
 

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
CREATE  PROCEDURE [dbo].[sp_Web_InsertFactorMasterDetails]  (  
            
			@Branch [int],
			@SaveInArya [bit],
			@DiscountTotal [float],
			@FreigtTotal [float] ,
			@DetailString [varchar](4000),
			@NvcDescription [nvarchar](150),
			@intCode [int], 
			@Tempaddress [nvarchar](255),
			@nvcName [nvarchar](50),
			@TelNo [varchar](10),
			@MobileNO [varchar](10),
			@nvcAddress [nvarchar](255),
			@SumPrice [float],
			@nvcDate [NVARCHAR](10),
			@nvcTime [NVARCHAR](5),
			@intSerialNo bigint  OUT
			)  

AS 
	--DECLARE 	@Branch [int]
	--	DECLARE	@SaveInArya [bit]
	--	DECLARE	@DiscountTotal [float]
	--DECLARE		@FreigtTotal [float] 
	--DECLARE		@DetailString [varchar](max)
	--DECLARE		@NvcDescription [nvarchar](150)
	--DECLARE		@intCode [int]
	--DECLARE		@Tempaddress [nvarchar](255)
	--DECLARE		@nvcName [nvarchar](50)
	--DECLARE		@TelNo [varchar](10)
	--DECLARE		@MobileNO [varchar](10)
	--DECLARE		@nvcAddress [nvarchar](255)
	--DECLARE		@SumPrice [float]
	--DECLARE		@nvcDate [NVARCHAR](50)
	--DECLARE		@nvcTime [NVARCHAR](50)

 -- SET  @Branch = 1 -- int
 --   SET @SaveInArya = 0 -- bit
 --  SET  @DiscountTotal = 0.0 -- float
 -- SET   @FreigtTotal = 0.0 -- float
 -- SET   @DetailString = '20;440;220000;/1;1101002202;190000;/3;1101002203;180000;/'-- varchar(max)
 --  SET  @NvcDescription = N'' -- nvarchar(150)
 -- SET   @intCode = 100, -- int
 -- SET   @Tempaddress = N'Tehran' -- nvarchar(255)
 -- SET   @nvcName = N' aaa' -- nvarchar(50)
 -- SET   @TelNo = ' ' -- varchar(10)
 -- SET   @MobileNO = '' -- varchar(10)
 -- SET   @nvcAddress = N'' -- nvarchar(255)
 -- SET   @SumPrice = 0.0 -- float
 -- SET   @nvcDate = N'' -- nvarchar(50)
 -- SET   @nvcTime = N'' -- nvarchar(50)

BEGIN 

    -- Check @intCode  -> IF Not Exist ->> Add New Customer To Tcust
	   
    DECLARE @CustCode BIGINT 

    IF @intCode = 0 OR ( NOT EXISTS ( SELECT * FROM dbo.tCust WHERE Code = @intCode AND Branch = @Branch ) )
        BEGIN
	   
            DECLARE @MembershipId BIGINT 
            DECLARE @User INT 

            --SELECT TOP 1 @User = UID FROM    dbo.tUser
			SET @User = -1
            SELECT  @MembershipId = ISNULL(MAX(MembershipId), 0) + 1 FROM    dbo.tCust
	
            EXEC dbo.Insert_CustomerFast @MembershipId = @MembershipId,
                @Name = '', @Family = nvcName, @Address = @nvcAddress,
                @Tel1 = @TelNo, @Mobile = @MobileNO, @Description = '',
                @User = @User, @Code = @CustCode OUT
   
        END
    ELSE
        BEGIN
            SET @CustCode = @intCode
        END  

    select @nvcTime = dbo.setTimeFormat(getdate())      
    
    SET @nvcDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      


INSERT INTO tbl_Web_tFacM(
		
		[Branch],
		[SaveInArya],
		[DiscountTotal],
		[FreigtTotal],
		[NvcDescription],
		[intCode],
		[Tempaddress],
		[nvcName],
		[TelNo],
		[MobileNO],
		[nvcAddress],
		[SumPrice],
		[nvcDate],
		[nvcTime]
)
VALUES (
		
		@Branch,
		@SaveInArya,
		@DiscountTotal,
		@FreigtTotal,
		@NvcDescription,
		CAST(@CustCode AS INT), 
		@Tempaddress,
		@nvcName,
		@TelNo,
		@MobileNO,
		@nvcAddress,
		@SumPrice,
		@nvcDate,
		@nvcTime

)
set @intSerialNo =@@identity


	declare @nvcMainString nvarchar(4000)

--declare @DetailString nvarchar(4000)

--set @DetailString='20;440;220000;/1;1101002202;190000;/3;1101002203;180000;/'
	DECLARE @intDelimiterPosField  INT
    DECLARE @intDelimiterPosRecord INT
	Declare @ReturnTable TABLE(row int IDENTITY (1, 1) NOT NULL,Amount FLOAT, GoodCode INT , FeeUnit Float )
    
 	DECLARE @Amount FLOAT
    DECLARE @GoodCode INT
    DECLARE @FeeUnit Float
 
    DECLARE @TempTable Table (nvcMainString nText)

    insert into @TempTable values (@DetailString)
   
    SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
    SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
    --if @intDelimiterPosRecord <> 0-- and @intDelimiterPosField < @intDelimiterPosRecord
    WHILE @intDelimiterPosRecord <> 0
		BEGIN 
--**********
        	SET @Amount = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS FLOAT)  from @TempTable )
        	SET @Amount =  ROUND(CAST(@Amount AS DECIMAL(15,3)),3)
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--SET @nvcMainString = (Select nvcMainString from @TempTable)
--PRINT @nvcMainString
--PRINT @Amount
--PRINT @intDelimiterPosField
--PRINT @intDelimiterPosRecord
--**********
			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
			SET @GoodCode = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--SET @nvcMainString = (Select nvcMainString from @TempTable)
--PRINT @nvcMainString
--PRINT @GoodCode
--PRINT @intDelimiterPosField
--PRINT @intDelimiterPosRecord
--**********
			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
			SET @FeeUnit = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Float)  from @TempTable )
      	   Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosRecord + 1, DataLength(nvcMainString) - @intDelimiterPosRecord  ) from @TempTable )
--SET @nvcMainString = (Select nvcMainString from @TempTable)
--PRINT @nvcMainString
--PRINT @FeeUnit
--PRINT @intDelimiterPosField
--PRINT @intDelimiterPosRecord

		INSERT INTO @ReturnTable(Amount , GoodCode , FeeUnit) VALUES(@Amount,@GoodCode,@FeeUnit)
        SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable )
        SET @intDelimiterPosRecord = ( Select patindex('%/%' , nvcMainString)  from @TempTable )
	END 

INSERT INTO dbo.tbl_Web_tFacD (

		[intSerialNo],
		[Branch],
		[intRow],
		[amount],
		[GoodCode],
		[Feeunit] )
select  @intSerialNo , @Branch , row , Amount , GoodCode , FeeUnit from @ReturnTable

END 
  return @intSerialNO
  
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetMojodiWithGoodCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[GetMojodiWithGoodCode]
GO

CREATE PROCEDURE [dbo].[GetMojodiWithGoodCode]
AS

DECLARE @AccountYear INT 
SET @AccountYear = dbo.Get_AccountYear()

    select GoodCode , Mojodi from tInventory_Good 
    INNER JOIN dbo.tGood ON dbo.tInventory_Good.GoodCode = tgood.Code
    WHERE  BitWebShow = 1 AND (tgood.GoodType = 2 OR tgood.GoodType = 3) AND AccountYear = @AccountYear and InventoryNo = 1
    
    
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER      PROCEDURE dbo.UpdatetGood
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

    UPDATE tgood SET BitWebTransfer = 0 WHERE code = @Code

COMMIT TRANSACTION


Return @Result


EventHandler:
    ROLLBACK TRAN
    Set @Result = 0
    RETURN @Result



GO


if NOT exists (select * from dbo.tPer where pPno = -1)
 
INSERT INTO dbo.tPer
        ( pPno ,
          PersonnelNumber ,
          nvcFirstName ,
          nvcSurName ,
          Gender ,
          IdNumber ,
          Job ,
          InsuranceNo ,
          Address ,
          Tel ,
          Date ,
          Time ,
          [User] ,
          Branch ,
          Tafsili ,
          MaxCredit ,
          ActDeAct
        )
VALUES  ( -1 ,
		  N'-1' , -- PersonnelNumber - nvarchar(50)
          N'کاربر' , -- nvcFirstName - nvarchar(50)
          N'وب' , -- nvcSurName - nvarchar(50)
          1 , -- Gender - udt_Gender
          N'-1' , -- IdNumber - nvarchar(50)
          1 , -- Job - int
          N' ' , -- InsuranceNo - nvarchar(50)
          N' ' , -- Address - nvarchar(300)
          N' ' , -- Tel - nvarchar(30)
          N' ' , -- Date - nvarchar(50)
          N' ' , -- Time - nvarchar(50)
          1 , -- User - int
          1 , -- Branch - int
          NULL  , -- Tafsili - int
          0 , -- MaxCredit - int
          1  -- ActDeAct - bit
        )

GO

--SELECT * FROM dbo.tPer
--go 


if NOT exists (select * from dbo.tUser where [UID] = -1)
   INSERT INTO dbo.tUser
        ( UID ,
          Username ,
          PassWord ,
          nvcHint ,
          nvcAnswer ,
          pPno ,
          AddUser ,
          intAccessLevel ,
          Branch ,
          CountRePrint ,
          CountInvoicePrint ,
          CountInvoiceEditable ,
          CountInvoiceRefferable
        )
VALUES  ( -1 ,
		  N'web' , -- Username - nvarchar(50)
          N'web' , -- PassWord - nvarchar(50)
          N' ' , -- nvcHint - nvarchar(200)
          N' ' , -- nvcAnswer - nvarchar(50)
          -1 , -- pPno - int
          1 , -- AddUser - int
          3 , -- intAccessLevel - int
          1 , -- Branch - int
          10 , -- CountRePrint - int
          10 , -- CountInvoicePrint - int
          10 , -- CountInvoiceEditable - int
          10  -- CountInvoiceRefferable - int
        )
GO


--SELECT * FROM dbo.tUser
--go 


if NOT exists (select * from dbo.tInventory where InventoryNo = -1)
  INSERT INTO dbo.tInventory
        ( InventoryNo ,
          MasterCode ,
          Type ,
          Description ,
          LatinDescription ,
          Active ,
          Branch ,
          Tafsili
        )
VALUES  ( -1 ,
		  1 , -- MasterCode - int
          NULL , -- Type - int
          N'انبار وب' , -- Description - nvarchar(50)
          N'Web Inventory' , -- LatinDescription - nvarchar(50)
          1 , -- Active - int
          1 , -- Branch - int
          NULL   -- Tafsili - int
        )
go 

--SELECT * FROM dbo.tInventory
--go 



--ِاضافه کردن کالاها  به انبار وب 
insert into .dbo.[tInventory_Good](GoodCode,InventoryNo,Branch ,AccountYear)
   select Code,-1,dbo.Get_Current_Branch() , dbo.get_AccountYear()  
	from dbo.[tGood]Where Code  Not In (Select GoodCode From tInventory_Good Where InventoryNo = -1 
			And Branch = dbo.Get_Current_Branch() And AccountYear = dbo.get_AccountYear())
	-- And GoodType = 4


go


--SELECT * FROM dbo.tStations
--go
if NOT exists (select * from dbo.tStations where StationID = -1)
INSERT INTO dbo.tStations
        ( StationID ,
          PortCode ,
          Description ,
          CashNo ,
          IsActive ,
          IP ,
          Dir ,
          Machine_Name ,
          StationType ,
          Branch ,
          TemporaryNo
        )
VALUES  ( -1 , -- StationID - int
          NULL  , -- PortCode - int
          N'WebStation' , -- Description - nvarchar(50)
          1 , -- CashNo - int
          1 , -- IsActive - bit
          N'' , -- IP - nvarchar(50)
          N'' , -- Dir - nvarchar(50)
          N'' , -- Machine_Name - nvarchar(50)
          2 , -- StationType - int
          dbo.Get_Current_Branch() , -- Branch - int
          1  -- TemporaryNo - bit
        )
GO


--SELECT * FROM dbo.tOrderType
--go
if NOT exists (select * from dbo.tOrderType where Code = 3)
INSERT INTO dbo.tOrderType
        ( Code ,
          Description ,
          LatinDescription
        )
VALUES  ( 3 , -- Code - int
          N'اینترنتی' , -- Description - nvarchar(50)
          N'by Web'  -- LatinDescription - nvarchar(50)
        )

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WS_Get_All_Goods]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[WS_Get_All_Goods]
GO


CREATE PROCEDURE [dbo].[WS_Get_All_Goods]

AS 

SELECT * FROM dbo.tGood WHERE GoodType = 2 OR GoodType = 3


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WS_Get_Mojodi_ByCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[WS_Get_Mojodi_ByCode]
GO


CREATE PROCEDURE [dbo].[WS_Get_Mojodi_ByCode]

@GoodCode INT ,
@Mojodi INT OUT 

AS 

SELECT @Mojodi = Mojodi FROM dbo.tInventory_Good 
	WHERE InventoryNo = -1 
		AND Branch = dbo.Get_Current_Branch() 
		And AccountYear = dbo.get_AccountYear()
		AND GoodCode = @GoodCode
	SET @Mojodi = ISNULL(@Mojodi , 0)

GO



--if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WS_InsertFactorMasterDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
--drop procedure [dbo].[WS_InsertFactorMasterDetails]
--GO


--CREATE PROCEDURE [dbo].[WS_InsertFactorMasterDetails]  (      

--        @DiscountTotal FLOAT ,      
--        @CarryFeeTotal FLOAT ,      
--        @ServiceTotal FLOAT ,      
--        @PackingTotal FLOAT ,      
--        @DetailsString nText,      
--        @NvcDescription Nvarchar(150) = Null ,      
--        @TempAddress Nvarchar(255) = '',  
--        @intserialNo INT OUT      
--     )      

--AS      

--DECLARE @SumPrice  float      
--Set @SumPrice = 0      

--DECLARE @proper_time nvarchar(5)      

--DECLARE @AccountYear SMALLINT
--Set @AccountYear = dbo.get_AccountYear() 
    
--BEGIN TRAN      

--    DECLARE @MasterServePlace INT      
--    DECLARE @newtime nvarchar(5)      
--    select @newtime=dbo.setTimeFormat(getdate())      
--    SET  @MasterServePlace = 16      

------------------------------------------Date From Server-----------------------------------------------------------------      
--	DECLARE @Date NVARCHAR(8)    
--	SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      

--	Declare @intBranch  int      
--	select @intBranch = dbo.Get_Current_Branch()      

--	DECLARE @No1  INT     
--	SET @NO1 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=11  And Branch =  @intBranch AND AccountYear = @AccountYear)      

--	Declare @ShiftNo int      
--	SET @ShiftNo= dbo.Get_Shift(GETDATE())      

--	DECLARE @TempNo INT 
--	SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=11  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      


--     INSERT INTO tFacM (      
--		[No] ,      
--		[Date] ,      
--		RegDate ,      
--		Status ,      
--		Customer ,      
--		SumPrice ,      
--		OrderType ,      
--		ServePlace ,      
--		StationId ,      
--		ServiceTotal ,      
--		Recursive ,      
--		CarryFeeTotal ,      
--		PackingTotal ,      
--		DiscountTotal ,      
--		[Time] ,      
--		[User] ,      
--		TableNo ,      
--		shiftNo ,      
--		incharge,      
--		owner ,      
--		FacPayment ,       
--		Balance ,       
--		Branch,      
--		AccountYear ,      
--		NvcDescription,      
--		TempAddress ,
--		GuestNo ,
--		TempNo    
		
-- )      
--     Values       

--(      
--                @NO1 ,      
--                @Date ,      
--                dbo.Shamsi(GETDATE()) ,      
--                11,      
--                -2 ,      
--                @SumPrice ,      
--                3 ,      
--                32 ,      
--                -1 ,      
--                @ServiceTotal ,      
--                0 ,      
--                @CarryFeeTotal ,      
--                @PackingTotal ,      
--                @DiscountTotal ,      
--                @newtime,      
--                -1 ,      
--                NULL ,      
--                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
--                NULL  ,      
--                NULL  ,      
--                0 ,      
--                0 ,      
--		@intBranch , --dbo.Get_Current_Branch(),      
--		@AccountYear ,      
--		@NvcDescription,      
--		@TempAddress,
--		0,
--		@TempNo  
-- )      
--    SET @intserialNo=@@IDENTITY      
--     IF @@ERROR <>0      
--        GoTo EventHandler       

------------------------------------Fill Details Factor  --------------------------------------------------------------      
-- exec InsertFactorDetail @DetailsString , @intserialNo , 0, -2 , @intBranch      

--     IF @@ERROR <>0      
--        GoTo EventHandler      
----------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------      

------------------------------------Total SumPrice Calculate  --------------------------------------------------------------      
--DECLARE @DiscountD INT 
--Set @DiscountD = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * discount/100 ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
--Set @SumPrice = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * (1 - discount/100) ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

--     IF @@ERROR <>0      
--        GoTo EventHandler      
--DECLARE @TaxTotal FLOAT  
--SET @TaxTotal = 0
--DECLARE @ValueGoodsTax FLOAT
--SET @ValueGoodsTax = 0
--	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
--	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
--	              WHERE     intSerialNo = @intSerialNo
--	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
--	            )  
--SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
--IF @@ERROR <> 0 
--GOTO EventHandler

--DECLARE @DutyTotal INT 
--SET @DutyTotal = 0
--DECLARE @ValueGoodsDuty FLOAT
--SET @ValueGoodsDuty = 0
--	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
--	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
--	              WHERE     intSerialNo = @intSerialNo
--	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
--	            )  
--SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
--IF @@ERROR <> 0 
--GOTO EventHandler

------------------------------------ServiceRate Calculate  --------------------------------------------------------------      

-- If @ServiceTotal <> 0      
--       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)      
--     IF @@ERROR <>0      
--        GoTo EventHandler       
------------------------------------Round Sumprice  --------------------------------------------------------------      
--  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
--  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
--    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal     

--    Declare @Remain INT
--    SET @Remain = 0  
--    Set @Remain = dbo.RoundSumPrice(@SumPrice )         
--    Set @SumPrice = @SumPrice - @Remain      
--    Set @DiscountTotal = @DiscountTotal + @Remain    
-----select @Remain as remain      
------------------------------------Net Price Update  --------------------------------------------------------------      

--Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch      
--Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch       
--Update tFacm Set DiscountTotal = @DiscountTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
--Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch       
--Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
--Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

---------------------------------------Fill Detail Cash , ..........--------------------------------------------------      
------------------------------History---------------------------      

--Exec InsertHistory  @No1, 11 , -1 , 1 , @AccountYear , @intBranch      

------------------------------Cash ---------------------------      

--------------------------Mojodi Control Online--------------------------------------------      

--Exec InsertMojodiCalculate 11 ,  @intserialNo , @AccountYear , @intBranch      
--IF @@ERROR <>0      
-- GoTo EventHandler      

--------------------------------------------Update Balance After Recived----------------------------

--COMMIT TRAN


--Return @intserialNo      

--EventHandler:      

--    ROLLBACK TRAN      
--    SET @intserialNo = -1      

--    RETURN @intserialNo



--GO



--if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WS_VoidFactor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
--drop procedure [dbo].WS_VoidFactor
--GO

--CREATE PROCEDURE WS_VoidFactor
--@intserialNo INT 

--AS 

--UPDATE dbo.tFacM
--	SET Recursive = 1 WHERE intSerialNo = @intserialNo AND Branch = dbo.Get_Current_Branch()

--DECLARE @AccountYear SMALLINT  
--DECLARE @Branch INT 
--SET @AccountYear = dbo.Get_AccountYear()
--SET @Branch = dbo.Get_Current_Branch()
--Exec DeleteMojodiCalculate 11 , @intserialNo  ,  1 , @AccountYear , @Branch  


--GO 




--if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WS_EditFactorMasterDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
--drop procedure [dbo].[WS_EditFactorMasterDetails]
--GO


--CREATE  PROCEDURE [dbo].[WS_EditFactorMasterDetails]  (  


--	@DiscountTotal Float ,  
--	@CarryFeeTotal Float ,  
--	@ServiceTotal  Float ,  
--	@PackingTotal  Float ,  
--	@DetailsString  nText,  
--	@NvcDescription Nvarchar(150) = Null ,  
--	@TempAddress Nvarchar(255) = '', 
--	@intSerialNo  INT ,
--	@Result INT OUT 
--  )  


--AS  
--	DECLARE @SumPrice BIGINT  
--	SET @Sumprice = 0  

--	DECLARE  @AccountYear int 
--	Set @AccountYear = dbo.Get_AccountYear()

--	Declare @intBranch  int  
--	SET  @intBranch =  dbo.Get_Current_Branch() 

--	DECLARE @No INT 
--	SELECT @No = No FROM dbo.tFacM WHERE intSerialNo = @intSerialNo AND Branch = @intBranch
	
--BEGIN TRANSACTION  

--       INSERT INTO tRepFacEditM (Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance , OrderType,  
--                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate, AccountYear , TaxTotal , DutyTotal  )  
--          SELECT Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance, OrderType,  
--                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate , AccountYear , TaxTotal , DutyTotal    
--   FROM tFacM WHERE tFacM.intSerialNo = @intSerialNo and Branch = @intBranch  

--      IF @@ERROR <>0  
--          GoTo EventHandler  

--      INSERT INTO tFacD2(Code , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate], intInventoryNo )   
--    SELECT @@identity , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate],intInventoryNo  
--                 From tFacD  
--                 WHERE intSerialNo = @intSerialNo  And Branch = @intBranch

--      IF @@ERROR <>0  
--          GoTo EventHandler  

-----------------------------------------Mojodi Control Online---------------------------------------------------------  
--Exec DeleteMojodiCalculate 11 , @intserialNo  ,  1 , @AccountYear , @intBranch  
--IF @@ERROR <>0  
-- GoTo EventHandler  
------------------------------------------Delete Old Details -----------------------------------------------------------  
--    DELETE FROM tFacD  
--    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
--    IF @@ERROR <>0  
--        GoTo EventHandler  
------------------------------------------Date From Server-----------------------------------------------------------------  
------------------------------------------Update Master-----------------------------------------------------------------  

--    Update tFacM  
--        SET 
--        DiscountTotal   = @DiscountTotal,  
--        CarryFeeTotal   = @CarryFeeTotal,  
--        SumPrice        = @SumPrice,  
--        ServiceTotal    = @ServiceTotal,  
--        PackingTotal    = @PackingTotal,  
--        [Time]          = dbo.SetTimeFormat(GETDATE()),  
--        NvcDescription  = @NvcDescription ,  
-- 		TempAddress     = @TempAddress
--    WHERE tFacM.intSerialNo = @intSerialNo  AND Branch =  @intBranch  

--    IF @@ERROR <>0  
--        GoTo EventHandler  


------------------------------------Fill Details Factor ----------------------------------------------------------------------  
-- exec InsertFactorDetail @DetailsString , @intserialNo , 0 , -2 , @intBranch        

--     IF @@ERROR <>0  
--        GoTo EventHandler  
----------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------  


------------------------------------Total SumPrice Calculate  --------------------------------------------------------------  
--DECLARE @DiscountD INT 
--Set @DiscountD = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * discount/100 ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
--Set @SumPrice = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * (1 - discount/100) ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

--     IF @@ERROR <>0      
--        GoTo EventHandler      
--DECLARE @TaxTotal FLOAT  
--SET @TaxTotal = 0
--DECLARE @ValueGoodsTax FLOAT
--SET @ValueGoodsTax = 0
--	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
--	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
--	              WHERE     intSerialNo = @intSerialNo
--	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
--	            )  
--SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
--IF @@ERROR <> 0 
--GOTO EventHandler

--DECLARE @DutyTotal INT 
--SET @DutyTotal = 0
--DECLARE @ValueGoodsDuty FLOAT
--SET @ValueGoodsDuty = 0
--	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
--	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
--	              WHERE     intSerialNo = @intSerialNo
--	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
--	            )  
--SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
--IF @@ERROR <> 0 
--GOTO EventHandler

------------------------------------ServiceRate Calculate  --------------------------------------------------------------  
-- If @ServiceTotal <> 0  
--       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)  

--     IF @@ERROR <>0  
--        GoTo EventHandler   
------------------------------------Round Sumprice  --------------------------------------------------------------  
--	SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
--	SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
--	Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal   

--    Declare @Remain INT  
--    SET @Remain = 0
--    Set @Remain = dbo.RoundSumPrice(@SumPrice )     
--    Set @SumPrice = @SumPrice - @Remain  
--    Set @DiscountTotal = @DiscountTotal + @Remain  
------------------------------------Net Price Update  --------------------------------------------------------------  

--    Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch   
-- IF @@ERROR <>0  
--         GoTo EventHandler  

--Update tFacm Set DiscountTotal = @DiscountTotal Where intSerialNo = @intserialNo  And Branch = @intBranch   
--Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch   
--Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch   
--Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
--Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

-------------------------------------------Fill Detail Cash ,....---------------------------------------------------  
-------------------------------------------History  --------------------------------------------------------------  

--Exec InsertHistory  @No, 11 , -1 , 2 ,@AccountYear  , @intBranch
-- IF @@ERROR <>0  
--        GoTo EventHandler  

-------------------------------------------Cash  --------------------------------------------------------------  

--------------------------------------------Mojodi Control Online-----------------------------------------------------  

--Exec InsertMojodiCalculate 11 , @intserialNo , @AccountYear , @intBranch  
--IF @@ERROR <>0  
-- GoTo EventHandler  
--------------------------------------------Update Balance After Recived----------------------------


--COMMIT TRANSACTION  

--Set @Result = @No  
--Return @Result  


--EventHandler:  
--    ROLLBACK TRAN  
--    SET @Result = -1   

--    RETURN @Result
--GO





if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WS_OrderPosition]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[WS_OrderPosition]
GO


CREATE  PROCEDURE [dbo].[WS_OrderPosition]  (  


	@intSerialNo  INT ,
	@Result int  OUT 
  )  


AS  


--0 Add or Edit
--1 Delivered By Payk
--2 Tasvieh 

DECLARE @Branch INT 
SET @Branch = dbo.Get_Current_Branch()
DECLARE @Payk int 
SET @Result = 0

SELECT @Payk = InCharge FROM dbo.tFacM WHERE intSerialNo = @intSerialNo AND Branch = @Branch
IF @Payk IS NOT NULL 
	BEGIN 
		SET @Result = 1
		if exists (select * from dbo.tHistory WHERE intSerialNo = @intSerialNo AND Branch = @Branch AND ActionCode = 4 ) 
			SET @Result = 1 

		if exists (select * from dbo.tHistory WHERE intSerialNo = @intSerialNo AND Branch = @Branch AND (ActionCode = 5  OR ActionCode =6 )) 
			SET @Result = 2 
	END 

GO
 
 
 
 
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].WS_All_PaykName') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[WS_All_PaykName]
GO


CREATE  PROCEDURE [dbo].[WS_All_PaykName]  
AS 
 SELECT * FROM dbo.tPer
	WHERE Job = 3
	AND ActDeAct = 1
	
	
GO


--تغییر برای چک نکردن ایستگاه بزرگتر از صفر
ALTER procedure dbo.Update_tStation_Inventory 

(@Branch Int , @InventoryNo int ,@AccountYear	Smallint ,@StationID nvarchar(400) )
as
Begin Tran

Insert into dbo.tStation_Inventory_Good ( branch ,InventoryNo, AccountYear ,StationID,  GoodCode , Active)

select t.Branch ,t.InventoryNo ,t.AccountYear ,t.StationID ,t.GoodCode ,1 as Active From
(Select  tInventory_Good.Branch , tInventory_Good.AccountYear , tInventory_Good.InventoryNo , t1.stationid  , tInventory_Good.Goodcode  , 1 as Active
	From tInventory_Good 
	 Inner join (Select cast(word as int) as StationID from dbo.SplitWithDelimiterNVarChar(@StationID , ','))t1  On inventoryno = @InventoryNo and Branch = @Branch And AccountYear = @AccountYear
   )t
        Where (t.GoodCode Not In (Select GoodCode  From tStation_Inventory_Good where inventoryno = t.inventoryno and Branch = t.Branch and StationId = t.StationId And AccountYear = t.AccountYear))

if @@Error <> 0 
	 GOTO ErrHandler

Commit Tran


Return

ErrHandler:
RollBack Tran 
Return



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER VIEW dbo.vw_NotPaidFactors      
AS      
SELECT     tfacm.intSerialNo, tfacm.[No], tfacm.Status, tfacm.Owner, tfacm.Customer, tfacm.DiscountTotal, tfacm.SumPrice, tfacm.CarryFeeTotal,      
  tfacm.Recursive, tfacm.FacPayment, tfacm.InCharge, tfacm.OrderType, tfacm.ServePlace, tfacm.StationID, tfacm.ServiceTotal, tfacm.PackingTotal,      
  tfacm.BascoleNo, tfacm.ShiftNo, tfacm.TableNo, tfacm.[Date], tfacm.[Time], tfacm.[User], tfacm.RegDate, tfacm.Branch, tfacm.Balance, tfacm.AccountYear, tfacm.NvcDescription, tfacm.RefFacM,       


  CASE dbo.tCust.[Name] + ' ' + dbo.tCust.Family WHEN ' ' THEN tCust.WorkName ELSE dbo.tCust.[Name] + ' ' +       
  dbo.tCust.Family END AS [Full Name],       
   dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tPer.Job, dbo.tCust.MembershipId AS Code, dbo.tCust.Address, dbo.tCust.Credit      
    ,dbo.tServePlace.[Description] AS ServePlaceName,dbo.tCust.distance,CASE LTRIM(RTRIM(dbo.tFacM.NvcDescription)) WHEN N'پيغام' THEN 1       
  WHEN N'' THEN 1 ELSE -1 END AS intWarn,ISNULL(LTRIM(RTRIM(dbo.tFacM.TempAddress)),'') AS TempAddress,      
  (DATEPART(hour, GETDATE()) * 60 + DATEPART(minute, GETDATE()))-      
  (CAST(SUBSTRING(dbo.tFacM.[Time], 1, 2) AS int) * 60 + CAST(SUBSTRING(dbo.tFacM.[Time], 4, 2) AS int))  AS RemainMinute      
  ,t.DateSend,t.TimeSend      
  ,      
  (DATEPART(hour, GETDATE()) * 60 + DATEPART(minute, GETDATE()))-      
  (CAST(SUBSTRING(t.TimeSend, 1, 2) AS int) * 60 + CAST(SUBSTRING(t.TimeSend, 4, 2) AS int))  AS RemainMinuteSend
    ,ISNULL(LTRIM(RTRIM(dbo.[tCust].Mobile)),'') AS Mobile        
, ISNULL(tfacm.GuestNo , '') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.No) AS TempNo	
    , tshift.Description AS ShiftDescription , dbo.tCust.Tafsili 
	, Tel1 + Tel2 + Tel3 + Tel4 +Mobile AS TelNumber
FROM         dbo.tFacM       
  INNER JOIN dbo.tServePlace ON dbo.tServePlace.intServePlace= dbo.tfacm.ServePlace      
  INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code
  LEFT OUTER JOIN dbo.tPer ON dbo.tFacM.InCharge = dbo.tPer.pPno and dbo.tper.ActDeact=1 --AND dbo.tFacM.Branch = dbo.tPer.Branch 
  LEFT OUTER JOIN dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code --AND (dbo.tFacM.Branch = dbo.tCust.Branch OR dbo.tCust.Branch IS NULL)      

 left outer  JOIN (SELECT MAX(RegDate) AS DateSend,MAX(RegTime) AS TimeSend,intserialno FROM thistory       
        WHERE ActionCode=4 GROUP BY intserialno) t      
 ON [tfacm].[intSerialNo] = t.[intSerialNo]        
 WHERE     (dbo.tFacM.Balance = 0 OR tfacM.StationID = -1) And Status =2 and Recursive=0  --for list peik in recived    




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER Proc Get_vw_NotPaidFactors_By_Job (@Job int , @AccountYear Smallint)  
as  
 Select vw_NotPaidFactors.intSerialNo,vw_NotPaidFactors.No,vw_NotPaidFactors.Status,  
        vw_NotPaidFactors.Owner,vw_NotPaidFactors.Customer,vw_NotPaidFactors.DiscountTotal,  
        vw_NotPaidFactors.SumPrice,vw_NotPaidFactors.CarryFeeTotal,vw_NotPaidFactors.Recursive,  
        vw_NotPaidFactors.FacPayment,vw_NotPaidFactors.InCharge,vw_NotPaidFactors.OrderType,  
        vw_NotPaidFactors.ServePlace,vw_NotPaidFactors.StationID,vw_NotPaidFactors.ServiceTotal,  
        vw_NotPaidFactors.PackingTotal,vw_NotPaidFactors.BascoleNo,vw_NotPaidFactors.ShiftNo,  
        vw_NotPaidFactors.TableNo,t.Date,t.TIME,vw_NotPaidFactors.RegDate,vw_NotPaidFactors.[USER],  
           vw_NotPaidFactors.Branch,vw_NotPaidFactors.Balance,vw_NotPaidFactors.ServePlaceName,  
        vw_NotPaidFactors.AccountYear,vw_NotPaidFactors.NvcDescription,vw_NotPaidFactors.RefFacM,  
        vw_NotPaidFactors.[Full NAME],vw_NotPaidFactors.nvcFirstName,vw_NotPaidFactors.nvcSurName,  
        vw_NotPaidFactors.Job,vw_NotPaidFactors.Code,vw_NotPaidFactors.Address,vw_NotPaidFactors.Credit,  
        vw_NotPaidFactors.distance,vw_NotPaidFactors.intWarn,vw_NotPaidFactors.RemainMinute,  
        vw_NotPaidFactors.TempAddress,vw_NotPaidFactors.mobile ,  
        vw_NotPaidFactors.GuestNo , vw_NotPaidFactors.TempNo , vw_NotPaidFactors.ShiftDescription 
        , vw_NotPaidFactors.Tafsili
  from vw_NotPaidFactors  
  INNER  JOIN (SELECT MAX(RegDate) AS Date,MAX(RegTime) AS Time,intserialno FROM thistory   
        WHERE ActionCode=4 GROUP BY intserialno) AS t   
		ON [vw_NotPaidFactors].[intSerialNo] = t.[intSerialNo]   
   Where  ((Balance = 0 And FacPayment = 0) OR StationID = -1 ) And AccountYear = @AccountYear
			AND vw_NotPaidFactors.InCharge > 0  --Job = @Job And 




GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER Proc Get_vw_NotPaidFactors_By_Job_InCharge (@Job int , @InCharge Int, @AccountYear Smallint)  
as  
 Select vw_NotPaidFactors.intSerialNo,vw_NotPaidFactors.No,vw_NotPaidFactors.Status,  
        vw_NotPaidFactors.Owner,vw_NotPaidFactors.Customer,vw_NotPaidFactors.DiscountTotal,  
        vw_NotPaidFactors.SumPrice,vw_NotPaidFactors.CarryFeeTotal,vw_NotPaidFactors.Recursive,  
        vw_NotPaidFactors.FacPayment,vw_NotPaidFactors.InCharge,vw_NotPaidFactors.OrderType,  
        vw_NotPaidFactors.ServePlace,vw_NotPaidFactors.StationID,vw_NotPaidFactors.ServiceTotal,  
        vw_NotPaidFactors.PackingTotal,vw_NotPaidFactors.BascoleNo,vw_NotPaidFactors.ShiftNo,  
        vw_NotPaidFactors.TableNo,t.Date,t.TIME,vw_NotPaidFactors.RegDate,vw_NotPaidFactors.[USER],  
           vw_NotPaidFactors.Branch,vw_NotPaidFactors.Balance,vw_NotPaidFactors.ServePlaceName,  
        vw_NotPaidFactors.AccountYear,vw_NotPaidFactors.NvcDescription,vw_NotPaidFactors.RefFacM,  
        vw_NotPaidFactors.[Full NAME],vw_NotPaidFactors.nvcFirstName,vw_NotPaidFactors.nvcSurName,  
        vw_NotPaidFactors.Job,vw_NotPaidFactors.Code,vw_NotPaidFactors.Address,vw_NotPaidFactors.Credit,  
        vw_NotPaidFactors.distance,vw_NotPaidFactors.intWarn,vw_NotPaidFactors.RemainMinute,  
		vw_NotPaidFactors.TempAddress,vw_NotPaidFactors.[Mobile] ,  
         vw_NotPaidFactors.GuestNo , vw_NotPaidFactors.TempNo , vw_NotPaidFactors.ShiftDescription
         , vw_NotPaidFactors.Tafsili
 from vw_NotPaidFactors  
  INNER  JOIN (SELECT MAX(RegDate) AS Date,MAX(RegTime) AS Time,intserialno FROM thistory   
        WHERE ActionCode=4 GROUP BY intserialno) AS t   
  ON [vw_NotPaidFactors].[intSerialNo] = t.[intSerialNo]   
   Where Job = @Job And InCharge = @InCharge And ((Balance = 0 And FacPayment = 0) OR StationId = -1) And AccountYear = @AccountYear
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER PROCEDURE dbo.GetCustomersInfo  

@AccountYear SmallInt   , @Branch INT 

AS  
 SELECT   dbo.tCust.MembershipId As Code, dbo.tFacM.intSerialNo ,dbo.tFacM.[No] ,  
  CASE  dbo.tCust.Family + dbo.tCust.[Name] WHEN ''  THEN tCust.WorkName  
                       ELSE dbo.tCust.[Name] + ' ' + dbo.tCust.Family  
  END AS [Full Name],  
  dbo.tFacM.SumPrice , dbo.tFacM.[Time] , dbo.tCust.Address ,dbo.tFacM.[Date] ,dbo.tfacm.ServePlace ,dbo.tServePlace.[Description] AS ServePlaceName  
  ,dbo.tCust.distance,CASE LTRIM(RTRIM(dbo.tFacM.NvcDescription)) WHEN N'پيغام' THEN 1   
  WHEN N'' THEN 1 ELSE -1 END AS intWarn,dbo.tFacM.NvcDescription,ISNULL(LTRIM(RTRIM(dbo.tFacM.TempAddress)),'') AS TempAddress  
  ,ISNULL(LTRIM(RTRIM(dbo.[tCust].Mobile)),'') AS Mobile
, ISNULL(tfacm.GuestNo , '') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.No) AS TempNo	
    , tshift.Description AS ShiftDescription , dbo.tCust.Tafsili , tfacM.StationID
FROM   dbo.tFacM  
	Left Outer JOIN tCust ON dbo.tCust.Code = dbo.tFacM.Customer   
	INNER JOIN dbo.tServePlace ON dbo.tServePlace.intServePlace= dbo.tfacm.ServePlace  
    INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code

 WHERE   ( dbo.tFacM.Incharge IS NULL OR dbo.tFacM.Incharge = '' ) AND dbo.tFacM.Status = 2  
		AND (dbo.tFacM.facPayment = 0  OR tfacM.StationID = -1)  
		and dbo.tFacM.TableNo is null   
		--And dbo.tfacm.Customer > 0  
		And dbo.tfacm.Recursive <> 1  
		And (dbo.tfacm.ServePlace = 2 OR dbo.tfacm.ServePlace = 4)  
		And AccountYear = @AccountYear   
		AND dbo.tFacM.Branch =  @Branch  
 ORDER BY dbo.tFacM.[No] , dbo.tFacM.[Date] ,dbo.tFacM.[Time]  




GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER VIEW dbo.VwTotal_NotDelivers
AS
SELECT     dbo.tCust.MembershipId As Code,dbo.tFacM.intSerialNo, dbo.tFacM.[No], 
                      CASE dbo.tCust.Family + dbo.tCust.[Name] WHEN '' THEN tCust.WorkName ELSE dbo.tCust.[Name] + ' ' + dbo.tCust.Family END AS [Full Name], 
                      dbo.tFacM.SumPrice, dbo.tFacM.[Time], dbo.tCust.Address, dbo.tFacM.[Date], dbo.tfacm.ServePlace, 
                      dbo.tServePlace.[Description] AS ServePlaceName,dbo.tFacM.AccountYear,
		 (DATEPART(hour, GETDATE()) * 60 + DATEPART(minute, GETDATE()))-
		(CAST(SUBSTRING(dbo.tFacM.[Time], 1, 2) AS int) * 60 + CAST(SUBSTRING(dbo.tFacM.[Time], 4, 2) AS int))  AS RemainMinute
		,dbo.tCust.distance,CASE LTRIM(RTRIM(dbo.tFacM.NvcDescription)) WHEN N'پيغام' THEN 1 
		WHEN N'' THEN 1 ELSE -1 END AS intWarn,LTRIM(RTRIM(dbo.tFacM.TempAddress)) AS TempAddress
		, Tel1 + Tel2 + Tel3 + Tel4 +Mobile AS TelNumber , tfacM.StationID
FROM         dbo.tFacM LEFT OUTER JOIN
                      tCust ON dbo.tCust.Branch = dbo.tFacM.Branch AND dbo.tCust.Code = dbo.tFacM.Customer INNER JOIN
                      dbo.tServePlace ON dbo.tServePlace.intServePlace = dbo.tfacm.ServePlace
WHERE     (dbo.tFacM.Incharge IS NULL OR
                      dbo.tFacM.Incharge = '') AND dbo.tFacM.Status = 2 
                      AND (dbo.tFacM.facPayment = 0 OR tfacM.StationId = -1)
                      AND dbo.tFacM.TableNo IS NULL AND dbo.tfacm.Recursive <> 1 
                      AND (dbo.tfacm.ServePlace = 2 OR dbo.tfacm.ServePlace = 4)
                      AND dbo.tFacM.Branch = dbo.Get_Current_Branch()




GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE dbo.GetTotal_Delivers 
	@AccountYear SMALLINT,
	@Job INT
AS
Select Code,intSerialNo,[no],[Full Name],SumPrice,[Time],Address,[Date],ServePlace,ServePlaceName,N'سفارش' as DeliverStatus,
	CAST(RemainMinute/60  AS VARCHAR(4))+':'+CAST(RemainMinute%60  AS VARCHAR(4)) AS RemainTime
	,N'ارسال نشده' as RemainTimesend,N'ارسال نشده' as Timesend
	,distance,intWarn,isnull(TempAddress,'') as TempAddress , '' AS InchargeName , '' AS Incharge
	, ISNULL(VwTotal_NotDelivers.TelNumber , N'') AS TelNumber , StationID
	from VwTotal_NotDelivers WHERE AccountYear = @AccountYear
UNION
Select  Code,intSerialNo,[no],[Full Name],SumPrice,[Time],Address,[Date],ServePlace,ServePlaceName,N'ارسال شده' as DeliverStatus,
	CAST(RemainMinute/60  AS VARCHAR(4))+':'+CAST(RemainMinute%60  AS VARCHAR(4)) AS RemainTime
	,CAST(RemainMinutesend/60  AS VARCHAR(4))+':'+CAST(RemainMinutesend%60  AS VARCHAR(4)) AS RemainTimesend,Timesend
	,distance,intWarn,isnull(TempAddress,'')as TempAddress , ISNULL(nvcFirstName , '') + ' ' + ISNULL(nvcSurName, '') AS InchargeName , InCharge
	, ISNULL(vw_NotPaidFactors.TelNumber , N'') AS TelNumber , StationID
	from vw_NotPaidFactors Where Job = @Job And (Balance = 0 OR StationID = -1)
ORDER BY [No] , [Date] ,[Time]



GO


ALTER FUNCTION [dbo].[fn_Get_Mojodi_IntermediateGoods]
    (
      @GoodCode INT ,
      @InventoryNo INT ,
      @Branch INT ,
      @AccountYear INT
    )
RETURNS INT
AS
    BEGIN

        DECLARE @GoodFirstCode INT
        DECLARE @Mojodi AS INT 

        SET @GoodFirstCode = ( SELECT TOP 1
                                        [GoodFirstCode]
                               FROM     [dbo].[tUsePercent]
                               WHERE    [GoodCode] = @GoodCode
                                        AND [GoodFirstCode] IN ( SELECT
                                                              [Code]
                                                              FROM
                                                              [dbo].[tGood]
                                                              WHERE
                                                              [GoodType] = 4 )
                             )
        IF @GoodFirstCode IS NULL
            BEGIN 
                SELECT  @Mojodi = [Mojodi] / 2
                FROM    [dbo].[tInventory_Good]
                WHERE   [GoodCode] = @GoodCode
                        AND [InventoryNo] = @InventoryNo
                        AND [Branch] = @Branch
                        AND [AccountYear] = @AccountYear
                IF @Mojodi <= 0
                    AND @GoodCode IN ( SELECT   [Code]
                                       FROM     [dbo].[tGood]
                                       WHERE    ([GoodType] = 2 OR [GoodType] = 3 ))
                    SET @Mojodi = 1 
             END 
        ELSE
            SELECT  @Mojodi = [Mojodi] / 2
            FROM    [dbo].[tInventory_Good]
            WHERE   [GoodCode] = @GoodFirstCode
                    AND [InventoryNo] = @InventoryNo
                    AND [Branch] = @Branch
                    AND [AccountYear] = @AccountYear

        RETURN(@Mojodi)

    END

GO


--SELECT dbo.fn_Get_Mojodi_IntermediateGoods(11010001, 1, 1,1395)

--GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetMojodiWithGoodCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetMojodiWithGoodCode]
GO


CREATE  PROCEDURE [dbo].[GetMojodiWithGoodCode]
AS

    SELECT  Code AS GoodCode ,
            dbo.[fn_Get_Mojodi_IntermediateGoods](Code, 1, 1,
                                                  dbo.Get_AccountYear()) AS Mojodi
    FROM    tgood
    WHERE   GoodType IN ( 2, 3 ) AND BitWebShow = 1 --AND code = 11010001
    
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Web_SetSaveInArya]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_Web_SetSaveInArya]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

create PROCEDURE [dbo].[sp_Web_SetSaveInArya] (@intSerialNo bigint , @Branch int)
AS

	update tbl_Web_tFacM set SaveInArya = 1 where intSerialNo  = @intSerialNo and Branch = @Branch
    
GO

 
    
 