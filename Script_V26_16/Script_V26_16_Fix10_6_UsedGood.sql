
--Script_V26_16_Fix10_6_UsedGood
--چاپ لیبل خاص برای کالاها
--ساخت جدولی که عناصر تشکیل دهنده کالا می باشد
--tGood_Used  نام جدول جدید
-- ممکن است یک کالا از چند کالای دیگر تشکیل شده باشد
-- و در اینصورت این اقلام مصرفی در پرینت لیبل شرکت می کنند
--برای جلوگیری از اشتباه از دستکاری جدول ضریب مصرف پرهیز می گردد
--درسایر تنظیمات استفاده از جدول مصرف برای لیبل تیک بخورد
-- اگر کالایی در جدول مصرف تعیین نشود همان روال قبلی پابرجاست و فقط آن کالا لیبلش چاپ می شود
--93/10/26


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tGood_tGood_Used]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tGood_Used] DROP CONSTRAINT FK_tGood_tGood_Used
GO

/****** Object:  Table [dbo].[tGood]    Script Date: 01/14/2015 10:27:55 ******/

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tGood_Used]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tGood_Used]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE TABLE [dbo].[tGood_Used] (
	[GoodCode] [int] NOT NULL ,
	[GoodFirstCode] [int] NOT NULL ,
	[Auto_Id] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tGood_Used] ON [dbo].[tGood_Used]([GoodCode]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tGood_Used] ADD 
	CONSTRAINT [FK_tGood_tGood_Used] FOREIGN KEY 
	(
		[GoodCode]
	) REFERENCES [dbo].[tGood] (
		[Code]
	) ON UPDATE CASCADE 
GO


--INSERT INTO dbo.tGood_Used
--        ( GoodCode , GoodFirstCode )
--SELECT Code , Code FROM dbo.tGood
--GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Good_Used]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Good_Used]
GO

CREATE PROCEDURE Get_Good_Used
@GoodCode INT 

AS

SELECT RTRIM(lTrim(tGood.Name)) AS nvcName, (SELECT COUNT(*) FROM dbo.tGood_Used WHERE GoodCode = @GoodCode)  AS UsedCount  FROM dbo.tGood_Used INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tGood_Used.GoodFirstCode
WHERE GoodCode = @GoodCode

GO


ALTER  PROCEDURE [dbo].[Get_FacMD_Good] (@No Bigint , @Status int , @intLanguage int , @AccountYear Smallint  , @Branch INT ) 

AS
DECLARE @intSerialNo INT 
SELECT @intSerialNo = intSerialNo FROM tfacM where No = @No  And  Status = @Status And  AccountYear =  @AccountYear AND Branch = @Branch


Select Sum(vw_FacMD_Good.Amount)As Amount  , vw_FacMD_Good.GoodCode    , Max(vw_FacMD_Good.ServePlace) As Serveplace , Max(vw_FacMD_Good.DifferencesCodes) As DifferencesCodes  ,
	Max(vw_FacMD_Good.DifferencesDescription) As DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]  ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	MAX(vw_FacMD_Good.Rate) as rate ,vw_FacMD_Good.ChairName  , Max(vw_FacMD_Good.FeeUnit)As FeeUnit ,
	Max(vw_FacMD_Good.intinventoryNo)As intinventoryNo ,Max(vw_FacMD_Good.DestInventoryNo)As DestInventoryNo,Max(vw_FacMD_Good.[ExpireDate])As [ExpireDate] , Max( IsNull(vw_FacMD_Good.[NvcDescription], ''))As [NvcDescription]
        , case @intLanguage when 0 then Name 
			    when 1 then LatinName end as nvcName , vw_FacMD_Good.intRow
	,Max(vw_FacMD_Good.NumberOfUnit) As NumberOfUnit , vw_FacMD_Good.maintype,Max( IsNull(vw_FacMD_Good.[TempAddress], ''))As [TempAddress]
	, Max(vw_FacMD_Good.[Description]) AS UnitDescription , ISNULL(T.Mojodi , 0 ) AS Mojodi
	 , TaxBuy , TaxSale , DutyBuy , DutySale , ISNULL(DestinationId , 0) AS DestinationId
	 , (SELECT SUM(Amount) FROM dbo.tFacD WHERE intSerialNo = @intSerialNo AND Branch = @Branch) AS SumAmount

 	from vw_FacMD_Good 
	LEFT OUTER JOIN
	(SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
	AND t.Branch = @Branch AND t.GoodCode = vw_FacMD_Good.GoodCode AND t.InventoryNo = vw_FacMD_Good.intinventoryNo
 	 
	where No = @No  And  Status = @Status And  vw_FacMD_Good.AccountYear =  @AccountYear AND vw_FacMD_Good.Branch = @Branch
	 Group By vw_FacMD_Good.GoodCode     , vw_FacMD_Good.DifferencesCodes ,vw_FacMD_Good.DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]    ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	vw_FacMD_Good.ChairName   , vw_FacMD_Good.ServePlace , vw_FacMD_Good.FeeUnit  , vw_FacMD_Good.intRow , vw_FacMD_Good.MainType
	,T.Mojodi  , TaxBuy , TaxSale , DutyBuy , DutySale , DestinationId
Order By  vw_FacMD_Good.intRow




GO
