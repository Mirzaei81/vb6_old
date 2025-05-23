

--فقط در ورژن های الماس 
--Script_V26_16_Fix8_GoodAnalyze
--آنالیز کالای ناخالص 
--در فاکتور خرید با انتخاب آنالیز کالا می توان کالای ناخالص را با مقدار آن انتخاب کرز
-- سپس با مشخص کردن کالاهای آماده شده از کالای اول مقادیر آنالیز شده در دیتابیس ثبت شده
-- و برای کالای ناخالص از انبار مشخص شده حواله صادر شده 
-- و برای کالاهای آماده شده در انبار مشخص شده رسید صادر می گردد
--93/08/11


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTotal_Good_Analyze_tGood]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTotal_Good_Analyze] DROP CONSTRAINT FK_tblTotal_Good_Analyze_tGood
GO


IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_tblTotal_Good_Analyze_Pert]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[tblTotal_Good_Analyze] DROP CONSTRAINT [DF_tblTotal_Good_Analyze_Pert]
END

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTotal_Good_Analyze]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblTotal_Good_Analyze]
GO

CREATE TABLE [dbo].[tblTotal_Good_Analyze] (
	[GoodCode] [int] NOT NULL ,
	[GoodFirstCode] [int] NOT NULL ,
	[fltUsedValue] [float] NOT NULL ,
	[Pert] [int] NULL ,
	[nvcDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblTotal_Good_Analyze] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblTotal_Good_Analyze] PRIMARY KEY  CLUSTERED 
	(
		[GoodCode],
		[GoodFirstCode]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTotal_Good_Analyze] ADD 
	CONSTRAINT [DF_tblTotal_Good_Analyze_Pert] DEFAULT (0) FOR [Pert]
GO

ALTER TABLE [dbo].[tblTotal_Good_Analyze] ADD 
	CONSTRAINT [FK_tblTotal_Good_Analyze_tGood] FOREIGN KEY 
	(
		[GoodCode]
	) REFERENCES [dbo].[tGood] (
		[Code]
	) ON UPDATE CASCADE 
GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_undefined_Good') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_undefined_Good
GO



CREATE   PROCEDURE dbo.Get_undefined_Good (@GoodCode int)

AS


SELECT     tGood.*      
 FROM         dbo.tGood Where tGood.Code not in    
                          (SELECT     GoodCode   
                             FROM         dbo.tblTotal_Good_Analyze    
                             WHERE     GoodCode = @GoodCode)   

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_Defined_Good') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_Defined_Good
GO



CREATE   PROCEDURE dbo.Get_Defined_Good (@GoodCode int)


AS

SELECT     *  FROM         dbo.tblTotal_Good_Analyze  
		INNER JOIN dbo.tGood ON tblTotal_Good_Analyze.GoodFirstCode = dbo.tGood.Code  
                WHERE     tblTotal_Good_Analyze.GoodCode = @GoodCode


GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Delete_tblTotal_Good_Analyze') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Delete_tblTotal_Good_Analyze
GO

CREATE   PROCEDURE dbo.Delete_tblTotal_Good_Analyze 
(
@GoodCode INT 

)


AS

DELETE FROM tblTotal_Good_Analyze WHERE GoodCode = @GoodCode


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Insert_tblTotal_Good_Analyze') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Insert_tblTotal_Good_Analyze
GO


CREATE   PROCEDURE dbo.Insert_tblTotal_Good_Analyze 
(
@GoodCode INT ,
@nvcGoodFirstCode nvarchar(4000)  ,
@nvcfltUsedValue nvarchar(4000)  

)


AS

EXEC Delete_tblTotal_Good_Analyze @GoodCode

INSERT INTO dbo.tblTotal_Good_Analyze
        ( GoodCode ,
          GoodFirstCode ,
          fltUsedValue ,
          Pert ,
          nvcDate
        )
SELECT  @GoodCode ,
	cast(s.word as int) as GoodFirstCode ,
	cast(u.word as float) as fltUsedValue ,
	0 , -- Pert - int
	dbo.shamsi(GETDATE())  -- nvcDate - nvarchar(10)
from 	dbo.SplitWithDelimiterNVarChar (@nvcGoodFirstCode , ',')s 
	inner join 
	dbo.SplitWithDelimiterNVarChar(@nvcfltUsedValue, ',')u 
	on  s.row = u.row


GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_Good_By_GoodType') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_Good_By_GoodType
GO


CREATE   PROCEDURE dbo.Get_Good_By_GoodType (@GoodType int)
AS

SELECT  *  FROM dbo.tGood
Where GoodType = @GoodType
Order By [Name]



GO



