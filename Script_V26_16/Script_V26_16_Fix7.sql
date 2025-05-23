
--Script_V26_16_Fix7
--در نسخه های پیشرفته و طلایی
--کنترل زمان اصلاح فیش با تایم کاربر و مدیر
--[Get_UserEditTime] = 30  زمان کاربربه دقیقه
--[Get_ManagerEditTime] = 50000   زمان مدیر سیستم به دقیقه
--بعد از تایم مدیر دیگر فیش قابل اصلاح نیست
--به روزرسانی فی و مبلغ هنگام تغییر نرخ در حالت مشترکین و دستی
--اصلاح ورود و ثبت اطلاعات پوزهای بانکی
--ثبت تغییرات کالاها مستقیمادر فاکتور فروش
--اصلاح فرم اختصاص به پیک
-- 93/-7/18


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
          7
        )
GO

ALTER TABLE tblPub_Pos
ALTER COLUMN nvcAccountNo NVARCHAR(50)
GO

ALTER TABLE tblPub_Pos
ALTER COLUMN  nvcBankName NVARCHAR(50)
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Insert_tblPub_Pos] 
(
	@PosId INT ,
	@NvcPosNo nvarchar(20) , 
	@nvcBankName nvarchar(50) , 
	@nvcAccountNo nvarchar(50) , 
	@AccountId INT ,
	@intStatus int out)
AS


Begin Tran


Insert Into dbo.tblPub_Pos
        ( PosId ,
          NvcPosNo ,
          nvcBankName ,
          nvcAccountNo ,
          AccountId
        )
VALUES  ( @PosId , -- PosId - int
          @NvcPosNo , -- NvcPosNo - nvarchar(20)
          @nvcBankName , -- BankName - nvarchar(20)
          @nvcAccountNo , -- nvcAccountNo - nvarchar(20)
          @AccountId
        )

if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus=@PosId
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Update_tblPub_Pos] (
	@PosId INT ,
	@NvcPosNo nvarchar(20) , 
	@nvcBankName nvarchar(50) , 
	@nvcAccountNo nvarchar(50) ,
	@AccountId INT , 
	@NewPosId INT ,
	@intStatus int out)

AS

Begin Tran

UPDATE dbo.tblPub_Pos SET
	PosId = @NewPosId ,
	NvcPosNo = @NvcPosNo  , 
	nvcBankName = @nvcBankName , 
	nvcAccountNo = @nvcAccountNo ,
	AccountId = @AccountId 

   WHERE PosId = @PosId

if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus = 1
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return



GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_UserEditTime]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Get_UserEditTime]
GO

CREATE FUNCTION [dbo].[Get_UserEditTime]()
RETURNS int 
AS  
BEGIN 
Return 1500
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_ManagerEditTime]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Get_ManagerEditTime]
GO

CREATE FUNCTION [dbo].[Get_ManagerEditTime]()
RETURNS int 
AS  
BEGIN 
Return 1000000
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_CurrentEditTime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_CurrentEditTime
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE Get_CurrentEditTime
@intserialNo INT ,
@Baranch INT 
 
AS

DECLARE @MinuteUseDiff INT
SELECT @MinuteUseDiff = 
( CAST(SUBSTRING(dbo.shamsi(GETDATE()), 4, 2) AS INT) - 1 * 30
                + CAST(SUBSTRING(dbo.shamsi(GETDATE()), 7, 2) AS INT) ) * 1440
+ ( DATEPART(HOUR, GETDATE()) * 60 + DATEPART(minute, GETDATE()) ) 
 - ( CAST(SUBSTRING(T.nvcDate, 4, 2) AS INT) - 1 * 30
                + CAST(SUBSTRING(T.nvcDate, 7, 2) AS INT) )  * 1440
           - ( CAST(SUBSTRING(T.nvcTime, 1, 2) AS INT) * 60
                + CAST(SUBSTRING(T.nvcTime, 4, 2) AS INT) ) 
from 
(
SELECT TOP 1 ISNULL(dbo.tRepFacEditM.Time , dbo.tFacM.Time) AS nvcTime , 
	         ISNULL(dbo.tRepFacEditM.RegDate , dbo.tFacM.RegDate) AS nvcDate 
 FROM dbo.tFacM LEFT OUTER JOIN dbo.tRepFacEditM ON dbo.tRepFacEditM.Branch = dbo.tFacM.Branch AND dbo.tRepFacEditM.intSerialNo = dbo.tFacM.intSerialNo
WHERE dbo.tFacM.intSerialNo = @intserialNo AND dbo.tFacM.Branch = @Baranch	
	And Code = (Select MIN(Code) from tRepFacEditM where intSerialNo = @intserialNo AND Branch = @Baranch)
) T

--SET @MinuteUseDiff = 25
SELECT @MinuteUseDiff AS  MinuteUseDiff,
 CASE WHEN @MinuteUseDiff < 0 THEN 0 WHEN @MinuteUseDiff > dbo.[Get_UserEditTime]() THEN 0 ELSE 1 END AS UserDiffTme ,
 CASE WHEN @MinuteUseDiff < 0 THEN 0 WHEN @MinuteUseDiff > dbo.[Get_ManagerEditTime]() THEN 0 ELSE 1 END AS ManagerDiffTme


GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--EXEC Get_CurrentEditTime 10000068 , 1



SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Insert_Differences](
	@intLanguage int ,
	@Defference NvarChar(200) ,
	@NegativeDefference NvarChar(200),
	@CostDifference int=0 ,
	@LastCode INT OUT 
 ) 
AS
set @LastCode = (Select ISNULL( Max(abs(Code)),0) +1 from tDifferences)
Insert into tDifferences (Code , [Difference] , LatinDifference,CostDifference) Values
(
@LastCode ,
case @intLanguage
	when 0 then @Defference
	else ''
	end , 
case @intLanguage
	when 0 then ''
	else @Defference
	end ,
ABS(@CostDifference)
)

Insert into tDifferences (Code , [Difference] , LatinDifference,CostDifference) Values
(
@LastCode ,
case @intLanguage
	when 0 then @NegativeDefference
	else ''
	end , 
case @intLanguage
	when 0 then ''
	else @NegativeDefference
	end ,
0
)

RETURN @LastCode

GO







