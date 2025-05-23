
--Script_V26_16_Fix8_NewPos
--اضافه شدن ایستگاه به  جدول پوز بانکی
--فقط استفاده از آسان پرداخت در کد های نرم افزاری
--استفاده از PosInterface.dll جدید برای 
-- 93/10/29


delete FROM dbo.tblPub_Pos
GO

CREATE TABLE [dbo].[tblPub_PosType]
(
[PosId] [int] NOT NULL,
PosName NVARCHAR(50) NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblPub_PosType] ADD CONSTRAINT [PK_tblPub_PosType] PRIMARY KEY CLUSTERED  ([PosId] ) ON [PRIMARY]
GO

INSERT INTO dbo.tblPub_PosType
        ( PosId, PosName )
VALUES  ( 1, -- PosId - int
          N'پوز آسان پرداخت'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosId, PosName )
VALUES  ( 2, -- PosId - int
          N'پوز بانکي پاسارگاد'  -- PosName - nvarchar(50)
          )
GO


IF COL_LENGTH('[tblPub_Pos]','AutoId') IS NULL
	ALTER TABLE [tblPub_Pos]
	ADD AutoId INT IDENTITY(1,1)
GO


IF COL_LENGTH('[tblPub_Pos]','[NvcPosNo]') IS NOT NULL
	UPDATE [tblPub_Pos] SET [NvcPosNo] = AutoId
GO


IF COL_LENGTH('[tblPub_Pos]','NvcPosNo') IS NOT NULL
	EXEC sp_rename '[tblPub_Pos].[NvcPosNo]', 'StationId', 'COLUMN';
GO

IF COL_LENGTH('[tblPub_Pos]','StationId') IS NOT NULL
	ALTER TABLE tblPub_Pos ALTER COLUMN StationId INT NOT NULL 
GO

ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT [FK_tFacCard_tblPub_Pos]
GO

ALTER TABLE [dbo].[tblAcc_ReceivedSummary] DROP CONSTRAINT [FK_tblAcc_ReceivedSummary_tPos]
GO

ALTER TABLE [dbo].[tblPub_Pos] DROP CONSTRAINT [PK_tblPub_Pos]
GO

ALTER TABLE [dbo].[tblPub_Pos] ADD CONSTRAINT [PK_tblPub_Pos] PRIMARY KEY CLUSTERED  ([PosId] , StationId ) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tFacCard] ADD CONSTRAINT [FK_tFacCard_tblPub_PosType] FOREIGN KEY ([PosId]) REFERENCES [dbo].[tblPub_PosType] ([PosId]) ON DELETE CASCADE
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Insert_tblPub_Pos] 
(
	@PosId INT ,
	@StationId INT  , 
	@nvcBankName nvarchar(50) , 
	@nvcAccountNo nvarchar(50) , 
	@AccountId INT ,
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@intStatus int out)
AS


Begin Tran


Insert Into dbo.tblPub_Pos
        ( PosId ,
          StationId ,
          nvcBankName ,
          nvcAccountNo ,
          AccountId ,
          PosType ,
          ComunicationType ,
          PosAddress
        )
VALUES  ( @PosId , -- PosId - int
          @StationId , -- NvcPosNo - nvarchar(20)
          @nvcBankName , -- BankName - nvarchar(20)
          @nvcAccountNo , -- nvcAccountNo - nvarchar(20)
          @AccountId ,
          @PosId ,
          @CommunicationType ,
          @PosAddress
        )

if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus=@@IDENTITY
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   PROCEDURE [dbo].[Update_tblPub_Pos] (
	@AutoId INT ,
	@StationId INT  , 
	@nvcBankName nvarchar(50) , 
	@nvcAccountNo nvarchar(50) ,
	@AccountId INT , 
	@NewPosId INT ,
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@intStatus int out)

AS

Begin Tran

UPDATE dbo.tblPub_Pos SET
	PosId = @NewPosId ,
	StationId = @StationId  , 
	nvcBankName = @nvcBankName , 
	nvcAccountNo = @nvcAccountNo ,
	AccountId = @AccountId ,
	PosType = @NewPosId ,
	ComunicationType = @CommunicationType ,
	PosAddress = @PosAddress

   WHERE AutoId = @AutoId

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

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_tblPub_Pos_ById] 
@AutoId INT 
AS
select * from [tblPub_Pos] WHERE  AutoId = @AutoId

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Delete_tblPub_Pos](
	@AutoId INT )
AS
	DELETE FROM dbo.tblPub_Pos WHERE AutoId = @AutoId

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_All_tblPub_Pos] AS
select * from [tblPub_Pos] ORDER BY AutoId

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_PosType]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_All_PosType
GO

CREATE PROC Get_All_PosType
AS 
select * from [dbo].[tblPub_PosType] ORDER BY PosId

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tblPub_Pos_ByStationId]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tblPub_Pos_ByStationId]
GO

CREATE   PROCEDURE [dbo].[Get_tblPub_Pos_ByStationId] 
@StationId INT 
AS
select * from [tblPub_Pos] WHERE  StationId = @StationId

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Get_tfacm_Card_Detail]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @User1 INT,
      @User2 INT,
      @Station1 INT,
      @Station2 INT,   
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50)
      
    )
AS 
    DECLARE @tmp1 INT  
    DECLARE @tmp2 NVARCHAR(50)  
    DECLARE @Time3 NVARCHAR(50)  
    DECLARE @Time4 NVARCHAR(50)  
    SET @Time3 = @Time1  
    SET @Time4 = @time2  
  
    IF @User2 < @User1 
        BEGIN   
            SET @tmp1 = @User2  
            SET @User2 = @User1  
            SET @User1 = @tmp1   
        END   
  
    IF @Time2 < @Time1 
        BEGIN  
  /*SET @tmp2 = @Time2  
  SET @Time2 = @Time1  
  SET @Time1 = @tmp2*/  
            SET @Time3 = '00:00'  
            SET @Time4 = '24:00'  
        END  
    DECLARE @TimeTitle NVARCHAR(10)  
    SET @TimeTitle = N' ساعت : '  

    SELECT  
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate,
			tfacm.* , dbo.tFacCard.* ,dbo.tblPub_Pos.* 
        FROM    tfacm  
			INNER JOIN dbo.tFacCard ON dbo.tFacM.Branch = dbo.tFacCard.Branch AND dbo.tFacM.intSerialNo = dbo.tFacCard.intSerialNo
			INNER JOIN dbo.tblPub_Pos ON dbo.tFacCard.PosId = dbo.tblPub_Pos.PosId AND dbo.tblPub_Pos.StationId = dbo.tFacM.StationID
    WHERE   [date ] >= @Date1
            AND [date] <= @Date2
            AND [User] >= @User1
            AND [User] < = @User2
            AND ( ( [Time] >= @Time1
                    AND [Time] <= @Time4
                  )
                  OR ( [Time] <= @Time2
                       AND [Time] >= @Time3
                     )
                )
            AND tfacM.StationID >= @Station1
            AND tfacM.StationID <= @Station2
    ORDER BY [Date]  
 
GO



