
--اضافه شدن نام همه بانک ها به دیتابیس
--تغییر در پوز بانکی که بتواند همه بانک ها را پوشش دهد
--940116

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPub_Pos_tblPub_PosPort]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblPub_Pos] DROP CONSTRAINT FK_tblPub_Pos_tblPub_PosPort
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_PosPort]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_PosPort]
GO

CREATE TABLE [dbo].[tblPub_PosPort]
(
[PortId] [int] NOT NULL,
PortName NVARCHAR(50) NOT NULL 
) ON [PRIMARY]
GO


ALTER TABLE [dbo].[tblPub_PosPort] ADD CONSTRAINT [PK_tblPub_PosPort] PRIMARY KEY CLUSTERED  ([PortId] ) ON [PRIMARY]
GO


INSERT INTO [dbo].[tblPub_PosPort]
        ( [PortId], [PortName] )
VALUES  ( 1, -- PortId - int
          N'Usb'  -- PortName - nvarchar(50)
          )

GO

INSERT INTO [dbo].[tblPub_PosPort]
        ( [PortId], [PortName] )
VALUES  ( 2, -- PortId - int
          N'Serial'  -- PortName - nvarchar(50)
          )

GO

INSERT INTO [dbo].[tblPub_PosPort]
        ( [PortId], [PortName] )
VALUES  ( 3, -- PortId - int
          N'Lan'  -- PortName - nvarchar(50)
          )

GO

----###

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_All_PosPort') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_All_PosPort
GO

CREATE PROC Get_All_PosPort
AS 
select * from [dbo].[tblPub_PosPort] ORDER BY PortId

GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacCard_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT FK_tFacCard_tblPub_PosType
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPub_Pos_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblPub_Pos] DROP CONSTRAINT FK_tblPub_Pos_tblPub_PosType
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsPrimary') = 1)
ALTER TABLE [dbo].[tblPub_PosType] DROP CONSTRAINT [PK_tblPub_PosType]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_PosType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_PosType]
GO

CREATE TABLE [dbo].[tblPub_PosType]
(
[PosTypeId] [int] NOT NULL,
PosName NVARCHAR(50) NOT NULL 
) ON [PRIMARY]
GO


ALTER TABLE [dbo].[tblPub_PosType] ADD CONSTRAINT [PK_tblPub_PosType] PRIMARY KEY CLUSTERED  ([PosTypeId] ) ON [PRIMARY]
GO


INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 1, -- PosId - int
          N'پوز آسان پرداخت'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 2, -- PosId - int
          N'پوز بانکي پاسارگاد'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 3, -- PosId - int
          N'پوز بانکي ایران کیش'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 4, -- PosId - int
          N'پوز بانکي ملت'  -- PosName - nvarchar(50)
          )
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblAcc_ReceivedSummary_tPos]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblAcc_ReceivedSummary] DROP CONSTRAINT FK_tblAcc_ReceivedSummary_tPos
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacCard_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT [FK_tFacCard_tblPub_PosType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacCard_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT FK_tFacCard_tblPub_Pos
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_Pos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_Pos]
GO

CREATE TABLE [dbo].[tblPub_Pos] (
	[PosId] [int] IDENTITY (1, 1) NOT NULL ,
	[StationId] [int] NOT NULL ,
	[intBank] [int] NULL ,
	[nvcAccountNo] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AccountId] [int] NULL ,
	[PosType] [int] NULL ,
	[ComunicationType] [int] NULL ,
	[PosAddress] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PosPort] [int] NULL ,
	[nvcPosSerialNo] nvarchar(20)COLLATE SQL_Latin1_General_CP1_CI_AS NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblPub_Pos] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblPub_Pos] PRIMARY KEY  CLUSTERED 
	(
		[PosId]
	)  ON [PRIMARY] 
GO

--####

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tblAcc_Bank]') and OBJECTPROPERTY(id, N'IsPrimaryKey') = 1)
ALTER TABLE [dbo].[tblAcc_Bank] DROP CONSTRAINT [PK_tblAcc_Bank]
GO

ALTER TABLE [dbo].[tblAcc_Bank]
ALTER COLUMN [tintBank] INT NOT NULL 
GO

ALTER TABLE [dbo].[tblAcc_Bank] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblAcc_Bank] PRIMARY KEY  CLUSTERED 
	(
		[tintBank]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPub_Pos] ADD 
	CONSTRAINT [FK_tblPub_Pos_tblAcc_Bank] FOREIGN KEY 
	(
		[intBank]
	) REFERENCES [dbo].[tblAcc_Bank] (
		[tintBank]
	) ON DELETE CASCADE  ON UPDATE CASCADE ,
	CONSTRAINT [FK_tblPub_Pos_tblPub_PosType] FOREIGN KEY 
	(
		[PosType]
	) REFERENCES [dbo].[tblPub_PosType] (
		[PosTypeId]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO


ALTER TABLE [dbo].[tblPub_Pos] ADD 
	CONSTRAINT [FK_tblPub_Pos_tblPub_PosPort] FOREIGN KEY 
	(
		[ComunicationType]
	) REFERENCES [dbo].[tblPub_PosPort] (
		[PortId]
	) ON UPDATE CASCADE 
GO


ALTER TABLE dbo.tFacCard WITH NOCHECK ADD CONSTRAINT
	FK_tFacCard_tblPub_Pos FOREIGN KEY
	(
	PosId
	) REFERENCES dbo.tblPub_Pos
	(
	PosId
	) ON UPDATE CASCADE
	
GO



DELETE FROM tblAcc_Bank WHERE tintBank >= 13 
GO

INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 13, -- tintBank - int
          N'پست بانک'  -- nvcBankName - nvarchar(25)
          )
GO

INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 14, -- tintBank - int
          N'توسعه صادرات'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 15, -- tintBank - int
          N'صنعت و معدن'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 16, -- tintBank - int
          N'بانک مسکن'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 17, -- tintBank - int
          N'توسعه تعاون'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 18, -- tintBank - int
          N'کارآفرین'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 19, -- tintBank - int
          N'پاسارگاد'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 20, -- tintBank - int
          N'بانک سرمایه'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 21, -- tintBank - int
          N'بانک سینا'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 22, -- tintBank - int
          N'بانک شهر'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 23, -- tintBank - int
          N'بانک دی'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 24, -- tintBank - int
          N'بانک انصار'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 25, -- tintBank - int
          N'بانک حکمت ایرانیان'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 26, -- tintBank - int
          N'بانک ایران زمین'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 27, -- tintBank - int
          N'بانک قوامین'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 28, -- tintBank - int
          N'بانک خاورمیانه'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 29, -- tintBank - int
          N'بانک آینده'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 30, -- tintBank - int
          N'بانک مهر اقتصاد'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 31, -- tintBank - int
          N'بانک مهر ایران'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 32, -- tintBank - int
          N'بانک رسالت '  -- nvcBankName - nvarchar(25)
          )
GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tBanks]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tBanks]
GO

CREATE PROCEDURE [dbo].[Get_All_tBanks] AS
select * from dbo.tblAcc_Bank ORDER BY tintBank

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tRecvType_Acc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tRecvType_Acc]
GO


CREATE  PROCEDURE [dbo].[Get_All_tRecvType_Acc] AS
SELECT *
FROM tRecvType
WHERE (tintIsShow = 1)
AND tintType = 1 OR tintType = 5
ORDER BY tintType


GO

UPDATE tRecvType SET nvcDescription = N'کارت بانکی' WHERE tintType = 5
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insert_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Insert_tblPub_Pos]
GO


CREATE  PROCEDURE [dbo].[Insert_tblPub_Pos] 
(	@PosType INT ,
	@StationId INT  , 
	@BankNo INT , 
	@nvcAccountNo nvarchar(50) , 
	@AccountId INT ,
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@nvcPosSerialNo NVARCHAR(20) ,
	@intStatus int out)
AS


Begin Tran


Insert Into dbo.tblPub_Pos
        ( 
          StationId ,
          intBank ,
          nvcAccountNo ,
          AccountId ,
          PosType ,
          ComunicationType ,
          PosAddress ,
          nvcPosSerialNo
        )
VALUES  ( 
          @StationId , -- NvcPosNo - nvarchar(20)
          @BankNo , -- BankName - nvarchar(20)
          @nvcAccountNo , -- nvcAccountNo - nvarchar(20)
          @AccountId ,
          @PosType ,
          @CommunicationType ,
          @PosAddress ,
          @nvcPosSerialNo
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

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_tblPub_Pos]
GO

CREATE  PROCEDURE [dbo].[Update_tblPub_Pos] (
	@PosId INT ,
	@StationId INT  , 
	@BankNo INT , 
	@nvcAccountNo nvarchar(50) ,
	@AccountId INT , 
	@PosType INT ,
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@nvcPosSerialNo NVARCHAR(20) ,
	@intStatus int out)

AS

Begin Tran

UPDATE dbo.tblPub_Pos SET
	StationId = @StationId  , 
	intBank = @BankNo , 
	nvcAccountNo = @nvcAccountNo ,
	AccountId = @AccountId ,
	PosType = @PosType ,
	ComunicationType = @CommunicationType ,
	PosAddress = @PosAddress ,
	nvcPosSerialNo = @nvcPosSerialNo

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

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tblPub_Pos_ById]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tblPub_Pos_ById]
GO

CREATE  PROCEDURE [dbo].[Get_tblPub_Pos_ById] 
@PosId INT 
AS
select * from [tblPub_Pos] WHERE  PosId = @PosId

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Delete_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Delete_tblPub_Pos]
GO

CREATE   PROCEDURE [dbo].[Delete_tblPub_Pos](
	@PosId INT )
AS
	DELETE FROM dbo.tblPub_Pos WHERE PosId = @PosId

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tblPub_Pos]
GO

CREATE  PROCEDURE [dbo].[Get_All_tblPub_Pos] AS
select * from [tblPub_Pos] ORDER BY PosId

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_PosType]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_All_PosType
GO

CREATE PROC Get_All_PosType
AS 
select * from [dbo].[tblPub_PosType] ORDER BY PosTypeId

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tblPub_Pos_ByStationId]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tblPub_Pos_ByStationId]
GO

CREATE   PROCEDURE [dbo].[Get_tblPub_Pos_ByStationId] 
@StationId INT 
AS
select * , CAST(PosType AS VARCHAR(2)) AS nvcPosType  from [tblPub_Pos] INNER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblPub_Pos.intBank
INNER JOIN dbo.tblPub_PosType ON dbo.tblPub_PosType.PosTypeId = dbo.tblPub_Pos.PosType
INNER JOIN dbo.tblPub_PosPort ON dbo.tblPub_PosPort.PortId = dbo.tblPub_Pos.ComunicationType
WHERE  StationId = @StationId 
ORDER BY PosId

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tfacm_Card_Detail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tfacm_Card_Detail]
GO

CREATE   PROCEDURE [dbo].[Get_tfacm_Card_Detail]
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

IF COL_LENGTH('tblPub_Pos','BankTafsili') IS NULL
BEGIN
	ALTER TABLE dbo.tblPub_Pos
	ADD [BankTafsili] [int] NULL
END

GO

