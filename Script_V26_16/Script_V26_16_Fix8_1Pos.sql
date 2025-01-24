
--Script_V26_16_Fix8_Pos
--93/07/28

-- Date  : 93/06/24
-- Add Columns To tblPub_Pos For AryaPos
-- Add Stored Procedure Update tblPub_Pos For update pos config

ALTER TABLE dbo.tblPub_Pos
ALTER COLUMN nvcAccountNo NVARCHAR(50)
GO

ALTER TABLE dbo.tblPub_Pos
ALTER COLUMN nvcBankName NVARCHAR(50)
GO


IF COL_LENGTH('tblPub_Pos','PosType') IS NULL
BEGIN
	ALTER TABLE tblPub_Pos
	ADD PosType INT NULL
END

GO

IF COL_LENGTH('tblPub_Pos','ComunicationType') IS NULL
BEGIN
	ALTER TABLE tblPub_Pos
	ADD ComunicationType INT NULL
END

GO

IF COL_LENGTH('tblPub_Pos','PosAddress') IS NULL
BEGIN
	ALTER TABLE tblPub_Pos
	ADD PosAddress NVARCHAR(50) NULL
END

GO

IF COL_LENGTH('tblPub_Pos','PosPort') IS NULL
BEGIN
	ALTER TABLE tblPub_Pos
	ADD PosPort INT NULL
END

GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sp_Pos_Update_DeviceSetting]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sp_Pos_Update_DeviceSetting]
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE Procedure [dbo].[Sp_Pos_Update_DeviceSetting]
	@PosId INT,
	@PosType INT,
	@ComunicationType INT,
	@PosAddress nvarchar(50),
	@PosPort INT
As
Begin

	UPDATE dbo.tblPub_Pos 
	SET 
	PosAddress = @PosAddress,
	PosPort = @PosPort,
	PosType = @PosType,
	ComunicationType = @ComunicationType 
	WHERE PosId = @PosId
	
	RETURN 1
End
GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   PROCEDURE [dbo].[Insert_tblPub_Pos] 
(
	@PosId INT ,
	@NvcPosNo nvarchar(20) , 
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
          NvcPosNo ,
          nvcBankName ,
          nvcAccountNo ,
          AccountId ,
          PosType ,
          ComunicationType ,
          PosAddress
        )
VALUES  ( @PosId , -- PosId - int
          @NvcPosNo , -- NvcPosNo - nvarchar(20)
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
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@intStatus int out)

AS

Begin Tran

UPDATE dbo.tblPub_Pos SET
	PosId = @NewPosId ,
	NvcPosNo = @NvcPosNo  , 
	nvcBankName = @nvcBankName , 
	nvcAccountNo = @nvcAccountNo ,
	AccountId = @AccountId ,
	PosType = @NewPosId ,
	ComunicationType = @CommunicationType ,
	PosAddress = @PosAddress

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










