

--Print Server Log & Print
--V26_16


DROP TABLE tblTotal_PrintFich
Go

CREATE TABLE [dbo].tblTotal_PrintFich
(
[intPrintFichNo] [int] NOT NULL IDENTITY(1, 1),
[intLanguage] [int] NOT NULL,
[FichNumber] [int] NOT NULL,
[PrintFormat] [int] NOT NULL,
[intStationId] [int] NOT NULL,
[Status] [int] NOT NULL,
[PrinterNo] [int] NOT NULL,
[AddEditMode] [int] NOT NULL,
[AccountYear] [int] NOT NULL,
[PartitionNo] [int] NOT NULL,
[Branch] [int] NOT NULL,
[IsPrinted] [bit] NOT NULL,
[nvcTime] [datetime] NOT NULL ,
[nvcError] NVARCHAR(255) NULL 
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[tblTotal_PrintFich] ADD CONSTRAINT [PK_tblTotal_PrintFich] PRIMARY KEY CLUSTERED  ([intPrintFichNo]) WITH FILLFACTOR=90 ON [PRIMARY]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    PROCEDURE [dbo].Insert_tblTotal_PrintFich
    (
@intLanguage [int] ,
@FichNumber [int] ,
@PrintFormat [int] ,
@intStationId [int] ,
@Status [int] ,
@PrinterNo [int] ,
@AddEditMode [int] ,
@AccountYear [int] ,
@PartitionNo [int] ,
@Branch [int] ,
@intPrintFichNo INT OUT 
    )
AS 
Begin Tran

INSERT INTO dbo.tblTotal_PrintFich
        ( intLanguage ,
          FichNumber ,
          PrintFormat ,
          intStationId ,
          [Status] ,
          PrinterNo ,
          AddEditMode ,
          AccountYear ,
          PartitionNo ,
          Branch ,
          IsPrinted ,
          nvcTime
        )
VALUES  ( @intLanguage , -- intLanguage - int
          @FichNumber , -- FichNumber - int
          @PrintFormat , -- PrintFormat - int
          @intStationId , -- intStationId - int
          @Status , -- Status - int
          @PrinterNo , -- PrinterNo - int
          @AddEditMode , -- AddEditMode - int
          @AccountYear , -- AccountYear - int
          @PartitionNo , -- PartitionNo - int
          @Branch , -- Branch - int 
          0 ,
          GETDATE()
        )

if @@Error <> 0 
	Goto ErrHandler

Commit Tran
SET @intPrintFichNo=@@IDENTITY
Return @intPrintFichNo

ErrHandler:
RollBack Tran
SET @intPrintFichNo=-1
Return @intPrintFichNo


GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Update_tblTotal_printFich
	(
	 @intPrintFichNo INT ,
	 @TypeFlag BIT ,
	 @nvcError NVARCHAR(255)
	 )

AS

IF @TypeFlag = 0
	UPDATE [dbo].[tblTotal_PrintFich]  
		SET IsPrinted=1 , nvcError = @nvcError
	 	WHERE intPrintFichNo=@intPrintFichNo 

ELSE
	DELETE FROM  [dbo].[tblTotal_PrintFich]  
	 	WHERE intPrintFichNo=@intPrintFichNo 
	

Return 1



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE dbo.Get_tblTotal_PrintFich 

AS

SELECT   * 	FROM      dbo.tblTotal_PrintFich
		WHERE  IsPrinted = 0
		ORDER BY intPrintFichNo

GO

