
ALTER    PROCEDURE [dbo].[Insert_tPrinters_Log]
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
@PrintPosition BIT ,
@PingStatus VARCHAR(255) ,
@Result INT OUT 
    )
AS 

INSERT INTO dbo.tPrinters_Log
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
          PrintPosition ,
          PingStatus ,
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
          @PrintPosition ,
          @PingStatus ,
          GETDATE()
        )
        
	SET @Result = @@IDENTITY 
        RETURN @Result
GO


