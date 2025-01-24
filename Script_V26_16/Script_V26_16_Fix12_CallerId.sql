SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Insert_tblTotal_CallerId] 
@nvcDate NVARCHAR(8) ,
@LineNumber SMALLINT ,
@nvcCallerId NVARCHAR(20)


 AS

DECLARE @newtime nvarchar(5)      
select @newtime=dbo.setTimeFormat(getdate())      

DECLARE @intRow INT 
SELECT @intRow = ISNULL(MAX(introw) ,0) + 1 FROM dbo.tblTotal_CallerId WHERE nvcDate = @nvcDate
INSERT INTO dbo.tblTotal_CallerId
        ( intRow ,
          nvcDate ,
          nvcTime ,
          nvcCallerId ,
          LineNumber 
        )
VALUES  ( @intRow , -- intRow - int
          @nvcDate , -- nvcDate - nvarchar(8)
          @newtime , -- nvcTime - nvarchar(8)
          @nvcCallerId , -- nvCallerId - nvarchar(20)
          @LineNumber 
        )

GO
