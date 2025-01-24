SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[GoodLable]
    (
      @FichNo  INT  ,
      @StrGood  NVARCHAR(250) ,
      @StrDescription NVARCHAR(250) 
    )
AS

DECLARE @TempNo INT 
DECLARE @Serveplace NVARCHAR(50) 
SELECT @TempNo = TempNo , @Serveplace = dbo.tServePlace.Description FROM dbo.tFacM
	INNER JOIN dbo.tServePlace ON dbo.tServePlace.intServePlace = dbo.tFacM.ServePlace
 WHERE No = @FichNo AND Status = 2 AND AccountYear = dbo.Get_AccountYear() AND Branch = dbo.Get_Current_Branch() 

    SELECT  
      @FichNo AS FichNo , 
      LTRIM(rtrim(@StrGood)) AS StrGood , 
      LTRIM(RTRIM(@StrDescription)) AS StrDescription ,
      ISNULL(@TempNo , 0) AS TempNo ,
      LTRIM(RTRIM(@Serveplace)) AS Serveplace 
      
      
    
--===============================================



GO
