ALTER   PROCEDURE [dbo].[Get_Previous_Factor_Detail] (@intLanguage  int , @Code INT , @Branch INT= NULL  ) AS

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
SELECT     dbo.tFacD2.*, case @intLanguage when  0 then dbo.tGood.Name 
when 1 then dbo.tGood.LatinName end AS Name
FROM         dbo.tFacD2 INNER JOIN
                      dbo.tGood ON dbo.tFacD2.GoodCode = dbo.tGood.Code
where tFacD2.Code = @Code and Branch =  @Branch 

GO
