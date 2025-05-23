

--ScriptV26_16_Fix16_آپشن منفی.sql

--95/04/23

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Insert_Differences](
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
@LastCode * -1 ,
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

--declare @p5 int
--set @p5=71
--exec Insert_Differences 0,N'آپشن مثبت',N'آپشن منفي',1000,@p5 output
--select @p5
--GO

