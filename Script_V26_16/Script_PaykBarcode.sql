
--کد پرسنلی پیک 4 رقمی شده و بارکد آن اصلاح شد
--برای همه ورژن ها قابل استقاده است


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  FUNCTION dbo.PersonelBarcodeGenerator
(
	@JobID INT,
	@PPNO   INT
)
RETURNS  NVARCHAR(20)

AS

BEGIN


	DECLARE @strJobID    NVARCHAR(10)
	DECLARE @strPPNO     NVARCHAR(10)
	DECLARE @Tmp         NVARCHAR(20)
	DECLARE @ZeroCount   INT


	SET @ZeroCount = 2 - LEN(CAST(@JobID AS NVARCHAR(10)))
	SET @strJobID  = (SELECT dbo.Repeater('0',@ZeroCount)) + CAST(@JobID AS NVARCHAR(10)) 

	SET @ZeroCount = 4 - LEN(CAST(@PPNO AS NVARCHAR(10)))
	SET @strPPNO   = (SELECT dbo.Repeater('0',@ZeroCount)) + CAST(@PPNO AS NVARCHAR(10)) 

	--SET @Tmp = @strJobID + (SELECT dbo.Repeater('0',7)) + @strPPNO   --12 number is correct
	SET @Tmp = @strJobID + (SELECT dbo.Repeater('0',6)) + @strPPNO
	SET @Tmp = '*' + @TMP + '*'
 	RETURN(@Tmp)
END



GO



