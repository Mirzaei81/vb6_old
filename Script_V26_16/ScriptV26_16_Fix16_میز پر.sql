
--ScriptV26_16_Fix16_میز پر.SQL
--95/04/14


-- این فانکشن میزهای پر امروز را شناسایی می کند
ALTER   Function [dbo].FN_NoEmptyTables
(@CurrentDay NVARCHAR(8))

RETURNS  @ReturnTable TABLE(
	[No] INT ,
	Branch INT  ,
	Code BIGINT PRIMARY KEY       ) 
As

BEGIN

    INSERT INTO @ReturnTable 
	SELECT tTable.[No]  , dbo.tTable.Branch , max(intSerialNo) as Code
	FROM   dbo.tFacM INNER JOIN dbo.tTable
	ON dbo.tTable.Branch = dbo.tFacM.Branch AND dbo.tTable.No = dbo.tFacM.TableNo
	where tfacm.Recursive = 0 AND dbo.tFacM.Date = @CurrentDay AND dbo.tFacM.FacPayment = 0 AND Status = 2
	GROUP BY tTable.No  , dbo.tTable.Branch


RETURN 

END



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   PROC GetInvoiceByTable
(
	@Branch int,
	@TableNo int
 )
as
BEGIN

DECLARE @nvcDate NVARCHAR(8)
SET @nvcDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())
	SELECT TOP 1 *  
	FROM  dbo.tFacM 
	WHERE Branch=@Branch 
		AND TableNo=@TableNo
		AND Balance = 0 
		AND [Recursive]<>1 
		AND dbo.tFacM.[Date] = @nvcDate
		AND FacPayment = 0
		AND Status = 2
	ORDER BY intSerialNo desc
END
--EXEC dbo.GetInvoiceByTable @Branch = 3, -- int
--    @TableNO = 42 -- int
--exec GetInvoiceByTable 3,236


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--93/01/28
--برای اینکه در میز گرافیکی دوباره فاکتور ظاهر نشود
-- وقتی انتقال به حساب مشتری اعتباری انجام می شود


ALTER    PROC CheckTableStatus(@TableNo INT , @nvcDate NVARCHAR(8) )
AS
--	SELECT * FROM dbo.tTable WHERE [No]=@TableNO

	SELECT * 
	FROM dbo.tTable 
	WHERE [No]=@TableNo
		AND No IN(
				SELECT TableNo
				FROM dbo.tFacM 
				WHERE dbo.tFacM.Balance = 0 
						AND dbo.tFacM.Recursive<>1 
						AND dbo.tFacM.Branch = dbo.Get_Current_Branch()
						AND dbo.tFacM.[Date] = @nvcDate
						AND tfacm.FacPayment = 0
						AND tfacM.Status = 2
						)
						
						--AND dbo.tFacM.[Date] = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())
						

GO



