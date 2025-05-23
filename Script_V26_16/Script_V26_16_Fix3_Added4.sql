

--Script V26_16_Fix3_Added4
--93/01/28
--برای اینکه در میز گرافیکی دوباره فاکتور ظاهر نشود
-- وقتی انتقال به حساب مشتری اعتباری انجام می شود
-- در فرم جستجو کاربران میزهای پارتیشن دیگررااگردسترسی ندارتد نبینند

ALTER   PROC CheckTableStatus(@TableNo INT , @nvcDate NVARCHAR(8) )
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
						)
						
						--AND dbo.tFacM.[Date] = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())
						

GO



ALTER  PROC GetInvoiceByTable
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
	ORDER BY intSerialNo desc
END
--EXEC dbo.GetInvoiceByTable @Branch = 3, -- int
--    @TableNO = 42 -- int
--exec GetInvoiceByTable 3,236


GO



ALTER  VIEW dbo.vw_Invoice_Table
AS
SELECT     tFacM.intSerialNo, tFacM.[No], tFacM.Status, tFacM.Owner, tFacM.Customer, tFacM.SumPrice, tFacM.Recursive, tFacM.InCharge, tFacM.FacPayment, 
              tFacM.OrderType, tFacM.ServePlace, tFacM.StationId, tFacM.BascoleNo, tFacM.ShiftNo, tFacM.TableNo, tFacD.intInventoryNo, tFacM.[Date], 
              tFacM.[Time], tFacM.[User], tFacM.RegDate, dbo.tShift.[Description] AS ShiftDescription, dbo.tStations.[Description] AS StationDescription, 
              dbo.tOrderType.[Description] AS OrderTypeDescription, dbo.tOrderType.LatinDescription AS OrderTypeLatinDescription, 
              dbo.tServePlace.[Description] AS ServePlaceDescription, dbo.tServePlace.LatinDescription AS ServePlaceLatinDescription, 
              dbo.tTable.[Name] AS TableName, dbo.tPer.nvcFirstName + SPACE(2) + dbo.tPer.nvcSurName AS FullName, dbo.tPer.pPno, 
               tFacM.DiscountTotal   , tFacM.CarryFeeTotal 
             ,  tFacM.ServiceTotal   ,  tFacM.PackingTotal ,   tFacM.Balance , ISNULL(tfacm.GuestNo , '') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.No) AS TempNo
              , ISNULL(dbo.tTable.PartitionID , 0) AS PartitionID
              FROM          tFacM INNER JOIN
			  dbo.tfacD ON tFacM.intserialno = dbo.tfacD.intserialno AND tFacM.Branch = dbo.tfacD.Branch inner join
              dbo.tOrderType ON tFacM.OrderType = dbo.tOrderType.Code INNER JOIN
              dbo.tServePlace ON tFacM.ServePlace = dbo.tServePlace.intServePlace INNER JOIN
              dbo.tShift ON tFacM.ShiftNo = dbo.tShift.Code AND tFacM.Branch = dbo.tShift.Branch LEFT OUTER JOIN
              dbo.tTable ON tFacM.TableNo = dbo.tTable.[No] AND tFacM.Branch = dbo.tTable.Branch Left Outer JOIN
              dbo.tPer ON tFacM.[Incharge] = dbo.tPer.pPno  INNER JOIN
          
              dbo.tStations ON tFacM.StationId = dbo.tStations.StationId AND tFacM.Branch = dbo.tStations.Branch
	--WHERE     tfacm.Branch = dbo.Get_Current_Branch()





GO


ALTER  Proc Get_Factors_Tables ( @intLanguage int , @nvcDate NVARCHAR(8) , @Branch INT )

as

SELECT distinct   Vw_Invoice_Table.TableNo ,  Vw_Invoice_Table.[Date] ,  Vw_Invoice_Table.TableName ,  Vw_Invoice_Table.FullName ,  
                   Vw_Invoice_Table.SumPrice ,  Vw_Invoice_Table.[Time] ,  Vw_Invoice_Table.[No] ,  Vw_Invoice_Table.intSerialNo
		, Vw_Invoice_Table.GuestNo , Vw_Invoice_Table.TempNo , Vw_Invoice_Table.ShiftDescription , Vw_Invoice_Table.FacPayment
		, Vw_Invoice_Table.PartitionId
FROM         Vw_Invoice_Table

Where  Vw_Invoice_Table.[TableNo] >0  and Vw_Invoice_Table.[Date] = @nvcDate  And Vw_Invoice_Table.Recursive = 0 --And Vw_Invoice_Table.Facpayment = 0 
ORDER BY Vw_Invoice_Table.TableNo --Vw_Invoice_Table.[No] , Vw_Invoice_Table.[Date] ,Vw_Invoice_Table.[Time]




GO


