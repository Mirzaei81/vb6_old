
--Script_V26_16_Fix3_Added2

ALTER  Procedure Get_DailyCountDeliveredFactors(@Today Nvarchar(8), @Count int out )
AS
DECLARE @Branch INT 
SET @Branch = dbo.Get_Current_Branch()
Set @Count = ( Select Count(*) from vw_NotPaidFactors where job = 3 and [Date] = @Today AND Branch =  @Branch )



GO



ALTER  Procedure Get_DailyCountNotDeliveredFactors(@Today Nvarchar(8), @Count int out )
as
DECLARE @Branch INT 
SET @Branch = dbo.Get_Current_Branch()
Set @Count =
 ( select count( intserialno ) from 	 dbo.tFacM  
	
	          WHERE   ( dbo.tFacM.Incharge IS NULL OR dbo.tFacM.Incharge = '' ) 
			AND dbo.tFacM.Status = 2									  
			AND dbo.tFacM.facPayment = 0 	
            And dbo.tfacM.ServePlace = 2
            And dbo.tfacm.Recursive <> 1
			AND dbo.tFacM.TableNo is null 
			AND dbo.tFacM.[Date] = @Today  --dbo.Get_ShamsiDate_For_Current_Shift(Getdate())
			AND dbo.tFacM.Branch =  @Branch
			AND dbo.tFacM.Customer > 0 )

	



GO


