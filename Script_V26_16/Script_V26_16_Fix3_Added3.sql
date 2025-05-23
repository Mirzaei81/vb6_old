
--Script V26_16_Fix3_Added3
-- فاکتور خرید حواله روزانه

--92/11/30



drop INDEX tfacm.[IX_Branch_Status_Date_Shift_TempNo]   

GO 


ALTER   Proc Get_New_FacM_No ( @Status int, @AccountYear smallint, @Branch INT )


AS

	DECLARE @No INT 
	DECLARE @TempNo INT 
	DECLARE @ShiftNo INT 
	DECLARE @Date NVARCHAR(8)  ---problem with miladi
	SET @ShiftNo= dbo.Get_Shift(GETDATE())     
	SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      

	Select @No = isnull(max([No]),0)+ 1 From tFacM  Where  Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear 

IF @Status = 2 OR @Status = 5 
	Select @TempNo =  isnull(max([TempNo]),0)+ 1 From tFacM  Where  Status = @Status and Branch =  @Branch AND Date = @Date  AND shiftNo = @ShiftNo
ELSE 
	SET @TempNo = 0

SELECT @No AS [No] , @TempNo AS TempNo


GO



