


--اضافه کردن آدرس موقت به فیش موقت
--Script_V26_16_Fix6_Added1


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER   procedure dbo.Get_TempFactors(@intLanguage int , @Status INT,  @Branch INT  )
AS
If @Status = 1
Begin

SELECT     	dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tFacMTemp.intSerialNo, dbo.tFacMTemp.[No], dbo.tFacMTemp.Status, dbo.tFacMTemp.Owner, 
		dbo.tFacMTemp.Customer, dbo.tFacMTemp.DiscountTotal, dbo.tFacMTemp.CarryFeeTotal, dbo.tFacMTemp.SumPrice, 
		dbo.tFacMTemp.Recursive, dbo.tFacMTemp.InCharge, dbo.tFacMTemp.FacPayment, dbo.tFacMTemp.OrderType, dbo.tFacMTemp.ServePlace, 
		dbo.tFacMTemp.StationID, dbo.tFacMTemp.ServiceTotal, dbo.tFacMTemp.PackingTotal, dbo.tFacMTemp.ShiftNo, 
		dbo.tFacMTemp.TableNo, dbo.tFacMTemp.NvcDescription, dbo.tFacMTemp.[Date], dbo.tFacMTemp.[Time], dbo.tFacMTemp.[User], 
		dbo.tFacMTemp.RegDate, dbo.tSupplier.MembershipId, dbo.tSupplier.Address,
		case When  dbo.tSupplier.Name+dbo.tSupplier.Family <>'' Then  dbo.tSupplier.Name+ ' ' + dbo.tSupplier.Family 
		Else dbo.tSupplier.WorkName end as FullName ,
		case @intLanguage when 0 then dbo.tShift.Description when 1 then dbo.tShift.LatinDescription end AS ShiftDescription
		
FROM         	dbo.tFacMTemp INNER JOIN
		dbo.tUser ON dbo.tFacMTemp.[User] = dbo.tUser.UID and  dbo.tFacMTemp.[Branch] = dbo.tUser.Branch INNER JOIN
		dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and  dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
		dbo.tSupplier ON dbo.tFacMTemp.Owner = dbo.tSupplier.Code AND  (dbo.tFacMTemp.Branch = dbo.tSupplier.Branch  or tSupplier.Branch IS NULL ) left outer join
		dbo.tShift ON dbo.tFacMTemp.ShiftNo = dbo.tShift.Code and  dbo.tFacMTemp.Branch = dbo.tShift.Branch   
		
Where  dbo.tFacMTemp.Branch =  @Branch And dbo.tFacMTemp.Status = @Status
order by dbo.tFacMTemp.[Date] , dbo.tFacMTemp.[Time]


End

else If @Status = 4

Begin

SELECT     	dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tFacMTemp.intSerialNo, dbo.tFacMTemp.[No], dbo.tFacMTemp.Status, dbo.tFacMTemp.Owner, 
		dbo.tFacMTemp.Customer, dbo.tFacMTemp.DiscountTotal, dbo.tFacMTemp.CarryFeeTotal, dbo.tFacMTemp.SumPrice, 
		dbo.tFacMTemp.Recursive, dbo.tFacMTemp.InCharge, dbo.tFacMTemp.FacPayment, dbo.tFacMTemp.OrderType, dbo.tFacMTemp.ServePlace, 
		dbo.tFacMTemp.StationID, dbo.tFacMTemp.ServiceTotal, dbo.tFacMTemp.PackingTotal, dbo.tFacMTemp.ShiftNo, 
		dbo.tFacMTemp.TableNo, dbo.tFacMTemp.NvcDescription, dbo.tFacMTemp.[Date], dbo.tFacMTemp.[Time], dbo.tFacMTemp.[User], 
		dbo.tFacMTemp.RegDate, dbo.tSupplier.MembershipId, dbo.tSupplier.Address,
		case When  dbo.tSupplier.Name+dbo.tSupplier.Family <>'' Then  dbo.tSupplier.Name+ ' ' + dbo.tSupplier.Family 
		Else dbo.tSupplier.WorkName end as FullName ,
		case @intLanguage when 0 then dbo.tShift.Description when 1 then dbo.tShift.LatinDescription end AS ShiftDescription
		, dbo.tFacMTemp.TempAddress
		
FROM         	dbo.tFacMTemp INNER JOIN
		dbo.tUser ON dbo.tFacMTemp.[User] = dbo.tUser.UID and  dbo.tFacMTemp.[Branch] = dbo.tUser.Branch INNER JOIN
		dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and  dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
		dbo.tSupplier ON dbo.tFacMTemp.Owner = dbo.tSupplier.Code AND ( dbo.tFacMTemp.Branch = dbo.tSupplier.Branch   or tSupplier.Branch IS NULL ) left outer join
		dbo.tShift ON dbo.tFacMTemp.ShiftNo = dbo.tShift.Code and  dbo.tFacMTemp.Branch = dbo.tShift.Branch   
		
Where  dbo.tFacMTemp.Branch =   @Branch And dbo.tFacMTemp.Status = @Status
order by dbo.tFacMTemp.[Date] , dbo.tFacMTemp.[Time]


End

Else If @Status = 2 
Begin
SELECT     	dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tFacMTemp.intSerialNo, dbo.tFacMTemp.[No], dbo.tFacMTemp.Status, dbo.tFacMTemp.Owner, 
		dbo.tFacMTemp.Customer, dbo.tFacMTemp.DiscountTotal, dbo.tFacMTemp.CarryFeeTotal, dbo.tFacMTemp.SumPrice, 
		dbo.tFacMTemp.Recursive, dbo.tFacMTemp.InCharge, dbo.tFacMTemp.FacPayment, dbo.tFacMTemp.OrderType, dbo.tFacMTemp.ServePlace, 
		dbo.tFacMTemp.StationID, dbo.tFacMTemp.ServiceTotal, dbo.tFacMTemp.PackingTotal, dbo.tFacMTemp.ShiftNo, 
		dbo.tFacMTemp.TableNo, dbo.tFacMTemp.NvcDescription , dbo.tFacMTemp.[Date], dbo.tFacMTemp.[Time], dbo.tFacMTemp.[User], 
		dbo.tFacMTemp.RegDate, dbo.tCust.MembershipId, dbo.tCust.Address,
		case When  dbo.tCust.Name+dbo.tCust.Family <>'' Then  dbo.tCust.Name+ ' ' + dbo.tCust.Family 
		Else dbo.tCust.WorkName end as FullName ,
		case @intLanguage when 0 then dbo.tShift.Description when 1 then dbo.tShift.LatinDescription end AS ShiftDescription
		, dbo.tFacMTemp.TempAddress
		
FROM         	dbo.tFacMTemp INNER JOIN
		dbo.tUser ON dbo.tFacMTemp.[User] = dbo.tUser.UID and  dbo.tFacMTemp.[Branch] = dbo.tUser.Branch INNER JOIN
		dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and  dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
		dbo.tCust ON dbo.tFacMTemp.Customer = dbo.tCust.Code AND  (dbo.tFacMTemp.Branch = dbo.tCust.Branch   or tCust.Branch IS NULL ) left outer join
		dbo.tShift ON dbo.tFacMTemp.ShiftNo = dbo.tShift.Code and  dbo.tFacMTemp.Branch = dbo.tShift.Branch   
		
Where  dbo.tFacMTemp.Branch =   @Branch And dbo.tFacMTemp.Status = @Status
order by dbo.tFacMTemp.[Date] , dbo.tFacMTemp.[Time]End

Else If @Status = 5 
Begin
SELECT     	dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tFacMTemp.intSerialNo, dbo.tFacMTemp.[No], dbo.tFacMTemp.Status, dbo.tFacMTemp.Owner, 
		dbo.tFacMTemp.Customer, dbo.tFacMTemp.DiscountTotal, dbo.tFacMTemp.CarryFeeTotal, dbo.tFacMTemp.SumPrice, 
		dbo.tFacMTemp.Recursive, dbo.tFacMTemp.InCharge, dbo.tFacMTemp.FacPayment, dbo.tFacMTemp.OrderType, dbo.tFacMTemp.ServePlace, 
		dbo.tFacMTemp.StationID, dbo.tFacMTemp.ServiceTotal, dbo.tFacMTemp.PackingTotal, dbo.tFacMTemp.ShiftNo, 
		dbo.tFacMTemp.TableNo, dbo.tFacMTemp.NvcDescription, dbo.tFacMTemp.[Date], dbo.tFacMTemp.[Time], dbo.tFacMTemp.[User], 
		dbo.tFacMTemp.RegDate, dbo.tCust.MembershipId, dbo.tCust.Address,
		case When  dbo.tCust.Name+dbo.tCust.Family <>'' Then  dbo.tCust.Name+ ' ' + dbo.tCust.Family 
		Else dbo.tCust.WorkName end as FullName ,
		case @intLanguage when 0 then dbo.tShift.Description when 1 then dbo.tShift.LatinDescription end AS ShiftDescription
		, dbo.tFacMTemp.TempAddress
		
FROM         	dbo.tFacMTemp INNER JOIN
		dbo.tUser ON dbo.tFacMTemp.[User] = dbo.tUser.UID and  dbo.tFacMTemp.[Branch] = dbo.tUser.Branch INNER JOIN
		dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and  dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
		dbo.tCust ON dbo.tFacMTemp.Customer = dbo.tCust.Code AND  (dbo.tFacMTemp.Branch = dbo.tCust.Branch   or tCust.Branch IS NULL )  left outer join
		dbo.tShift ON dbo.tFacMTemp.ShiftNo = dbo.tShift.Code and  dbo.tFacMTemp.Branch = dbo.tShift.Branch   
		
Where  dbo.tFacMTemp.Branch =   @Branch And dbo.tFacMTemp.Status = @Status
order by dbo.tFacMTemp.[Date] , dbo.tFacMTemp.[Time]End

Else 
Begin

SELECT     	dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tFacMTemp.intSerialNo, dbo.tFacMTemp.[No], dbo.tFacMTemp.Status, dbo.tFacMTemp.Owner, 
		dbo.tFacMTemp.Customer, dbo.tFacMTemp.DiscountTotal, dbo.tFacMTemp.CarryFeeTotal, dbo.tFacMTemp.SumPrice, 
		dbo.tFacMTemp.Recursive, dbo.tFacMTemp.InCharge, dbo.tFacMTemp.FacPayment, dbo.tFacMTemp.OrderType, dbo.tFacMTemp.ServePlace, 
		dbo.tFacMTemp.StationID, dbo.tFacMTemp.ServiceTotal, dbo.tFacMTemp.PackingTotal, dbo.tFacMTemp.ShiftNo, 
		dbo.tFacMTemp.TableNo, dbo.tFacMTemp.NvcDescription, dbo.tFacMTemp.[Date], dbo.tFacMTemp.[Time], dbo.tFacMTemp.[User], 
		dbo.tFacMTemp.RegDate, -1 As MembershipId ,' ' as FullName , ' ' As Address ,
		case @intLanguage when 0 then dbo.tShift.Description when 1 then dbo.tShift.LatinDescription end AS ShiftDescription
		, dbo.tFacMTemp.TempAddress
		
FROM         	dbo.tFacMTemp INNER JOIN
		dbo.tUser ON dbo.tFacMTemp.[User] = dbo.tUser.UID and  dbo.tFacMTemp.[Branch] = dbo.tUser.Branch INNER JOIN
		dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and  dbo.tUser.Branch = dbo.tPer.Branch 
		  left outer join
		dbo.tShift ON dbo.tFacMTemp.ShiftNo = dbo.tShift.Code and  dbo.tFacMTemp.Branch = dbo.tShift.Branch   
		
Where  dbo.tFacMTemp.Branch =   @Branch And dbo.tFacMTemp.Status = @Status
order by dbo.tFacMTemp.[Date] , dbo.tFacMTemp.[Time]


End


GO



