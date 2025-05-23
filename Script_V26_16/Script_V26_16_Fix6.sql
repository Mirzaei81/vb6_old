
--Script_V26_16_Fix6
--Kitchen Monitoring
--آدرس موقت
--تغییر کرایه حمل
-- 93/-7/04


INSERT  INTO dbo.tblPub_Script2
        ( Version ,
          Script ,
          LastScriptNo ,
          [Date] ,
          FixNumber
        )
VALUES  ( 26 ,
          16 ,
          0 ,
          dbo.shamsi(GETDATE()) ,
          6
        )
GO



ALTER  PROCEDURE [dbo].[Get_AddedGoods_To_Kitchen](@intLanguage int , @StationID INT , @Branch INT = NULL) AS

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
SELECT  dbo.tGood.Code, dbo.tGood.Level1, 
	case @intLanguage when 0 then dbo.tGoodLevel1.Description
			when 1 then dbo.tGoodLevel1.LatinDescription
			end as DesLevel1 ,
	dbo.tGood.Level2 ,

	case @intLanguage when 0 then dbo.tGoodLevel2.Description
			when 1 then dbo.tGoodLevel2.LatinDescription
			end as DesLevel2 ,

	case @intLanguage when 0 then dbo.tGood.Name
			when 1 then dbo.tGood.LatinName
			end as [Name] ,

	case @intLanguage when 0 then dbo.tGood.NamePrn
			when 1 then dbo.tGood.LatinNamePrn
			end as [NamePrn] ,
     

	                                         	      
	dbo.tGood.TechnicalNo, dbo.tGood.BarCode, dbo.tGood.Unit,
	dbo.tGood.Model, dbo.tGood.Weight, 
	dbo.tGood.NumberOfUnit, 
	dbo.tGood.ProductCompany, dbo.tGood.SellPrice, dbo.tGood.BuyPrice, 
	dbo.tGood.BtnAscDefault, 
	dbo.tGood.BtnTz1No, dbo.tGood.GoodType

FROM        dbo.tGood INNER JOIN
                      dbo.tGoodLevel1 ON dbo.tGood.Level1 = dbo.tGoodLevel1.Code INNER JOIN
                      dbo.tGoodLevel2 ON dbo.tGood.Level2 = dbo.tGoodLevel2.Code

where dbo.tGood.Code not in (select GoodCode from tKitchenGood where StationID = @StationID  and Branch = @Branch)




GO

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO




ALTER  PROCEDURE [dbo].[Get_DeletedGoods_From_Kitchen](@intLanguage int , @StationID INT , @Branch INT = NULL) AS

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()

SELECT  dbo.tGood.Code, dbo.tGood.Level1, 
	case @intLanguage when 0 then dbo.tGoodLevel1.Description
			when 1 then dbo.tGoodLevel1.LatinDescription
			end as DesLevel1 ,
	dbo.tGood.Level2 ,

	case @intLanguage when 0 then dbo.tGoodLevel2.Description
			when 1 then dbo.tGoodLevel2.LatinDescription
			end as DesLevel2 ,

	case @intLanguage when 0 then dbo.tGood.Name
			when 1 then dbo.tGood.LatinName
			end as [Name] ,

	case @intLanguage when 0 then dbo.tGood.NamePrn
			when 1 then dbo.tGood.LatinNamePrn
			end as [NamePrn] ,
                                              
	dbo.tGood.TechnicalNo, dbo.tGood.BarCode, dbo.tGood.Unit,
	dbo.tGood.Model, dbo.tGood.Weight, 
	dbo.tGood.NumberOfUnit, 
	dbo.tGood.ProductCompany, dbo.tGood.SellPrice, dbo.tGood.BuyPrice, 
	dbo.tGood.BtnAscDefault,  
	dbo.tGood.BtnTz1No, dbo.tGood.GoodType, dbo.tKitchenGood.StationID

FROM         dbo.tGood INNER JOIN
                      dbo.tKitchenGood ON dbo.tGood.Code = dbo.tKitchenGood.GoodCode INNER JOIN
                      dbo.tStations ON dbo.tKitchenGood.StationID = dbo.tStations.StationID INNER JOIN
                      dbo.tGoodLevel1 ON dbo.tGood.Level1 = dbo.tGoodLevel1.Code INNER JOIN
                      dbo.tGoodLevel2 ON dbo.tGood.Level2 = dbo.tGoodLevel2.Code

where  dbo.tStations.StationID = @StationID and  dbo.tKitchenGood.Branch =  @Branch




GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_Station_By_StationType](@StationType INT , @Branch INT = NULL) 
AS

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
SELECT [StationID], [PortCode], [Description], [CashNo], 
	[IsActive], [IP], [Dir], [Machine_Name],  
	[StationType], [Branch] 
FROM [tStations]
where ( (StationType & @StationType) = @StationType  AND Branch = @Branch)




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Delete_DeletedGoods_From_Kitchen] ( @StationID int , @GoodCode INT , @Branch INT = NULL) AS 
IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
delete from dbo.tKitchenGood
where StationID = @StationID and branch = @Branch
AND GoodCode = @GoodCode


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER PROCEDURE [dbo].[Insert_DeletedGoods_From_Kitchen]( @StationID int , @GoodCode INT , @Branch INT = NULL) AS
IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
Insert Into  tKitchenGood (StationID , GoodCode , Branch )
Values
(@StationID , @GoodCode ,  @Branch )

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   Proc dbo.Get_Stations  (@MaxStationNo int ,
				@MaxPocketPcNo int ,
				@MaxKitchenNo int)
as 
--Declare @S nvarchar(4000)
--set @S = '(select Top ' + cast (@MaxStationNo as nvarchar(50)) + 
--	' * from dbo.tStations  Where 
--	 ((StationType & 1) = 1 or  (StationType & 2 = 2))  And IsActive = 1 )
--	Union
--	(select Top ' + cast (@MaxPocketPcNo as nvarchar(50)) + 
--	' * from dbo.tStations  Where 
--	 (StationType & 8) = 8  And IsActive = 1)
--	Union
--	(select Top ' + cast (@MaxTz1No as nvarchar(50)) + 
--	' * from dbo.tStations  Where 
--	 (StationType & 4) = 4 And IsActive = 1)'
--Exec ( @S)

SELECT * FROM dbo.tStations WHERE (((StationType & 1) = 1 or  (StationType & 2 = 2))  And IsActive = 1 )
 UNION
SELECT * FROM dbo.tStations WHERE ((StationType & 8) = 8  And IsActive = 1)
 UNION 
SELECT * FROM dbo.tStations WHERE ((StationType & 4) = 4  And IsActive = 1)
 UNION 
SELECT * FROM dbo.tStations WHERE ((StationType & 16) = 16  And IsActive = 1)

GO







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



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  Procedure dbo.Update_Cust_By_NewCarryFee  
(   
 @OldCarryFee Float,   
 @NewCarryFee Float, 
 @PercentCarryFee  FLOAT,
 @User int,   
 @Updated Bigint out  

)  

as  

Begin Tran  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

IF @PercentCarryFee = 0 
	BEGIN 
		Update dbo.tCust SET  
		 CarryFee = @NewCarryFee ,  
		 [Date] = @Date ,  
		 [Time] = @Time ,  
		 [User] = @User   
		Where CarryFee = @OldCarryFee  

		if @@Error <> 0   
		 goto ErrHandler  
	END
ELSE
		Update dbo.tCust SET  
		 CarryFee = CarryFee + CAST(CarryFee * @PercentCarryFee / 100 AS INT ) ,  
		 [Date] = @Date ,  
		 [Time] = @Time ,  
		 [User] = @User   

		if @@Error <> 0   
		 goto ErrHandler  

Set @Updated = 1   

Commit Tran   

return @Updated  

ErrHandler:  
RollBack Tran  
return -1  




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  Procedure dbo.Update_Cust_By_NewPaykFee  
(   
 @OldPaykFee Float,   
 @NewPaykFee Float,   
 @PercentPaykFee  FLOAT,
 @User int,   
 @Updated Bigint out  

)  

as  

Begin Tran  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

IF @PercentPaykFee = 0
BEGIN 
Update dbo.tCust SET  
 PaykFee = @NewPaykFee ,  
 [Date] = @Date ,  
 [Time] = @Time ,  
 [User] = @User   
Where PaykFee = @OldPaykFee  

if @@Error <> 0   
 goto ErrHandler  
END 
ELSE
Update dbo.tCust SET  
 PaykFee = PaykFee + CAST(PaykFee * @PercentPaykFee / 100 AS INT ) ,  
 [Date] = @Date ,  
 [Time] = @Time ,  
 [User] = @User   
if @@Error <> 0   
 goto ErrHandler  

Set @Updated = 1   
--   

Commit Tran   

return @Updated  

ErrHandler:  
RollBack Tran  
return -1  





GO


