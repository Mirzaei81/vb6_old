
--ScriptV26_16_Fix17_ucepercent_950810 .sql

-- نمایش کالاهای واسطه در فرم جستجوی کالا در حالت خریدنی

ALTER   PROCEDURE dbo.Get_GoodInfo_By_GoodType (@intLanguage int  , @NotSupportedGoodType int , @StationId Int ,@Flag Bit , @AccountYear Smallint)
AS

DECLARE @Branch INT 
SET @Branch = (SELECT TOP 1 Branch FROM dbo.tStations WHERE StationID = @StationId )

If @Flag = 0
Begin
SELECT 
	DISTINCT  vw_Good.* , 	Case @intLanguage When 0 then [Name]
						Else [LatinName] end as [GoodName],
				Case @intLanguage When 0 then  [NamePrn]
						Else [LatinNamePrn] End as [NamePrn]

FROM [dbo].[vw_Good]
Inner Join tInventory_Level1 On tInventory_Level1.Level1 = vw_Good.Level1 
Inner Join tInventory On tInventory_Level1.Branch = tInventory.Branch And  tInventory_Level1.InventoryNo = tInventory.InventoryNo
Inner Join tStation_Inventory_Good On tStation_Inventory_Good.InventoryNo = tInventory.InventoryNo 
	And tStation_Inventory_Good.StationId = @StationId 
	And tStation_Inventory_Good.Branch = @Branch
	And tStation_Inventory_Good.AccountYear = @AccountYear
	And tStation_Inventory_Good.GoodCode = vw_Good.Code
Where GoodType <> @NotSupportedGoodType And GoodType <> 4
Order By [Name]

End

Else
Begin
SELECT 
	 vw_Good.* , 	Case @intLanguage When 0 then [Name]
						Else [LatinName] end as [GoodName],
				Case @intLanguage When 0 then  [NamePrn]
						Else [LatinNamePrn] End as [NamePrn]

FROM [dbo].[vw_Good]
Where GoodType <> @NotSupportedGoodType --And GoodType <> 4
Order By [Name]

End



GO

