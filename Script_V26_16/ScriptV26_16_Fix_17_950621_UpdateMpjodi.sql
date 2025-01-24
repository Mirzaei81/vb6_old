
--ScriptV26_16_Fix_17_950621_UpdateMpjodi.sql
-- 95/06/21

ALTER  PROCEDURE dbo.Update_tblTotal_tInventory_tGood_For_FinalPrice
(  
 @SystemDate   NVARCHAR(50),  
 @SystemDay    NVARCHAR(50),  
 @SystemTime   NVARCHAR(50),   
 @DateBefore   NVARCHAR(50),  
 @DateAfter    NVARCHAR(50),  
 @Type  int  ,  
 @InventoryNo Int ,  
 @AccountYear Smallint ,
 @ZeroNegative BIT 
)   
  
AS  
BEGIN TRAN  

	UPDATE dbo.tInventory_Good
	SET BuyPriceAverage = FirstPrice , MojodiPrice = FirstPrice
	WHERE InventoryNo = @InventoryNo AND AccountYear = @AccountYear
	
	UPDATE dbo.tInventory_Good
	SET BuyPriceAverage = T.AverageBuyFee , MojodiPrice = T.AverageBuyFee  FROM (
	Select (IsNull(Sum(FeeUnit * Amount) ,0) + ISNULL(FirstPrice * FirstMojodi , 0)) /(ISNULL(Sum(Amount),1) + ISNULL(FirstMojodi ,1)) AS AverageBuyFee , dbo.tInventory_Good.GoodCode , tInventory_Good.AccountYear , tInventory_Good.InventoryNo
--	Select IsNull(Sum(FeeUnit * Amount) ,0) /(ISNULL(Sum(Amount),1) ) AS AverageBuyFee , dbo.tInventory_Good.GoodCode , tInventory_Good.AccountYear , tInventory_Good.InventoryNo
	From tFacM inner join tfacd On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch AND dbo.tFacD.intInventoryNo = @InventoryNo
	INNER JOIN dbo.tInventory_Good ON tfacd.GoodCode = dbo.tInventory_Good.GoodCode 
		AND dbo.tInventory_Good.AccountYear = @AccountYear AND dbo.tInventory_Good.Branch = dbo.tFacD.Branch 
		AND dbo.tInventory_Good.InventoryNo = @InventoryNo
	Where tfacm.Status = 1 and Recursive = 0 And tfacm.AccountYear = @AccountYear AND tfacD.intInventoryNo = @InventoryNo 
	GROUP BY tfacd.GoodCode ,tInventory_Good.GoodCode, tInventory_Good.AccountYear , tInventory_Good.InventoryNo ,  tInventory_Good.FirstMojodi  ,  tInventory_Good.FirstPrice)T
	WHERE tInventory_Good.AccountYear = t.AccountYear  AND dbo.tInventory_Good.InventoryNo = t.InventoryNo AND tInventory_Good.GoodCode = t.GoodCode

UPDATE  tInventory_Good  
    
 Set    BuyAmount = T2.BuyAmount,  
		SaleAmount = T2.SaleAmount ,  
		LossAmount = T2.LossAmount ,
		BuyReturnAmount = T2.BuyReturnAmount ,  
		SaleReturnAmount = T2.SaleReturnAmount ,  
		FromStoreAmount = T2.FromStoreAmount ,  
		toStoreAmount = T2.toStoreAmount ,  
		Mojodi = T2.Mojodi , 
 	    MojodiPrice = CASE tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount  - tInventory_Good.FromStoreAmount + tInventory_Good.toStoreAmount  WHEN 0 THEN 0 ELSE ( firstMojodiRial + BuyRial - FromStoreRial + toStoreRial ) / (tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount  - tInventory_Good.FromStoreAmount + tInventory_Good.toStoreAmount) END 
   
 FROM dbo.tblTotal_tInventory_tGood_For_FinalPrice  
  (  
   @DateBefore   ,  
   @DateAfter    ,  
   @Type    ,  
   @InventoryNo  ,  
   @AccountYear  
  )  
   AS T2    
     Where tInventory_Good.GoodCode = T2.GoodCode And tInventory_Good.InventoryNo = T2.InventoryNo and tInventory_Good.AccountYear = @AccountYear  
	if @@Error <> 0   
	 goto ErrHandler  
  

 	  
UPDATE  tInventory_Good  
	SET MojodiPrice = 0 WHERE Mojodi = 0

IF @ZeroNegative = 1
UPDATE  tInventory_Good  
	SET MojodiPrice = 0 WHERE MojodiPrice < 0

Commit Tran   
  
return  
  
ErrHandler:  
RollBack Tran  
return 
  




GO


