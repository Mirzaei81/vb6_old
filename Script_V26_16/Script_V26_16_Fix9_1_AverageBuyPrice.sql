


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTotal_tInventory_tGood_For_FinalPrice]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].tblTotal_tInventory_tGood_For_FinalPrice
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE Function dbo.tblTotal_tInventory_tGood_For_FinalPrice    
(    
 @DateBefore   NVARCHAR(8),    
 @DateAfter    NVARCHAR(8),    
 @Type  int  ,    
 @InventoryNo Int ,    
 @AccountYear Smallint    
)     
RETURNS  @ReturnTable table    
   (    
   DateBefore  nvarchar(8)    
   ,DateAfter  nvarchar(8)    
   ,GoodName nvarchar(50)    
   ,goodtype  int    
   ,GoodCode  int    
   ,InventoryNo  int    
   ,firstMojodi  FLOAT    
   ,BuyAmount  FLOAT    
   ,LossAmount  FLOAT    
   ,BuyReturnAmount  FLOAT    
   ,FromStoreAmount  FLOAT    
   ,toStoreAmount  FLOAT    
   ,Mojodi  FLOAT     
   ,FirstMojodiRial  BIGINT    
   ,BuyRial  BIGINT    
   ,LossRial  BIGINT    
   ,BuyReturnRial  BIGINT    
   ,FromStoreRial  BIGINT    
   ,toStoreRial  BIGINT   
   ,SaleAmount FLOAT 
   ,SaleRial BIGINT 
   ,SaleReturnAmount FLOAT 
   )    
   
     
As    
BEGIN
	

 INSERT INTO @ReturnTable(DateBefore  ,DateAfter  ,GoodName , goodtype  ,GoodCode  ,InventoryNo  ,firstMojodi  ,BuyAmount  ,LossAmount  ,BuyReturnAmount  ,FromStoreAmount  ,toStoreAmount  ,Mojodi ,
 	BuyRial , LossRial , BuyReturnRial , FromStoreRial , ToStoreRial , FirstMojodiRial , SaleAmount , SaleRial , SaleReturnAmount )    
  Select                  @DateBefore  ,@DateAfter  , GoodName,    @Type ,GoodCode ,@InventoryNo  ,firstMojodi  ,BuyAmount  ,LossesAmount  ,BuyReturnAmount  , FromStoreAmount  ,toStoreAmount  ,MojodiAmount ,      
 	BuyRial , LossesRial , BuyReturnRial , FromStoreRial , ToStoreRial , FirstMojodiRial  , SaleAmount  , SaleRial ,SaleReturnAmount
  FROM (
SELECT  @InventoryNo AS InventoryNo , ISNULL(W.GoodCode ,  tInventory_Good.GoodCode) AS GoodCode  , ISNULL(W.GoodName , dbo.tGood.[Name]) AS GoodName,
	ISNULL(W.BuyAmount , 0) AS BuyAmount  , ISNULL(w.LossesAmount ,0) AS LossesAmount ,
	ISNULL(w.BuyReturnAmount , 0) AS BuyReturnAmount ,ISNULL(w.SaleReturnAmount , 0) AS SaleReturnAmount , 
	ISNULL(w.FromStoreAmount , 0) AS FromStoreAmount , ISNULL(w.ToStoreAmount , 0) AS ToStoreAmount , 
	ISNULL(W.BuyRial , 0) AS BuyRial  , ISNULL(W.SaleAmount ,0) AS SaleAmount , ISNULL(W.SaleRial ,0) AS SaleRial , ISNULL(w.LossesRial ,0) AS LossesRial ,
	ISNULL(w.BuyReturnRial , 0) AS BuyReturnRial , ISNULL(w.SaleReturnRial , 0) AS SaleReturnRial ,
	ISNULL(w.FromStoreRial , 0) AS FromStoreRial , ISNULL(w.ToStoreRial , 0) AS ToStoreRial , 
	dbo.tInventory_Good.FirstMojodi + ISNULL(W.BuyAmount , 0) - ISNULL(w.LossesAmount ,0) -
	ISNULL(w.BuyReturnAmount , 0) - ISNULL(w.FromStoreAmount , 0) + ISNULL(w.ToStoreAmount , 0) AS MojodiAmount , 
	(dbo.tInventory_Good.FirstMojodi * dbo.tInventory_Good.FirstPrice ) AS FirstMojodiRial , dbo.tInventory_Good.FirstMojodi
	
   From tInventory_Good     
    INNER join  dbo.tGood ON  dbo.tGood.Code = tInventory_Good.GoodCode and  tInventory_Good.InventoryNo = @InventoryNo    
    AND dbo.tInventory_Good.AccountYear = @AccountYear  And tGood.GoodType = @Type     
      
  Full Outer   join     
  (    
  select @InventoryNo AS InventoryNo,F.GoodCode,F.GoodName , sum(BuyAmount) as BuyAmount,sum(SaleAmount) as SaleAmount,
  sum(LossesAmount) as LossesAmount,sum(BuyReturnAmount)as BuyReturnAmount,sum(SaleReturnAmount)as SaleReturnAmount,sum(FromStoreAmount) as FromStoreAmount    
    ,sum(ToStoreAmount) as ToStoreAmount , sum(BuyRial) as BuyRial,sum(SaleRial) as SaleRial,
  sum(LossesRial) as LossesRial,sum(BuyReturnRial)as BuyReturnRial,sum(SaleReturnRial)as SaleReturnRial,sum(FromStoreRial) as FromStoreRial    
    ,sum(ToStoreRial) as ToStoreRial from    
  (

   SELECT   intInventoryNo ,GoodCode  ,tgood.NAME AS GoodName ,    
   Case Status When 1 Then Sum(Amount) Else 0 End As BuyAmount ,    
   Case Status When 2 Then Sum(Amount) Else 0 End As SaleAmount ,    
   Case Status When 3 Then Sum(Amount) Else 0 End As LossesAmount ,    
   Case Status When 4 Then Sum(Amount) Else 0 End As BuyReturnAmount ,    
   Case Status When 5 Then Sum(Amount) Else 0 End As SaleReturnAmount ,   
   Case Status When 6 Then Sum(Amount) Else 0 End As FromStoreAmount ,    
   Case Status When 7 Then Sum(Amount) Else 0 End As ToStoreAmount ,     

   Case Status When 1 Then Sum(Amount * (1 - (tFacd.Discount/100))* FeeUnit ) Else 0 End As BuyRial ,    
   Case Status When 2 Then Sum(Amount * (1 - (tFacd.Discount/100))* FeeUnit ) Else 0 End As SaleRial ,    
   Case Status When 3 Then Sum(Amount * FeeUnit ) Else 0 End As LossesRial ,    
   Case Status When 4 Then Sum(Amount * FeeUnit ) Else 0 End As BuyReturnRial ,    
   Case Status When 5 Then Sum(Amount * FeeUnit ) Else 0 End As SaleReturnRial ,   
   Case Status When 6 Then Sum(Amount * FeeUnit ) Else 0 End As FromStoreRial ,    
   Case Status When 7 Then Sum(Amount * FeeUnit ) Else 0 End As ToStoreRial     
   FROM  dbo.tFacM     
   INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND  dbo.tFacM.Branch = dbo.tFacD.Branch     
   inner join tGood On tGood.Code = tFacd.GoodCode     
   WHERE  dbo.tFacM.Recursive = 0  And dbo.tFacM.AccountYear = @AccountYear  AND tFacM.[Date] >=  @DateBefore  AND tFacM.[Date] <=  @DateAfter     
   And  ( dbo.tFacD.intInventoryNo = @InventoryNo ) AND dbo.tGood.GoodType  = @Type 
   Group By Goodcode , tfacm.Status ,intInventoryNo , tFacd.Branch , tGood.[name]   
      
   )F    
   Group By F.GoodCode , F.intInventoryNo , GoodName   

  )W    
          On tInventory_Good.GoodCode = W.GoodCode and tInventory_Good.InventoryNo = W.InventoryNo   
          
           --ORDER BY GoodCode
           )X
           
Return    
    
    
End    


GO



ALTER   PROCEDURE dbo.Update_tblTotal_tInventory_tGood_For_FinalPrice
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
 	    MojodiPrice = CASE tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount  WHEN 0 THEN 0 ELSE ( firstMojodiRial + BuyRial ) / (tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount) END 
   
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


--exec Update_tblTotal_tInventory_tGood_For_FinalPrice N'93/09/06', N'äÌ ÔäÈå', N'11:29', N'93/01/01', N'93/09/06', 3, 1, 1393, 1
--GO


