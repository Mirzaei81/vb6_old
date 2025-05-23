SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

-------------------------------------*******************************************
---------------------------------------------------اصلاح گزارش مواد مصرفی********


ALTER PROC Rep_UseOfGood
    (
      @SystemDate NVARCHAR(20),
      @SystemDay NVARCHAR(20),
      @SystemTime NVARCHAR(20),
      @Date1 NVARCHAR(8),
      @Date2 NVARCHAR(8),
      @Status1 INT,
      @InventoryNo1 INT,
      @Branch1 INT
							
    )
AS 

    BEGIN
        SELECT  SUM(T.Amount) AS Amount,
                T.[Name],
             --AS FeeUnit,
                T.GoodCode,
                ( SELECT    dbo.tUnitGood.Description
                  FROM      dbo.tUnitGood
                  WHERE     Code = T.Unit
                ) AS Unit,
                T.Discount,
                T.Weight,
                T.NumberOfUnit,
               -- Rate,
                ( SELECT    dbo.tInventory.Description
                  FROM      dbo.tInventory
                  WHERE     dbo.tInventory.InventoryNo = intInventoryNo
                ) AS InventoryName,
                SellPrice,
                BuyPrice,
                @SystemDate AS SystemDate,
                @SystemDay AS SystemDay,
                @SystemTime AS SystemTime 
				FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                            --dbo.tFacD.Rate,
                            dbo.tFacD.intInventoryNo,
                            dbo.tFacD.DestInventoryNo,
                            ( dbo.tFacD.Amount * dbo.tUsePercent.fltUsedValue ) AS Amount,
                            ( SELECT    Name
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [Name],
                            dbo.tGood.Weight,
                            ( SELECT    Unit
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS Unit,
                            dbo.tGood.NumberOfUnit,
                            dbo.tFacD.Discount,
                            ( SELECT    SellPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [SellPrice],
                            ( SELECT    BuyPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [BuyPrice]
						FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                    AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                            JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
						WHERE     tfacm.Branch = @Branch1
                            AND tfacm.Recursive = 0
                            AND tfacm.status = @Status1
                            AND tfacm.Date >= @Date1
                            AND tfacm.Date <= @Date2
                            AND tfacd.intInventoryNo = @InventoryNo1
                  UNION ALL
                  SELECT    dbo.tFacD.GoodCode,
                            --dbo.tFacD.Rate,
                            dbo.tFacD.intInventoryNo,
                            dbo.tFacD.DestInventoryNo,
                            dbo.tFacD.Amount,
                            dbo.tGood.Name,
                            dbo.tGood.Weight,
                            dbo.tGood.Unit,
                            dbo.tGood.NumberOfUnit,
                            dbo.tFacD.Discount,
                            dbo.tGood.SellPrice,
                            dbo.tGood.BuyPrice
                  FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                  WHERE     tfacm.Branch = @Branch1
                            AND tfacm.Recursive = 0
                            AND tfacm.status = @Status1
                            AND tfacm.Date >= @Date1
                            AND tfacm.Date <= @Date2
                            AND tfacd.intInventoryNo = @InventoryNo1
                            --AND dbo.tFacD.GoodCode NOT IN (
                            --SELECT  dbo.tUsePercent.GoodCode
                            --FROM    dbo.tUsePercent
                            --WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                            AND dbo.tGood.GoodType = 3
                ) T
        GROUP BY T.GoodCode,
                T.Name,
                T.Weight,
                T.Unit,
                --Rate,
                intInventoryNo,
                DestInventoryNo,
                T.Discount,
                T.NumberOfUnit,
                T.SellPrice,
                T.BuyPrice
	
    END


GO
