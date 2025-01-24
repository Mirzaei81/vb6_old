
--ScriptV26_14_Fix52_GetOrderGoodAmount.sql
--For All Version V26_14 & V26_15 & V26_16

ALTER     PROCEDURE [dbo].[GetOrderGoodAmountInfo]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @level11 INT,
      @level12 INT,
      @level21 INT,
      @level22 INT,
      @Inventory1 INT,
      @AccountYear1  INT  
    )

AS 
    DECLARE @intLanguage INT 
    SET @intLanguage = 0
    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage = 0 
        SET @TimeTitle = N' ”«⁄  : '
    ELSE 
        SET @TimeTitle = N'Time: '


    SELECT  [CompDes],
            [Code],
            [vw_Good].[Level1],
            [vw_Good].[Level2],
            CASE @intLanguage
              WHEN 0 THEN [Name]
              ELSE [LatinName]
            END AS [Name],
            CASE @intLanguage
              WHEN 0 THEN [NamePrn]
              ELSE [LatinNamePrn]
            END AS [NamePrn],
            CASE @intLanguage
              WHEN 0 THEN [Level1Description]
              ELSE [Level1LatinDescription]
            END AS [Level1Name],
            CASE @intLanguage
              WHEN 0 THEN [Level2Description]
              ELSE [Level2LatinDescription]
            END AS [Level2Name]	,	 
            [OrderPoint],
            [MinValue],
            [MaxValue],
            [ProductCompany],
            [UnitDescription],
            [TypeDescription],
            @SystemDay + ' ' + @SystemDate  AS Sysdate,
            tInventory_Good.[InventoryNo],
            tInventory.Description AS Inventoryname ,
            Mojodi
			, ISNULL((SELECT TOP 1 feeunit from tfacd 
				inner join (SELECT MAX(tfacd.intSerialNo) AS maxint,goodcode
							FROM tfacd  INNER JOIN  [tFacM] ON [tFacD].[intSerialNo] = [tFacM].[intSerialNo]    
							where tfacm.status=1 and tfacm.accountyear=@AccountYear1 and tfacd.intInventoryNo = @Inventory1
					        GROUP BY goodcode )k ON tfacd.goodcode=k.goodcode and tfacd.intserialno=k.maxint
				WHERE [tfacd].[GoodCode]=[vw_Good].[Code] AND tfacd.intInventoryNo = @Inventory1),[vw_Good].BuyPrice) AS feeunit

    FROM    [dbo].[vw_Good]
            INNER JOIN tInventory_Good ON vw_Good.Code = tInventory_Good.GoodCode 
					AND tInventory_Good.InventoryNo = 1 AND tInventory_Good.AccountYear = @AccountYear1
            INNER JOIN tInventory ON tInventory.Branch = tInventory_Good.Branch
                                     AND tInventory.InventoryNo = @Inventory1

    WHERE   [vw_Good].[Level1] >= @level11
            AND [vw_Good].[Level1] <= @level12
            AND [vw_Good].[Level2] >= @level21
            AND [Level2] <= @level22
            AND [GoodType] <> 2
            AND [GoodType] <> 4
            AND tInventory_Good.[InventoryNo] = @Inventory1
            AND tInventory.Branch = tInventory_Good.Branch
            AND tInventory.InventoryNo = tInventory_Good.InventoryNo
            AND tInventory_Good.AccountYear = @AccountYear1
  		    AND [tInventory_Good].mojodi <= OrderPoint 
  		    AND OrderPoint > 0

ORDER BY Name

GO

--exec dbo.GetOrderGoodAmountInfo;1 N'94/10/10',N'Å‰Ã ‘‰»Â',N'12:30',11,22,1101,2201,1,1394
--GO
