-- ÈÑØÑÝ ˜ÑÏä ãÔ˜á ÒíÑ
-- ÏÑ ÕæÑÊí ˜å ÓØÑí ÏÑ ÌÏæá ãæÌæÏí ÇäÈÇÑ ÇÒ ˜ÇáÇ äÈÇÔÏ äãí ÊæÇäÓÊ ÈÑ ÇÓÇÓ ÝÇ˜ÊæÑ åÇ ÏÇÏå ãæÌæÏí ÑÇ ÈíÇæÑÏ



ALTER  FUNCTION [dbo].[tblTotal_tInventory_tGood_For_Mojodi]
    (
      @intLanguage INT ,
      @SystemDate NVARCHAR(50) ,
      @SystemDay NVARCHAR(50) ,
      @SystemTime NVARCHAR(50) ,
      @DateBefore NVARCHAR(50) ,
      @DateAfter NVARCHAR(50) ,
      @Type INT ,
      @InventoryNo1 INT ,
      @InventoryNo2 INT ,
      @Branch INT ,
      @UsePercentFlag INT ,
      @AccountYear SMALLINT
    )
RETURNS @ReturnTable TABLE
    (
      DateBefore NVARCHAR(50) ,
      DateAfter NVARCHAR(50) ,
      Sysdate NVARCHAR(50) ,
      [Name] NVARCHAR(50) ,
      goodtype INT ,
      GoodCode INT ,
      Branch INT ,
      intInventoryNo INT ,
      firstMojodi FLOAT ,
      BuyAmount FLOAT ,
      SaleAmount FLOAT ,
      LossAmount FLOAT ,
      BuyReturnAmount FLOAT ,
      SaleReturnAmount FLOAT ,
      FromStoreAmount FLOAT ,
      toStoreAmount FLOAT ,
      Mojodi FLOAT
    )
AS
    BEGIN
-------------------------------------------------------------------------------------------------------------------
        DECLARE @TimeTitle NVARCHAR(10)
        IF @intLanguage = 0
            SET @TimeTitle = N' ÓÇÚÊ : '
        ELSE
            SET @TimeTitle = N'Time: '
        INSERT  INTO @ReturnTable
                ( DateBefore ,
                  DateAfter ,
                  Sysdate ,
                  [Name] ,
                  goodtype ,
                  GoodCode ,
                  Branch ,
                  intInventoryNo ,
                  firstMojodi ,
                  BuyAmount ,
                  SaleAmount ,
                  LossAmount ,
                  BuyReturnAmount ,
                  SaleReturnAmount ,
                  FromStoreAmount ,
                  toStoreAmount ,
                  Mojodi  
                )
                SELECT  DateBefore ,
                        DateAfter ,
                        Sysdate ,
                        GoodFirstName ,
                        @Type AS goodtype ,
                        GoodCode ,
                        Branch ,
                        InventoryNo ,
                        firstMojodi ,
                        BuyAmount ,
                        SaleAmount ,
                        LossAmount ,
                        BuyReturnAmount ,
                        SaleReturnAmount ,
                        FromStoreAmount ,
                        toStoreAmount ,
                        Mojodi
                FROM    ( SELECT    Y.* ,
                                    CASE @intLanguage
                                      WHEN 0 THEN tinventory.Description
                                      WHEN 1 THEN tinventory.LatinDescription
                                    END AS InventoryName
                          FROM      ( SELECT    ISNULL(tGood.Name,
                                                       W.GoodFirstName) AS GoodFirstName ,
                                                tUnitGood.[Description] ,
                                                ISNULL(tInventory_Good.GoodCode,
                                                       W.GoodFirstCode) AS GoodCode ,
                                                ISNULL(tInventory_Good.Branch,
                                                       W.Branch) AS Branch ,
                                                ISNULL(tInventory_Good.InventoryNo,
                                                       W.intInventoryNo) AS InventoryNo ,
                                                ISNULL(FirstMojodi, 0) AS firstMojodi ,
                                                ISNULL(MAX(W.BuyAmount), 0) AS BuyAmount ,
                                                ISNULL(SUM(W.SaleAmount), 0) AS SaleAmount ,
                                                ISNULL(MAX(W.LossAmount), 0) AS LossAmount ,
                                                ISNULL(MAX(W.BuyReturnAmount),
                                                       0) AS BuyReturnAmount ,
                                                ISNULL(MAX(W.SaleReturnAmount),
                                                       0) AS SaleReturnAmount ,
                                                ISNULL(MAX(W.FromStoreAmount),
                                                       0) AS FromStoreAmount ,
                                                ISNULL(MAX(W.toStoreAmount), 0) AS toStoreAmount ,
                                                ISNULL(FirstMojodi, 0)
                                                + ISNULL(MAX(W.BuyAmount), 0)
                                                - ISNULL(SUM(W.SaleAmount), 0)
                                                - ISNULL(MAX(W.LossAmount), 0)
                                                - ISNULL(MAX(W.BuyReturnAmount),
                                                         0)
                                                + ISNULL(MAX(W.SaleReturnAmount),
                                                         0)
                                                - ISNULL(MAX(W.FromStoreAmount),
                                                         0)
                                                + ISNULL(MAX(W.toStoreAmount),
                                                         0) AS Mojodi ,
                                                @DateBefore AS DateBefore ,
                                                @DateAfter AS DateAfter ,
                                                @SystemDay + ' ' + @SystemDate
                                                + ' ' + @TimeTitle
                                                + @SystemTime AS Sysdate
                                      FROM      tInventory_Good
                                                FULL OUTER   JOIN ( SELECT
                                                              MAX(ISNULL(X.Branch,
                                                              FirstGoods.branch)) AS branch ,
                                                              MAX(ISNULL(X.intInventoryNo,
                                                              FirstGoods.intInventoryNo)) AS intInventoryNo ,
                                                              ISNULL(FirstGoods.GoodFirstName,
                                                              X.GoodFirstName) AS GoodFirstName ,
                                                              ISNULL(X.GoodfirstCode,
                                                              FirstGoods.GoodFirst1) AS GoodfirstCode ,
                                                              MAX(ISNULL(FirstGoods.SaleAmount,
                                                              0))
                                                              + MAX(ISNULL(X.SaleAmount,
                                                              0)) AS saleamount ,
                                                              MAX(ISNULL(FirstGoods.SaleReturnAmount,
                                                              0))
                                                              + MAX(ISNULL(X.SaleReturnAmount,
                                                              0)) AS SaleReturnAmount ,
                                                              MAX(ISNULL(FirstGoods.Buy,
                                                              0)) AS BuyAmount ,
                                                              MAX(ISNULL(FirstGoods.Losses,
                                                              0)) AS LossAmount ,
                                                              MAX(ISNULL(FirstGoods.BuyReturn,
                                                              0)) AS BuyReturnAmount ,
                                                              MAX(ISNULL(FirstGoods.FromStore,
                                                              0)) AS FromStoreAmount ,
                                                              MAX(ISNULL(FirstGoods.ToStore,
                                                              0)) AS ToStoreAmount
                                                              FROM
                                                              ( SELECT
                                                              Branch ,
                                                              intInventoryNo ,
                                                              GoodfirstCode ,
                                                              SUM(SaleAmount) AS SaleAmount ,
                                                              SUM(SaleReturnAmount) AS SaleReturnAmount ,
                                                              GoodFirstName
                                                              FROM
                                                              ( SELECT
                                                              Branch ,
                                                              intInventoryNo ,
                                                              ISNULL(GoodfirstCode,
                                                              GoodCode) AS GoodfirstCode ,
                                                              ( ( SaleAmount
                                                              * fltUsedvalue )
                                                              + ( [SaleAmount]
                                                              * Pert ) ) AS SaleAmount ,
                                                              SaleReturnAmount ,
                                                              CASE @intLanguage
                                                              WHEN 0
                                                              THEN tGood.[Name]
                                                              WHEN 1
                                                              THEN tGood.LatinName
                                                              END AS GoodFirstName
                                                              FROM
                                                              ( SELECT
                                                              @Branch AS Branch , --tFacd.Branch ,
                                                              intInventoryNo ,
                                                              Goodcode ,
                                                              tGood.[Name] AS MainName ,
                                                              tfacd.Serveplace ,
                                                              CASE Status
                                                              WHEN 2
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleAmount ,
                                                              CASE Status
                                                              WHEN 5
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleReturnAmount
                                                              FROM
                                                              dbo.tFacM
                                                              INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                                              AND dbo.tFacM.Branch = dbo.tFacD.Branch
                                                              INNER JOIN tGood ON tGood.Code = tFacd.GoodCode
                                                              WHERE
                                                              dbo.tFacM.Recursive = 0
                                                              AND dbo.tFacM.AccountYear = @AccountYear
                                                              AND tFacM.[Date] >= @DateBefore
                                                              AND tFacM.[Date] <= @DateAfter
                                                              AND ( dbo.tFacD.intInventoryNo >= @InventoryNo1
                                                              AND dbo.tFacD.intInventoryNo <= @InventoryNo2
                                                              --AND tFacd.Branch = @Branch
                                                              )
					--AND dbo.tFacM.ShiftNo = dbo.Get_Current_Shift(@SystemTime) --Just only for Mashad(Malek Restaurant)
                                                              GROUP BY Goodcode ,
                                                              Name ,
                                                              tfacd.serveplace ,
                                                              tfacm.Status ,
                                                              intInventoryNo --,
                                                              --tFacd.Branch
                                                              ) T
                                                              LEFT OUTER JOIN usepercent ON T.GoodCode = usepercent.code
                                                              AND T.serveplace = usepercent.intserveplace
                                                              INNER JOIN tGood ON tGood.Code = usepercent.GoodFirstCode
                                                              ) AS U
                                                              GROUP BY GoodFirstCode ,
                                                              intInventoryNo ,
                                                              Branch ,
                                                              GoodFirstName --, ServePlace
                                                              ) X
                                                              FULL OUTER  JOIN ( SELECT
                                                              F.Branch ,
                                                              F.intInventoryNo ,
                                                              F.GoodFirst1 ,
                                                              F.GoodFirstName ,
                                                              SUM(Buy) AS Buy ,
                                                              SUM(SaleAmount) AS SaleAmount ,
                                                              SUM(Losses) AS Losses ,
                                                              SUM(BuyReturn) AS BuyReturn ,
                                                              SUM(SaleReturnAmount) AS SaleReturnAmount ,
                                                              SUM(FromStore) AS FromStore ,
                                                              SUM(ToStore) AS ToStore
                                                              FROM
                                                              ( SELECT
                                                              @Branch AS Branch ,--tFacd.Branch ,
                                                              intInventoryNo ,
                                                              GoodCode AS GoodFirst1 ,
                                                              tgood.NAME AS GoodFirstName ,
                                                              CASE Status
                                                              WHEN 1
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS Buy ,
                                                              CASE Status
                                                              WHEN 2
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleAmount ,
                                                              CASE Status
                                                              WHEN 3
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS Losses ,
                                                              CASE Status
                                                              WHEN 4
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS BuyReturn ,
                                                              CASE Status
                                                              WHEN 5
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleReturnAmount ,
                                                              CASE Status
                                                              WHEN 6
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS FromStore ,
                                                              CASE Status
                                                              WHEN 7
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS ToStore
                                                              FROM
                                                              dbo.tFacM
                                                              INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                                              AND dbo.tFacM.Branch = dbo.tFacD.Branch
                                                              INNER JOIN tGood ON tGood.Code = tFacd.GoodCode
                                                              WHERE
                                                              dbo.tFacM.Recursive = 0
                                                              AND dbo.tFacM.AccountYear = @AccountYear
                                                              AND tFacM.[Date] >= @DateBefore
                                                              AND tFacM.[Date] <= @DateAfter
                                                              AND ( dbo.tFacD.intInventoryNo >= @InventoryNo1
                                                              AND dbo.tFacD.intInventoryNo <= @InventoryNo2
                                                              --AND tFacd.Branch = @Branch
                                                              )
			--AND dbo.tFacM.ShiftNo = dbo.Get_Current_Shift(@SystemTime) --Just only for Mashad(Malek Restaurant)
                                                              GROUP BY Goodcode ,
                                                              tfacm.Status ,
                                                              intInventoryNo ,
                                                            --  tFacd.Branch ,
                                                              tGood.[name]
                                                              ) F
                                                              GROUP BY F.GoodFirst1 ,
                                                              F.intInventoryNo ,
                                                              F.Branch ,
                                                              GoodFirstName
                                                              ) FirstGoods ON FirstGoods.GoodFirst1 = X.GoodfirstCode
                                                              GROUP BY GoodFirstCode ,
                                                              FirstGoods.intInventoryNo ,
                                                              FirstGoods.Branch ,
                                                              FirstGoods.GoodFirstName ,
                                                              X.GoodFirstName ,
                                                              GoodFirst1
                                                              ) W ON tInventory_Good.GoodCode = W.GoodfirstCode
                                                              AND tInventory_Good.Branch = W.Branch
                                                              AND tInventory_Good.InventoryNo = W.intInventoryNo
                                                INNER JOIN dbo.tGood ON dbo.tGood.Code = ISNULL(tInventory_Good.GoodCode,
                                                              W.GoodfirstCode)
                                                              AND ISNULL(tInventory_Good.InventoryNo,
                                                              W.intInventoryNo) >= @InventoryNo1
                                                              AND ISNULL(tInventory_Good.InventoryNo,
                                                              W.intInventoryNo) <= @InventoryNo2
                                                              AND ISNULL(tInventory_Good.Branch,
                                                              W.Branch) = @Branch
                                                              AND ISNULL(dbo.tInventory_Good.AccountYear,
                                                              @AccountYear) = @AccountYear
                                                              AND tGood.GoodType = @Type
                                                INNER JOIN tUnitGood ON tGood.Unit = tUnitGood.Code
                                      WHERE     GoodType = @Type
                                      GROUP BY  tInventory_Good.GoodCode ,
                                                GoodFirstCode ,
                                                tGood.Name ,
                                                W.GoodFirstName ,
                                                FirstMojodi ,
                                                W.intInventoryNo ,
                                                tUnitGood.[Description] ,
                                                W.Branch ,
                                                tInventory_Good.InventoryNo ,
                                                tInventory_Good.Branch
                                    ) y
                                    INNER JOIN tInventory ON tInventory.InventoryNo = Y.InventoryNo
                                                             AND tInventory.Branch = Y.Branch
                        ) AS T

-----------------------------------------------------------------------------------------
--END
        RETURN


    END
--===============================================
