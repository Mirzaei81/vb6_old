
--ScriptV26_16_Fix16_کالای واسطه با ساعت.SQL

--95/03/20

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTotal_tInventory_tGood_For_Mojodi_Vaseteh]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[tblTotal_tInventory_tGood_For_Mojodi_Vaseteh]
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE FUNCTION [dbo].[tblTotal_tInventory_tGood_For_Mojodi_Vaseteh]
    (
      @intLanguage INT ,
      @SystemDate NVARCHAR(50) ,
      @SystemDay NVARCHAR(50) ,
      @SystemTime NVARCHAR(50) ,
      @DateBefore NVARCHAR(10) ,
      @DateAfter NVARCHAR(10) ,
      @TimeBefore NVARCHAR(5) ,
      @TimeAfter NVARCHAR(5) ,
      @Type INT ,
      @InventoryNo1 INT ,
      @InventoryNo2 INT ,
      @Branch INT ,
      @UsePercentFlag INT ,
      @AccountYear SMALLINT
    )
RETURNS @ReturnTable TABLE
    (
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
        DECLARE @TimeTitle NVARCHAR(10)
        IF @intLanguage = 0 
            SET @TimeTitle = N' ساعت : '
        ELSE 
            SET @TimeTitle = N'Time: '
        INSERT  INTO @ReturnTable
                ( Sysdate ,
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
                SELECT  Sysdate ,
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
                                                INNER JOIN dbo.tGood ON dbo.tGood.Code = tInventory_Good.GoodCode
                                                              AND tInventory_Good.InventoryNo >= @InventoryNo1
                                                              AND tInventory_Good.InventoryNo <= @InventoryNo2
                                                              AND tInventory_Good.Branch = @Branch
                                                              AND dbo.tInventory_Good.AccountYear = @AccountYear
                                                              AND tGood.GoodType = @Type
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
                                                              tFacd.Branch ,
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
                                                              AND (tFacM.[Date] >= @DateBefore AND tfacM.[Date] <= @DateAfter)
                                                              AND (tfacM.Time >= @TimeBefore or tFacM.[Date] > @DateBefore ) 
                                                              AND (tfacM.Time <= @TimeAfter or tFacM.[Date] < @DateAfter)
                                                              AND ( dbo.tFacD.intInventoryNo >= @InventoryNo1
                                                              AND dbo.tFacD.intInventoryNo <= @InventoryNo2
                                                              AND tFacd.Branch = @Branch
                                                              )
					--AND dbo.tFacM.ShiftNo = dbo.Get_Current_Shift(@SystemTime) --Just only for Mashad(Malek Restaurant)
                                                              GROUP BY Goodcode ,
                                                              Name ,
                                                              tfacd.serveplace ,
                                                              tfacm.Status ,
                                                              intInventoryNo ,
                                                              tFacd.Branch
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
                                                              tFacd.Branch ,
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
                                                              AND (tFacM.[Date] >= @DateBefore AND tfacM.[Date] <= @DateAfter)
                                                              AND (tfacM.Time >= @TimeBefore or tFacM.[Date] > @DateBefore ) 
                                                              AND (tfacM.Time <= @TimeAfter or tFacM.[Date] < @DateAfter)
                                                              AND ( dbo.tFacD.intInventoryNo >= @InventoryNo1
                                                              AND dbo.tFacD.intInventoryNo <= @InventoryNo2
                                                              AND tFacd.Branch = @Branch
                                                              )
			--AND dbo.tFacM.ShiftNo = dbo.Get_Current_Shift(@SystemTime) --Just only for Mashad(Malek Restaurant)
                                                              GROUP BY Goodcode ,
                                                              tfacm.Status ,
                                                              intInventoryNo ,
                                                              tFacd.Branch ,
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
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tblTotal_tInventory_tGood_For_Mojodi_Vaseteh]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_tblTotal_tInventory_tGood_For_Mojodi_Vaseteh]
GO




CREATE   PROCEDURE [dbo].[Update_tblTotal_tInventory_tGood_For_Mojodi_Vaseteh]
    (
      @intLanguage INT,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @DateBefore NVARCHAR(50),
      @DateAfter NVARCHAR(50),
      @TimeBefore NVARCHAR(5) ,
      @TimeAfter NVARCHAR(5) ,
      @Type INT,
      @InventoryNo1 INT,
      @InventoryNo2 INT,
      @Branch INT,
      @UsePercentFlag INT,
      @AccountYear SMALLINT
    )
AS 
    BEGIN TRAN

    SET @SystemTime = dbo.SetTimeFormat(GETDATE())

    INSERT  INTO tInventory_Good
            (
              Branch,
              InventoryNo,
              GoodCode,
              BuyAmount,
              SaleAmount,
              LossAmount,
              BuyReturnAmount,
              SaleReturnAmount,
              FromStoreAmount,
              toStoreAmount,
              Mojodi,
              AccountYear 
            )
            SELECT  T1.Branch,
                    T1.intInventoryNo,
                    T1.GoodCode,
                    T1.BuyAmount,
                    T1.SaleAmount,
                    T1.LossAmount,
                    T1.BuyReturnAmount,
                    T1.SaleReturnAmount,
                    T1.FromStoreAmount,
                    T1.toStoreAmount,
--                    T1.Mojodi,
                    CASE dbo.AutoHavale() WHEN 0 THEN  CASE WHEN @Type IN ( 1, 3 )
					                         THEN T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount
					                              - T1.FromStoreAmount - T1.LossAmount
					                              + T1.ToStoreAmount
					                         ELSE T1.Mojodi END 
                    			 ELSE  CASE WHEN @Type IN ( 1, 3 )
					                         THEN T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount
					                              - T1.FromStoreAmount - T1.LossAmount
					                              + T1.ToStoreAmount - T1.Saleamount
					                         ELSE T1.Mojodi END 
		    END  AS Mojodi,
                    @AccountYear
            FROM    dbo.tblTotal_tInventory_tGood_For_Mojodi_Vaseteh(@intLanguage,
                                                             @SystemDate,
                                                             @SystemDay,
                                                             @SystemTime,
                                                             @DateBefore,
                                                             @DateAfter,
                                                             @TimeBefore ,
                                                             @TimeAfter ,
                                                             @Type,
                                                             @InventoryNo1,
                                                             @InventoryNo2,
                                                             @Branch,
                                                             @UsePercentFlag,
                                                             @AccountYear) AS T1

--SELECT T1.GoodCode,T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount - T1.FromStoreAmount - T1.LossAmount+ T1.ToStoreAmount FROM  dbo.tblTotal_tInventory_tGood_For_Mojodi(0,'','','',N'89/01/01',N'89/02/23',1,1,1,1,0,1389) T1
                                                          




            WHERE   0 = ( SELECT    COUNT(GoodCode)
                          FROM      tInventory_Good
                          WHERE     GoodCode = T1.GoodCode
                                    AND InventoryNo = T1.intInventoryNo
                                    AND Branch = T1.Branch
                                    AND AccountYear = @AccountYear
                        )
		

    IF @@Error <> 0 
        GOTO ErrHandler
------------------------------------------------------------------------

    UPDATE  tInventory_Good
    SET     BuyAmount = T2.BuyAmount,
            SaleAmount = T2.SaleAmount,
            LossAmount = T2.LossAmount,
            BuyReturnAmount = T2.BuyReturnAmount,
            SaleReturnAmount = T2.SaleReturnAmount,
            FromStoreAmount = T2.FromStoreAmount,
            toStoreAmount = T2.toStoreAmount,
            Mojodi = CASE dbo.AutoHavale() WHEN 0 THEN  CASE WHEN @Type IN ( 1, 3 )
			                         THEN T2.firstMojodi+T2.BuyAmount - T2.BuyReturnAmount
			                              - T2.FromStoreAmount - T2.LossAmount
			                              + T2.ToStoreAmount
			                         ELSE T2.Mojodi END 
	    			 ELSE  CASE WHEN @Type IN ( 1, 3 )
				                         THEN T2.firstMojodi+T2.BuyAmount - T2.BuyReturnAmount
				                              - T2.FromStoreAmount - T2.LossAmount
				                              + T2.ToStoreAmount - T2.Saleamount
				                         ELSE T2.Mojodi END 
		    END  

    FROM    dbo.tblTotal_tInventory_tGood_For_Mojodi_Vaseteh(@intLanguage, @SystemDate,
                                                     @SystemDay, @SystemTime,
                                                     @DateBefore, @DateAfter,@TimeBefore , @TimeAfter ,
                                                     @Type, @InventoryNo1,
                                                     @InventoryNo2, @Branch,
                                                     @UsePercentFlag,
                                                     @AccountYear) AS T2
    WHERE   tInventory_Good.GoodCode = T2.GoodCode
            AND tInventory_Good.InventoryNo = T2.intInventoryNo
            AND tInventory_Good.Branch = T2.Branch
            AND tInventory_Good.AccountYear = @AccountYear
---------------------------------------------------------------------
    IF @@Error <> 0 
        GOTO ErrHandler

    UPDATE  tInventory_Good
    SET     Mojodi = 0
    FROM    ( SELECT    
                    GoodCode ,
                    InventoryNo ,
                    Branch
          FROM      tInventory_Good INNER JOIN dbo.tGood ON dbo.tInventory_Good.GoodCode = dbo.tGood.Code AND GoodType = 4
          WHERE     tInventory_Good.AccountYear = @AccountYear
                    AND tInventory_Good.Branch = @Branch
                                ) T3
    WHERE   tInventory_Good.GoodCode = T3.GoodCode
            AND tInventory_Good.InventoryNo = T3.InventoryNo
            AND tInventory_Good.Branch = T3.Branch
            AND tInventory_Good.AccountYear = @AccountYear
	    AND tInventory_Good.Mojodi < 0
---------------------------------------------------------------------
    IF @@Error <> 0 
        GOTO ErrHandler

    COMMIT TRAN 

    RETURN 1

    ErrHandler:
    ROLLBACK TRAN
    RETURN -1
	

GO


