
--بروز رسانی کالاها در انبار
--94/01/22


ALTER  PROCEDURE [dbo].[Update_tblTotal_tInventory_tGood_For_Mojodi]
    (
      @intLanguage INT,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @DateBefore NVARCHAR(50),
      @DateAfter NVARCHAR(50),
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
                    CASE WHEN @Type = 1
								THEN T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount
	                              - T1.FromStoreAmount - T1.LossAmount
	                              + T1.ToStoreAmount  
	                     WHEN @Type = 3
								THEN T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount
	                              - T1.FromStoreAmount - T1.LossAmount
	                              + T1.ToStoreAmount

					    ELSE T1.Mojodi  
					END  AS Mojodi,
                    @AccountYear
				FROM    dbo.tblTotal_tInventory_tGood_For_Mojodi(@intLanguage,
                                                             @SystemDate,
                                                             @SystemDay,
                                                             @SystemTime,
                                                             @DateBefore,
                                                             @DateAfter, @Type,
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
            Mojodi = CASE WHEN @Type = 1
			                         THEN T2.firstMojodi+T2.BuyAmount - T2.BuyReturnAmount
			                              - T2.FromStoreAmount - T2.LossAmount
			                              + T2.ToStoreAmount
	                     WHEN @Type = 3
			                         THEN T2.firstMojodi+T2.BuyAmount - T2.BuyReturnAmount
			                              - T2.FromStoreAmount - T2.LossAmount
			                              + T2.ToStoreAmount
			            ELSE T2.Mojodi  
					END  

    FROM    dbo.tblTotal_tInventory_tGood_For_Mojodi(@intLanguage, @SystemDate,
                                                     @SystemDay, @SystemTime,
                                                     @DateBefore, @DateAfter,
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

