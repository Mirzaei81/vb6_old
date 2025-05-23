
--ScriptV26_16_Fix_16_موجودی.SQL

-- کنترل موجودی ها
--95/04/23


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER PROCEDURE [dbo].[InsertMojodiCalculate]
    (
      @Status INT,
      @intserialNo BIGINT,
      @AccountYear SMALLINT,
      @Branch INT = NULL
    )
AS 
    IF @Branch IS NULL 
        SET @Branch = dbo.Get_Current_Branch()

---------------------------------------Mojodi Control Online---------------------------------------------------------

    IF @Status = 2 
	    --IF dbo.AutoHavale() = 0
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - CASE WHEN dbo.AutoHavale() = 0 THEN t.Amount ELSE 0 END,
                    SaleAmount = SaleAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - CASE WHEN dbo.AutoHavale() = 0 THEN X.Amount ELSE 0 END ,
                    SaleAmount = SaleAmount + X.Amount
            FROM    ( SELECT    SUM(( Amount * fltUsedValue ) + ( [Amount] * [Pert] )) AS Amount,
                                GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType = 4 
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) X
            WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = X.intInventoryNo
                    --AND tInventory_Good.Branch = X.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

 	    UPDATE  tInventory_Good	--Mojodi not less zero because in edit mode not show message
            SET     Mojodi = 0
            FROM    ( SELECT    GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType = 4
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) X
            WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = X.intInventoryNo
                    --AND tInventory_Good.Branch = X.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
		    AND tInventory_Good.Mojodi < 0


        END
    IF @Status = 1 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + t.Amount,
                    BuyAmount = BuyAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 3 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    LossAmount = LossAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear


        END
    IF @Status = 4 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    BuyReturnAmount = BuyReturnAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 5 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + t.Amount,
                    SaleReturnAmount = SaleReturnAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 6 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    FromStoreAmount = FromStoreAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

        END
    IF @Status = 7 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + t.Amount,
                    toStoreAmount = toStoreAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

        END
--===============================================

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[DeleteMojodiCalculate]
    (
      @Status INT ,
      @intserialNo BIGINT ,
      @Recursive INT ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS ---------------------------------------Mojodi Control Online---------------------------------------------------------
    IF @Recursive = 1 
        BEGIN
            IF @Status = 2 
		        --IF dbo.AutoHavale() = 0
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + CASE WHEN dbo.AutoHavale() = 0 THEN t.Amount ELSE 0 END  ,
                            SaleAmount = SaleAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + CASE WHEN dbo.AutoHavale() = 0 THEN Amount ELSE 0 END ,
                            SaleAmount = SaleAmount - Amount
                    FROM    ( SELECT    SUM(( Amount * fltUsedValue )
                                            + ( [Amount] * [Pert] )) AS Amount ,
                                        GoodFirstCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      ( SELECT    *
                                          FROM      tFacd
                                                    INNER JOIN UsePercent ON tFacd.GoodCode = usepercent.code
                                                              AND tFacd.serveplace = usepercent.intserveplace
                                          WHERE     intserialNo = @intserialNo
                                                    AND Branch = @Branch
                                        ) FirstGoods
                                        INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType = 4 
                              GROUP BY  FirstGoods.GoodFirstCode ,
                                        FirstGoods.intInventoryNo ,
                                        FirstGoods.Branch
                            ) X
                    WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                            AND tInventory_Good.InventoryNo = X.intInventoryNo
                            AND tInventory_Good.Branch = X.Branch
                            AND tInventory_Good.AccountYear = @AccountYear

                END
            IF @Status = 1 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - t.Amount ,
                            BuyAmount = BuyAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 3 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            LossAmount = LossAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 4 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            BuyReturnAmount = BuyReturnAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 5 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - t.Amount ,
                            SaleReturnAmount = SaleReturnAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 6 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            FromStoreAmount = FromStoreAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 7 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - t.Amount ,
                            toStoreAmount = FromStoreAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
        END

    ELSE 
        IF @Recursive = 0 
            BEGIN
                IF @Status = 2 
	   	            --IF dbo.AutoHavale() = 0
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - CASE WHEN dbo.AutoHavale() = 0 THEN t.Amount ELSE 0 END ,
                                SaleAmount = SaleAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - CASE WHEN dbo.AutoHavale() = 0 THEN Amount ELSE 0 END ,
                                SaleAmount = SaleAmount - Amount
                        FROM    ( SELECT    SUM(( Amount * fltUsedValue )
                                                + ( [Amount] * [Pert] )) AS Amount ,
                                            GoodFirstCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      ( SELECT    *
                                              FROM      tFacd
                                                        INNER JOIN UsePercent ON tFacd.GoodCode = usepercent.code
                                                              AND tFacd.serveplace = usepercent.intserveplace
                                              WHERE     intserialNo = @intserialNo
                                                        AND Branch = @Branch
                                            ) FirstGoods
                                            INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code  AND dbo.tGood.GoodType = 4  
                                  GROUP BY  FirstGoods.GoodFirstCode ,
                                            FirstGoods.intInventoryNo ,
                                            FirstGoods.Branch
                                ) X
                        WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                                AND tInventory_Good.InventoryNo = X.intInventoryNo
                                AND tInventory_Good.Branch = X.Branch
                                AND tInventory_Good.AccountYear = @AccountYear

                    END
                IF @Status = 1 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi + t.Amount ,
                                BuyAmount = BuyAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
                IF @Status = 3 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - t.Amount ,
                                LossAmount = LossAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
                IF @Status = 4 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - t.Amount ,
                                BuyReturnAmount = BuyReturnAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
                IF @Status = 5 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi + t.Amount ,
                                SaleReturnAmount = SaleReturnAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
            IF @Status = 6 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - t.Amount ,
                            FromStoreAmount = FromStoreAmount + t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 7 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            toStoreAmount = FromStoreAmount + t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            END
--===============================================

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    PROCEDURE Update_tFacM_Recursive
(
@No  Bigint,
@Status int,
@Recursive int,
@Uid int,
@Balance Bit,
@FacPayment Bit ,
@AccountYear Smallint = NULL ,
@Branch INT 
)

AS
Declare @TableNo int
DECLARE @intTableUsedNo INT      
IF @AccountYear Is Null 
	SET @AccountYear = dbo.Get_AccounYear()

DECLARE @intSerialNo BIGINT

--DECLARE @Branch INT
--	SET @Branch = dbo.Get_Current_Branch()

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch and AccountYear = @AccountYear)

UPDATE tFacM
     SET Recursive= @Recursive
         WHERE tFacM.intSerialNo = @intserialNo And  Branch = @Branch 


DECLARE @intserialNo2 BIGINT
If @Status = 6 OR (@Status = 2 AND dbo.AutoHavale() = 1)
BEGIN 
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear)  
	IF @intSerialNo2 > 0
		UPDATE tFacM
			 SET Recursive= @Recursive
				 WHERE tFacM.intSerialNo = @intserialNo2 And  Branch = @Branch 
END 

If @Recursive = 1 
Begin

UPDATE tFacM
     SET FacPayment = 0 , Balance = 0
         WHERE tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear

SET @TableNo = ISNULL((Select TableNo From tfacm   Where  tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear) , 0)
  UPDATE tTable
       SET Empty = 1 
           WHERE dbo.tTable.[No] = @TableNo
	If dbo.Get_TableMonitoring() = 1 AND @TableNo > 0		---Table Monitoring
	Begin
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.bitIsValid = 0      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	


Exec DeleteFactorChildren @intSerialNo , @Branch

UPDATE dbo.tblAcc_Recieved SET Bestankar = 0 WHERE intSerialNo = @intSerialNo And  Branch = @Branch  

End

If @Recursive = 0

Begin
   Update tFacm 
       SET FacPayment = @FacPayment , Balance = @Balance
         WHERE tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear

	SET @TableNo = ISNULL((Select TableNo From tfacm   Where  tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear) , 0)
    UPDATE tTable
       SET Empty = 0 
           WHERE dbo.tTable.[No] = @TableNo
	If dbo.Get_TableMonitoring() = 1 AND @TableNo > 0		---Table Monitoring
	Begin
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.bitIsValid = 1      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	

	IF @Balance = 1
	BEGIN 
	DELETE FROM tFacCash WHERE intSerialNo = @intSerialNo AND [Branch] = @Branch
	INSERT INTO tFacCash (intSerialNo, intAmount ,branch)
		SELECT @intSerialNo AS
	 intSerialNo, Sumprice,@Branch From tFacM  WHERE tFacM.[No]=@No   AND Status = 2 And  Branch = @Branch and AccountYear = @AccountYear

	END 
End

--Declare @Monitor1 Bit
--Declare @Monitor2 Bit

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())


--If @Monitor1 > 0 
--  exec Notify_to_Clients
--Else If @Monitor2 > 0 
--  exec Notify_to_Clients

If @Recursive = 0
   Exec InsertHistory  @No, @Status , @Uid , 8 ,@AccountYear , @Branch
Else if @Recursive = 1
   Exec InsertHistory  @No, @Status , @Uid , 3 ,@AccountYear , @Branch 

---------------------------------------Mojodi Control Online---------------------------------------------------------

Exec DeleteMojodiCalculate @Status , @intserialNo , @Recursive ,@AccountYear , @Branch
If @Status = 6  AND @intserialNo2 > 0
	EXEC DeleteMojodiCalculate 7, @intSerialNo2 , @Recursive, @AccountYear , @Branch
If  (@Status = 2 AND dbo.AutoHavale() = 1)  AND @intserialNo2 > 0
	EXEC DeleteMojodiCalculate 6, @intSerialNo2 , @Recursive, @AccountYear , @Branch

--------------------------------------------------------------------------------------------------------------------------------------

---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @intSerialNo , 3

--------------------------------------------------------------------------------------------------------------------------------------
GO
