
--For Jarchi
--Upgrade to V26_16_Fix10
--94/02/13

UPDATE tfacm SET transferAccounting = 0 WHERE AccountYear = 1394
GO


DELETE FROM dbo.tblAcc_DocumentHeader
Go

DELETE FROM dbo.tblAcc_Tafsili WHERE TafsiliId > 0
GO

UPDATE dbo.tPer SET Tafsili = NULL 
UPDATE dbo.tCust SET Tafsili = NULL
GO


----
--Script_V26_16_Fix8
--اضافه شدن دسترسی حواله و رسید موقت به دسترسی ها
--امکان صدور حواله روزانه برای انبارهای مختلف در فیش های چند انباره
-- 93/09/02

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
          9
        )
GO

DELETE FROM dbo.tObjects WHERE intObjectCode = 326 OR intObjectCode = 333 OR intObjectCode =334 OR intObjectCode =335 
GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 326 , -- intObjectCode - int
          N'frmSavePurchase' , -- ObjectId - nvarchar(50)
          N'ثبت خریدها' , -- ObjectName - nvarchar(50)
          N'frmSavePurchase' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
GO
        
        
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          326  -- intObjectCode - int
          )
GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 333 , -- intObjectCode - int
          N'frmHavaleh' , -- ObjectId - nvarchar(50)
          N'فرم حواله' , -- ObjectName - nvarchar(50)
          N'frmHavaleh' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
GO
        
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          333  -- intObjectCode - int
          )
GO

INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 334 , -- intObjectCode - int
          N'frmTempReceived' , -- ObjectId - nvarchar(50)
          N'فرم رسید موقت' , -- ObjectName - nvarchar(50)
          N'frmTempReceived' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
GO
        
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          334  -- intObjectCode - int
          )
GO

        
INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 335 , -- intObjectCode - int
          N'frmLosses' , -- ObjectId - nvarchar(50)
          N'ورود ضایعات' , -- ObjectName - nvarchar(50)
          N'frmLosses' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
GO
        
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          335  -- intObjectCode - int
          )
GO
        


IF COL_LENGTH('tFacD','BitHavaleResid') IS NULL
BEGIN
	ALTER TABLE tFacD
	ADD BitHavaleResid BIT NOT NULL DEFAULT(0)
END

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Update_BitHavaleResid]
    (
      @Branch INT,
      @InventoryNo INT,
      @AccountYear SMALLINT,
      @Status INT,
      @FromDate NVARCHAR(8),
      @ToDate NVARCHAR(8)
    )
AS 
    BEGIN
        UPDATE  tFacd
        SET     bitHavaleResid = 1
        FROM    ( SELECT    tfacm.intSerialNo,
                            tfacm.Branch , tfacd.intRow 
                  FROM      tfacm
                            INNER JOIN tfacd ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                  WHERE     tfacm.Branch = @Branch
                            AND tfacm.Recursive = 0
                            AND tfacm.status = @Status
                            AND tfacm.Date >= @FromDate
                            AND tfacm.Date <= @ToDate
                            AND dbo.tFacD.BitHavaleResid = 0
                            AND tfacd.intInventoryNo = @InventoryNo
                ) T
        WHERE   T.intSerialNo = tFacd.intSerialNo
                AND t.Branch = tFacd.Branch
                AND tfacd.intRow = T.intRow
                AND tfacd.intInventoryNo = @InventoryNo
    END
--===============================================
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   PROCEDURE [dbo].[Get_DailyGoodForHavale]
    (
      @Branch INT,
      @InventoryNo INT,
      @AccountYear SMALLINT,
      @Status INT,
      @FromDate NVARCHAR(8),
      @ToDate NVARCHAR(8)
    )
AS 
    BEGIN
	DECLARE @LossStatus INT 
	IF @Status = 2 SET @LossStatus = 3
	ELSE IF @Status <> 2 SET @LossStatus = 0

        SELECT  CAST(SUM(T.Amount) AS DECIMAL(19,3)) AS Amount,
                T.[Name],
                T.BuyPrice AS FeeUnit,
                T.GoodCode,
                T.Unit,
                0 AS Discount , --T.Discount,
                T.Weight,
                T.NumberOfUnit,
                Rate,
                ( SELECT    dbo.tInventory.Description
                  FROM      dbo.tInventory
                  WHERE     dbo.tInventory.InventoryNo = intInventoryNo
                ) AS intInventoryNo,
                ( SELECT    ISNULL(dbo.tInventory.Description, '')
                  FROM      dbo.tInventory
                  WHERE     dbo.tInventory.InventoryNo = DestInventoryNo
                ) AS DestInventoryNo,
                0 AS Duty
        FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                            dbo.tFacD.Rate,
                            dbo.tFacD.intInventoryNo,
                            dbo.tFacD.DestInventoryNo,
                            ( dbo.tFacD.Amount
                              * ( dbo.tUsePercent.fltUsedValue
                                  + ISNULL(dbo.tUsePercent.Pert, 0) ) ) AS Amount,
                            ( SELECT    Name
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [Name],
                            dbo.tGood.Weight,
                            ( SELECT    Unit
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS Unit,
                            ( SELECT    BuyPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS BuyPrice,
                            dbo.tGood.NumberOfUnit,
                            0 AS Discount --dbo.tFacD.Discount
                  FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                    AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                            JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                  WHERE     tfacm.Branch = @Branch
                            AND tfacm.Recursive = 0
                            AND (tfacm.status = @Status OR  dbo.tFacM.Status = @LossStatus)
                            AND tfacm.Date >= @FromDate
                            AND tfacm.Date <= @ToDate
                            AND tFacD.bitHavaleResid = 0
                            AND tfacd.intInventoryNo = @InventoryNo
                            AND ( SELECT    dbo.tGood.GoodType
                                  FROM      dbo.tGood
                                  WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                ) <> 4
                  UNION ALL
                  SELECT    dbo.tFacD.GoodCode,
                            dbo.tFacD.Rate,
                            dbo.tFacD.intInventoryNo,
                            dbo.tFacD.DestInventoryNo,
                            dbo.tFacD.Amount,
                            dbo.tGood.Name,
                            dbo.tGood.Weight,
                            dbo.tGood.Unit,
                            dbo.tGood.BuyPrice,
                            dbo.tGood.NumberOfUnit ,
                            0 AS Discount --dbo.tFacD.Discount
                  FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                  WHERE     dbo.tFacM.Branch = @Branch
                            AND tfacm.Recursive = 0
                            AND tfacm.status = @Status
                            AND tfacm.Date >= @FromDate
                            AND tfacm.Date <= @ToDate
                            AND tfacD.BitHavaleResid = 0
                            AND tfacd.intInventoryNo = @InventoryNo
                            AND dbo.tFacD.GoodCode NOT IN (
                            SELECT  dbo.tUsePercent.GoodCode
                            FROM    dbo.tUsePercent
                            WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                            AND dbo.tGood.GoodType = 3
                ) T
        GROUP BY T.GoodCode,
                T.Name,
                T.Weight,
                T.Unit,
                Rate,
                intInventoryNo,
                DestInventoryNo,
                --T.Discount,
                T.NumberOfUnit,
                T.BuyPrice

    END

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Insert_AutoHavale]
    (
      @Branch INT,
      @InventoryNo INT,
      @AccountYear SMALLINT,
      @Status INT,
      @FromDate NVARCHAR(8),
      @ToDate NVARCHAR(8),
      @Date NVARCHAR(8),
      @User INT,
      @NvcDescription NVARCHAR(150),
      @Result INT OUT
		
    )
AS 
    BEGIN
        BEGIN TRAN
	DECLARE @LossStatus INT 
	IF @Status = 2 SET @LossStatus = 3
	ELSE IF @Status <> 2 SET @LossStatus = 0

        DECLARE @intSerialNo INT
        DECLARE @No INT 

        IF @Status = 2 
            SELECT  @NO = ISNULL(MAX([NO]), 0) + 1
            FROM    tFacM
            WHERE   Status = 6
                    AND Branch = @Branch
                    AND AccountYear = @AccountYear

        DECLARE @ReturnTable TABLE
            (
              Row INT IDENTITY(1, 1)
                      NOT NULL,
              Amount FLOAT NOT NULL,
              GoodCode INT NOT NULL,
              BuyPrice FLOAT NOT NULL
            )
 
        INSERT  INTO @ReturnTable
                (
                  Amount,
                  GoodCode,
                  BuyPrice
                )
                SELECT  CAST(SUM(T.Amount) AS DECIMAL(19,3)) ,
                        T.GoodCode,
                        T.BuyPrice
                FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                                    ( dbo.tFacD.Amount
                                      * ( dbo.tUsePercent.fltUsedValue
                                          + ISNULL(dbo.tUsePercent.Pert, 0) ) ) AS Amount,
                                    ( SELECT    BuyPrice
                                      FROM      dbo.tGood
                                      WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                    ) AS BuyPrice
                          FROM      dbo.tFacM
                                    JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                      AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                            AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                                    JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                          WHERE     tfacm.Branch = @Branch
                                    AND tfacm.Recursive = 0
                                    AND (tfacm.status = @Status OR dbo.tFacM.Status = @LossStatus)
                                    AND tfacm.Date >= @FromDate
                                    AND tfacm.Date <= @ToDate
                                    AND tfacd.bitHavaleResid = 0
                                    AND tfacd.intInventoryNo = @InventoryNo
                                    AND ( SELECT    dbo.tGood.GoodType
                                          FROM      dbo.tGood
                                          WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                        ) <> 4
                          UNION ALL
                          SELECT    dbo.tFacD.GoodCode,
                                    dbo.tFacD.Amount,
                                    dbo.tGood.BuyPrice
                          FROM      dbo.tFacM
                                    JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                      AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                          WHERE     dbo.tFacM.Branch = @Branch
                                    AND tfacm.Recursive = 0
                                    AND tfacm.status = @Status
                                    AND tfacm.Date >= @FromDate
                                    AND tfacm.Date <= @ToDate
                                    AND tfacd.bitHavaleResid = 0
                                    AND tfacd.intInventoryNo = @InventoryNo
                                    AND dbo.tFacD.GoodCode NOT IN (
                                    SELECT  dbo.tUsePercent.GoodCode
                                    FROM    dbo.tUsePercent
                                    WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                                    AND dbo.tGood.GoodType = 3
                        ) T
                GROUP BY T.GoodCode,
                        T.BuyPrice


		
        DECLARE @SumTotal INT 
        SELECT  @SumTotal = SUM(BuyPrice * Amount )
        FROM    @ReturnTable

		DECLARE @ShiftNo INT 
		DECLARE @TempNo INT 

		SET @ShiftNo= dbo.Get_Shift(GETDATE())      
		SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @Branch AND Date = @Date AND ShiftNo = @ShiftNo)      
		   DECLARE @IdentityNo INT
		    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
		    FROM    tFacm
		    WHERE   Branch = @Branch 

    IF @IdentityNo < ( @Branch * 10000000 ) 
        SET @IdentityNo = ( @Branch * 10000000 )


        INSERT  INTO tFacM
                (
                  intSerialNo ,
                  [No],
                  [Date],
                  RegDate,
                  Status,
                  Customer,
                  SumPrice,
                  StationId,
                  Recursive,
                  CarryFeeTotal,
                  PackingTotal,
                  DiscountTotal,
                  [Time],
                  [User],
                  shiftNo,
                  incharge,
                  FacPayment,
                  Balance,
                  Branch,
                  AccountYear,
                  NvcDescription,
                  OrderType ,
                  TempNo
			  
                )
        VALUES  (
                  @IdentityNo ,
                  @NO,
                  @Date,
                  dbo.Shamsi(GETDATE()),
                  6,
                  -1,
                  @SumTotal,
                  1,
                  0,
                  0,
                  0,
                  0,
                  dbo.SetTimeFormat(GETDATE()),
                  @User,
                  @ShiftNo ,--dbo.Get_Shift(GETDATE()) ,
                  NULL,
                  0,
                  0,
                  @Branch, --dbo.Get_Current_Branch(),
                  @AccountYear,
                  @NvcDescription,
                  2 ,
                  @TempNo
                )
        IF @@ERROR <> 0 
            GOTO EventHandler

        SET @intserialNo = @IdentityNo 

 
        INSERT  INTO tFacD
                (
                  intRow,
                  Amount,
                  GoodCode,
                  FeeUnit,
                  Discount,
                  Rate,
                  [ExpireDate],
                  intInventoryNo,
                  DestInventoryNo, --Because Has a Relation and Can not insert for  Another Branch
                  DifferencesCodes,
                  DifferencesDescription,
                  intSerialNo,
                  Branch,
                  ServePlace
                )
                SELECT  tmpTable.Row,
                        tmpTable.Amount,
                        tmpTable.GoodCode,
                        BuyPrice, --FeeUnit ,
                        0, --Discount ,
                        1, --tmpTable.Rate ,
                        '', --tmpTable.[ExpireDate] ,
                        @InventoryNo,
                        NULL,
                        '', --DifferencesCode ,
                        '', --.DifferencesDescription ,
                        @intSerialNo,
                        @Branch, --dbo.Get_Current_Branch(),
                        1
                FROM    @ReturnTable tmpTable

	--DROP TABLE @ReturnTable
		Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @Branch      

		IF @@ERROR <>0      
		GoTo EventHandler      

        EXEC Update_BitHavaleResid @Branch, @InventoryNo, @AccountYear,
            @Status, @FromDate, @ToDate

        COMMIT TRAN
        SET @Result = 1

        RETURN @Result

        EventHandler:
        ROLLBACK TRAN
        SET @Result = -1
        RETURN @Result
    END


GO




IF COL_LENGTH('tInventory','Tafsili2') IS NULL
BEGIN
	ALTER TABLE dbo.tInventory
	ADD Tafsili2 INT NULL
END
--ALTER TABLE dbo.tInventory DROP COLUMN Tafsili2
--GO
	

GO

INSERT INTO dbo.tblAcc_Moein
        ( KolId ,
          MoeinId ,
          MoeinName ,
          Kind ,
          Active
        )
VALUES  ( 43 , -- KolId - int
          4304 , -- MoeinId - int
          N'قیمت تمام شده' , -- MoeinName - nvarchar(50)
          1 , -- Kind - tinyint
          1  -- Active - bit
        )
GO


INSERT INTO dbo.TblAcc_Sale
        ( Code ,
          Description ,
          Kol ,
          Moein ,
          Tafsili ,
          Active ,
          MoeinDesc
        )
VALUES  ( 39 , -- Code - int
          N'مرکز هزینه' , -- Description - nvarchar(50)
          43 , -- Kol - int
          4304 , -- Moein - int
          0 , -- Tafsili - int
          1 , -- Active - bit
          N'قیمت تمام شده'  -- MoeinDesc - nvarchar(50)
        )
        
GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    PROCEDURE [dbo].[Insert_AutoHavale]
    (
      @Branch INT,
      @InventoryNo INT,
      @AccountYear SMALLINT,
      @Status INT,
      @FromDate NVARCHAR(8),
      @ToDate NVARCHAR(8),
      @Date NVARCHAR(8),
      @User INT,
      @NvcDescription NVARCHAR(150),
      @Result INT OUT
		
    )
AS 
    BEGIN
        BEGIN TRAN
	DECLARE @LossStatus INT 
	IF @Status = 2 SET @LossStatus = 3
	ELSE IF @Status <> 2 SET @LossStatus = 0

        DECLARE @intSerialNo INT
        DECLARE @No INT 

        IF @Status = 2 
            SELECT  @NO = ISNULL(MAX([NO]), 0) + 1
            FROM    tFacM
            WHERE   Status = 6
                    AND Branch = @Branch
                    AND AccountYear = @AccountYear

        DECLARE @ReturnTable TABLE
            (
              Row INT IDENTITY(1, 1)
                      NOT NULL,
              Amount FLOAT NOT NULL,
              GoodCode INT NOT NULL,
              BuyPrice FLOAT NOT NULL
            )
 
        INSERT  INTO @ReturnTable
                (
                  Amount,
                  GoodCode,
                  BuyPrice
                )
                SELECT  CAST(SUM(T.Amount) AS DECIMAL(19,3)) ,
                        T.GoodCode,
                        T.BuyPrice
                FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                                    ( dbo.tFacD.Amount
                                      * ( dbo.tUsePercent.fltUsedValue
                                          + ISNULL(dbo.tUsePercent.Pert, 0) ) ) AS Amount,
                                    ( SELECT    BuyPrice
                                      FROM      dbo.tGood
                                      WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                    ) AS BuyPrice
                          FROM      dbo.tFacM
                                    JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                      AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                            AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                                    JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                          WHERE     tfacm.Branch = @Branch
                                    AND tfacm.Recursive = 0
                                    AND (tfacm.status = @Status OR dbo.tFacM.Status = @LossStatus)
                                    AND tfacm.Date >= @FromDate
                                    AND tfacm.Date <= @ToDate
                                    AND tfacd.bitHavaleResid = 0
                                    AND tfacd.intInventoryNo = @InventoryNo
                                    AND ( SELECT    dbo.tGood.GoodType
                                          FROM      dbo.tGood
                                          WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                        ) <> 4
                          UNION ALL
                          SELECT    dbo.tFacD.GoodCode,
                                    dbo.tFacD.Amount,
                                    dbo.tGood.BuyPrice
                          FROM      dbo.tFacM
                                    JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                      AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                          WHERE     dbo.tFacM.Branch = @Branch
                                    AND tfacm.Recursive = 0
                                    AND tfacm.status = @Status
                                    AND tfacm.Date >= @FromDate
                                    AND tfacm.Date <= @ToDate
                                    AND tfacd.bitHavaleResid = 0
                                    AND tfacd.intInventoryNo = @InventoryNo
                                    AND dbo.tFacD.GoodCode NOT IN (
                                    SELECT  dbo.tUsePercent.GoodCode
                                    FROM    dbo.tUsePercent
                                    WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                                    AND dbo.tGood.GoodType = 3
                        ) T
                GROUP BY T.GoodCode,
                        T.BuyPrice


		
        DECLARE @SumTotal INT 
        SELECT  @SumTotal = SUM(BuyPrice * Amount )
        FROM    @ReturnTable

		DECLARE @ShiftNo INT 
		DECLARE @TempNo INT 

		SET @ShiftNo= dbo.Get_Shift(GETDATE())      
		SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @Branch AND Date = @Date AND ShiftNo = @ShiftNo)      
		   DECLARE @IdentityNo INT
		    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
		    FROM    tFacm
		    WHERE   Branch = @Branch 

    IF @IdentityNo < ( @Branch * 10000000 ) 
        SET @IdentityNo = ( @Branch * 10000000 )


        INSERT  INTO tFacM
                (
                  intSerialNo ,
                  [No],
                  [Date],
                  RegDate,
                  Status,
                  Customer,
                  SumPrice,
                  StationId,
                  Recursive,
                  CarryFeeTotal,
                  PackingTotal,
                  DiscountTotal,
                  [Time],
                  [User],
                  shiftNo,
                  incharge,
                  FacPayment,
                  Balance,
                  Branch,
                  AccountYear,
                  NvcDescription,
                  OrderType ,
                  TempNo
			  
                )
        VALUES  (
                  @IdentityNo ,
                  @NO,
                  @Date,
                  dbo.Shamsi(GETDATE()),
                  6,
                  -1,
                  @SumTotal,
                  1,
                  0,
                  0,
                  0,
                  0,
                  dbo.SetTimeFormat(GETDATE()),
                  @User,
                  @ShiftNo ,--dbo.Get_Shift(GETDATE()) ,
                  NULL,
                  0,
                  0,
                  @Branch, --dbo.Get_Current_Branch(),
                  @AccountYear,
                  @NvcDescription,
                  2 ,
                  @TempNo
                )
        IF @@ERROR <> 0 
            GOTO EventHandler

        SET @intserialNo = @IdentityNo 

 
        INSERT  INTO tFacD
                (
                  intRow,
                  Amount,
                  GoodCode,
                  FeeUnit,
                  Discount,
                  Rate,
                  [ExpireDate],
                  intInventoryNo,
                  DestInventoryNo, --Because Has a Relation and Can not insert for  Another Branch
                  DifferencesCodes,
                  DifferencesDescription,
                  intSerialNo,
                  Branch,
                  ServePlace
                )
                SELECT  tmpTable.Row,
                        tmpTable.Amount,
                        tmpTable.GoodCode,
                        BuyPrice, --FeeUnit ,
                        0, --Discount ,
                        1, --tmpTable.Rate ,
                        '', --tmpTable.[ExpireDate] ,
                        @InventoryNo,
                        NULL,
                        '', --DifferencesCode ,
                        '', --.DifferencesDescription ,
                        @intSerialNo,
                        @Branch, --dbo.Get_Current_Branch(),
                        1
                FROM    @ReturnTable tmpTable

	--DROP TABLE @ReturnTable
		Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @Branch      

		IF @@ERROR <>0      
		GoTo EventHandler      

        EXEC Update_BitHavaleResid @Branch, @InventoryNo, @AccountYear,
            @Status, @FromDate, @ToDate

        COMMIT TRAN
        SET @Result = @NO

        RETURN @Result

        EventHandler:
        ROLLBACK TRAN
        SET @Result = -1
        RETURN @Result
    END


GO



--کسر مرجوعی ها از وجه نقد و کارت 
--93/10/20

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
--New

ALTER  VIEW [dbo].[VwStationSaleSummery]
AS 
    SELECT  dbo.tFacM.[No] ,
            dbo.tFacM.[Date] ,
            dbo.tFacM.[Time] ,
            dbo.tFacM.[User] ,
            CASE WHEN tfacm.Recursive = 0 THEN SumPrice ELSE 0 END AS SumPrice ,
            CASE WHEN tfacm.Recursive = 0 THEN CarryFeeTotal ELSE 0 END AS CarryFeeTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN DiscountTotal ELSE 0 END AS DiscountTotal ,
            StationID ,
            CASE WHEN tfacm.Recursive = 0 THEN ServiceTotal ELSE 0 END AS ServiceTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN PackingTotal ELSE 0 END AS PackingTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN TaxTotal ELSE 0 END AS TaxTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN DutyTotal ELSE 0 END AS DutyTotal ,
	        CASE WHEN tfacm.Recursive = 0 THEN dbo.tfacm.[RoundDiscount] ELSE 0 END AS RoundDiscount,
            FacPayment ,
            Balance ,
            dbo.tfacm.Customer ,
            CASE dbo.tCust.[Name] + SPACE(3) + dbo.tCust.family
              WHEN NULL THEN dbo.tCust.WorkName
              WHEN '' THEN dbo.tCust.WorkName
              ELSE dbo.tCust.[Name] + '  ' + dbo.tCust.family
            END AS CustomerName ,
            dbo.tCust.tafsili AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            CASE WHEN Recursive = 1 THEN 0
				WHEN ISNULL(Incharge, 0) = 0 THEN 0
                ELSE CASE ISNULL(TableNo, 0)
                     WHEN 0 THEN SumPrice
                     ELSE 0
                   END
            END AS CarrierSumPrice ,
            CASE WHEN Recursive = 1 THEN 0
				WHEN ISNULL(Incharge, 0) = 0 THEN 0
				ELSE CASE ISNULL(TableNo, 0)
                     WHEN 0 THEN 0
                     ELSE SumPrice
                   END
				END AS GarsonSumPrice ,
            CASE  WHEN Recursive = 1 THEN 0
				  WHEN FacPayment = 0 
                   THEN CASE Balance
                            WHEN 0 THEN CASE ISNULL(Incharge, 0)
                                          WHEN 0 THEN 0
                                          ELSE CASE ISNULL(TableNo, 0)
                                                 WHEN 0 THEN SumPrice
                                                 ELSE 0
                                               END
                                        END
                            ELSE 0
                          END
              ELSE 0
            END AS CarrierDebit ,
            CASE WHEN Recursive = 1 THEN 0
				 WHEN FacPayment = 0 
					THEN CASE ISNULL(Incharge, 0)
								WHEN 0 THEN 0
								ELSE CASE ISNULL(TableNo, 0)
									   WHEN 0 THEN 0
									   ELSE SumPrice
									 END
							  END
				  ELSE 0
            END AS GarsonDebit ,
--            CASE Balance
--              WHEN 0 THEN CASE FacPayment
--                            WHEN 0 THEN 0
--                            ELSE SumPrice
--                          END
--              ELSE 0
--            END AS CustomerDebit ,
			CASE WHEN Recursive = 1 THEN 0
				WHEN Balance = 0 THEN 
				CASE WHEN (Facpayment = 1 or (Incharge is NULL AND serveplace <> 2 AND serveplace <> 16))THEN 
					SumPrice - (ISNULL(Resived.Received , 0) +ISNULL(CardReceived.CardReceived , 0)+ISNULL(PreReceived2.PreReceived2 , 0))
					ELSE 0
					END 
				ELSE 0
			END AS CustomerDebit ,
            CASE WHEN Recursive = 1 THEN 0 
              WHEN Balance = 0 
					THEN CASE FacPayment
                            WHEN 0 THEN CASE ISNULL(Incharge, 0)
                                          WHEN 0 THEN SumPrice
                                          ELSE 0
                                        END
                            ELSE 0
                          END
              ELSE   0
            END AS UnBalanceFich ,
            dbo.tFacM.Branch ,
            0 AS Payment ,
            ISNULL(Resived.Received , 0) AS Recieved ,
            tper.ppno ,
            tfacm.status ,
            ISNULL(CardReceived.CardReceived , 0) AS CardReceived ,
            ISnull(tFactorAdditionalServices.amount , 0)  AS TipAmount ,
			0 AS ManualRecieved , 0 AS OrderPrice , 0 AS OrderReceived ,
			CASE WHEN tfacm.Recursive = 1 THEN SumPrice ELSE 0 END AS SumRecursive 
			
    FROM    dbo.tFacM
        INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
        INNER JOIN tCUst ON tfacM.Customer = tcust.code
       -- LEFT OUTER JOIN [tblAcc_Recieved] ON tfacm.[OrderRefrence] = [tblAcc_Recieved].[intSerialNo] 
        LEFT OUTER JOIN [tFactorAdditionalServices] ON [tFactorAdditionalServices].[Branch] = [tFacM].[Branch] AND [tFactorAdditionalServices].[intSerialNo] = [tFacM].[intSerialNo]
			AND tFactorAdditionalServices.intServiceNo = 3 AND dbo.tFacM.Recursive = 0
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS Received,intSerialNo , Branch FROM  [dbo].[tFacCash] 
			GROUP BY intSerialNo , Branch) AS Resived ON Resived.Branch = tfacm.Branch AND  Resived.intSerialNo = dbo.tFacM.intSerialNo  AND dbo.tFacM.Recursive = 0
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS CardReceived,intSerialNo , Branch FROM  [dbo].[tFacCard] 
			GROUP BY intSerialNo , Branch) AS CardReceived ON CardReceived.Branch = tfacm.Branch AND  CardReceived.intSerialNo = dbo.tFacM.intSerialNo  AND dbo.tFacM.Recursive = 0
	    LEFT outer JOIN (SELECT ISNULL(SUM(Bestankar),0) AS PreReceived2 ,intSerialNo , Branch FROM  [dbo].[tblAcc_Recieved] 
			GROUP BY intSerialNo , Branch) AS PreReceived2 ON PreReceived2.Branch = tfacm.Branch AND  PreReceived2.intSerialNo = dbo.tFacM.intSerialNo  AND dbo.tFacM.Recursive = 0

    WHERE  Status = 2

    UNION
    SELECT  dbo.tFacM.[No] ,
            dbo.tFacM.[Date] ,
            dbo.tFacM.[Time] ,
            dbo.tFacM.[User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
	        0 AS [RoundDiscount],
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tFacM.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    SumPrice AS OrderPrice ,
	    0 AS OrderReceived ,
	    0 AS SumRecursive
    FROM    dbo.tFacM 
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
            INNER JOIN tCust ON tfacM.Customer = tcust.code
    WHERE Recursive = 0 AND Status = 10
    UNION
    SELECT  tblAcc_Cash.Code AS [No] ,
            tblAcc_Cash.[Date] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Cash.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Cash.Branch AS Branch ,
            Bestankar AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived ,
	    0 AS SumRecursive
    FROM    tblAcc_Cash
            INNER JOIN dbo.tUser ON dbo.tblAcc_Cash.[UID] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,--list
            tFacM.[Date] AS [Date] ,  --            tblAcc_Recieved.[Date] AS [Date] , --Date From tFacm
            RegTime AS [Time] ,
            tblAcc_Recieved.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved.Branch AS Branch ,
            0 AS Payment ,
            Bestankar AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS OrderPrice ,
  	    0 AS OrderReceived ,
	    0 AS SumRecursive

    FROM    tblAcc_Recieved 
		INNER JOIN dbo.tFacM ON dbo.tblAcc_Recieved.Branch = dbo.tFacM.Branch AND dbo.tblAcc_Recieved.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
		INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
	WHERE tblAcc_Recieved.intSerialNo IS NOT NULL AND Status = 2
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,
            tblAcc_Recieved.[Date] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Recieved.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS OrderPrice ,
  	    Bestankar AS OrderReceived ,
	    0 AS SumRecursive

    FROM    tblAcc_Recieved 
	    INNER JOIN dbo.tFacM ON dbo.tblAcc_Recieved.Branch = dbo.tFacM.Branch AND dbo.tblAcc_Recieved.intSerialNo = dbo.tFacM.intSerialNo
        INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    WHERE tblAcc_Recieved.intSerialNo IS NOT NULL AND Status = 10
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,
            tblAcc_Recieved.[Date] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Recieved.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    Bestankar AS ManualRecieved ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived ,
	    0 AS SumRecursive
    FROM    tblAcc_Recieved 
        INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    WHERE intSerialNo IS NULL 


GO



SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


-------------------------------------------------
-------------------------گزارش خلاصه فروش صندوق
ALTER     PROCEDURE [dbo].[GetStationSaleSummeryInfo]
    (
      --@intLanguage1 INT = Null ,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50),
      @user1 INT,
      @user2 INT,
      @Station1 INT,
      @Station2 INT,
      @Branch1 INT,
      @Branch2 INT       
    )
AS 
    DECLARE @intLanguage1 INT
--    SET @intLanguage1 = 0   
    IF @intLanguage1 IS NULL 
        SET @intLanguage1 = 0

    DECLARE @strTmp NVARCHAR(50)
    DECLARE @intTmp INT
    DECLARE @Time3 NVARCHAR(50)
    DECLARE @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @Time2

    IF @Date2 < @Date1 
        BEGIN
            SET @strTmp = @Date2
            SET @Date2 = @Date1
            SET @Date1 = @strTmp
        END

    IF @Time2 < @Time1 
        BEGIN 
		/*SET @strTmp   = @Time2
		SET @Time2   = @Time1
		SET @Time1 = @strTmp*/
            SET @Time3 = '00:00'
            SET @Time4 = '24:00'
        END

    IF @user2 < @user1 
        BEGIN
            SET @intTmp = @user2
            SET @user2 = @user1
            SET @user1 = @intTmp
        END
	
    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage1 = 0 
        SET @TimeTitle = N' ساعت : '
    ELSE 
        SET @TimeTitle = N'Time: '

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
    SELECT  SUM(VwStationSaleSummery.SumPrice) AS SumPriceTotal,
            SUM(VwStationSaleSummery.CarryFeeTotal) AS SumCarryFee,
            SUM(VwStationSaleSummery.PackingTotal) AS SumPacking,
            SUM(VwStationSaleSummery.DiscountTotal) AS SumDiscount,
            SUM(VwStationSaleSummery.ServiceTotal) AS SumService,
            SUM(VwStationSaleSummery.GarsonSumPrice) AS SumGarsonSale,
            SUM(VwStationSaleSummery.CarrierSumPrice) AS SumCarrierSale,
            SUM(VwStationSaleSummery.CarrierDebit) AS SumCarrierDebit,
            SUM(VwStationSaleSummery.GarsonDebit) AS SumGarsonDebit,
            SUM(VwStationSaleSummery.CustomerDebit) AS SumCustomerDebit,
            SUM(VwStationSaleSummery.UnBalanceFich) AS SumUnBalanceFich,
            SUM(VwStationSaleSummery.Payment) AS SumPayment,
            SUM(VwStationSaleSummery.Recieved) AS SumRecieved,
            VwStationSaleSummery.[Date] ,
	    SUM(VwStationSaleSummery.TaxTotal) AS SumTax ,
	    SUM(VwStationSaleSummery.DutyTotal) AS SumDuty,
	    SUM(VwStationSaleSummery.TipAmount) AS SumTip,
	    SUM(VwStationSaleSummery.RoundDiscount) AS sumRoundDiscount,
	    SUM(VwStationSaleSummery.CardReceived) AS SumCardReceived,
	    SUM(VwStationSaleSummery.ManualRecieved) AS SumManualRecieved,
	    SUM(VwStationSaleSummery.OrderPrice) AS OrderPrice ,
	    SUM(VwStationSaleSummery.OrderReceived) AS SumOrderReceived ,
	    SUM(VwStationSaleSummery.SumRecursive) AS SumRecursive ,
            
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @Time1 AS FromTime,
            @Time2 AS ToTime,
            @user1 AS FormUser,
            @user2 AS ToUser,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate --,
			,VwStationSaleSummery.Branch , @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2
			, dbo.tBranch.nvcBranchName , VwStationSaleSummery.[User] AS Uid , VwStationSaleSummery.UserFullName
    FROM    VwStationSaleSummery 
    	INNER JOIN dbo.tBranch ON dbo.VwStationSaleSummery.Branch = dbo.tBranch.Branch 
    	LEFT OUTER JOIN dbo.tStations ON dbo.VwStationSaleSummery.StationID = dbo.tStations.StationID --AND dbo.VwStationSaleSummery.Branch = dbo.tStations.Branch
    WHERE   VwStationSaleSummery.[Date] >= @Date1
            AND VwStationSaleSummery.[Date] <= @Date2
            AND ( ( VwStationSaleSummery.[Time] >= @Time1
                    AND VwStationSaleSummery.[Time] <= @Time4
                  )
                  OR ( VwStationSaleSummery.[Time] <= @Time2
                       AND VwStationSaleSummery.[Time] >= @Time3
                     )
                )
            AND ( ( VwStationSaleSummery.StationID >= @Station1
                    AND VwStationSaleSummery.StationID <= @Station2
                  )
                  OR VwStationSaleSummery.StationID = 0
                )
            AND ((VwStationSaleSummery.[User] >= @user1 AND VwStationSaleSummery.[User] <= @user2) ) -- OR (dbo.tStations.StationType = 8 AND VwStationSaleSummery.Balance = 0))
            AND VwStationSaleSummery.Branch >= @Branch1
            AND VwStationSaleSummery.Branch <= @Branch2
    GROUP BY VwStationSaleSummery.[Date] , VwStationSaleSummery.Branch , dbo.tBranch.nvcBranchName , VwStationSaleSummery.[User] , VwStationSaleSummery.UserFullName
 --,VwStationSaleSummery.[User],VwStationSaleSummery.StationID,
	--VwStationSaleSummery.UserGender,VwStationSaleSummery.UserFullName



GO



--Script_V26_16_Fix10
--اضافه شدن مرکز هزینه به فرم سود وزیان حسابداری
--تولید سند حسابداری یکپارچه از داخل اسکریپت
--اضافه شدن بدهی فروش به مشتریان (اگر سند تولید نشده)در تولید سند حسابداری
--اضافه شدن فیلد تفضیلی به پارتیشن ها برای محاسبه عوارض و مالیات و سایر افزایش ها

-- 93/10/19

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
          10
        )
GO
-- Date 930920

IF COL_LENGTH('tPartitions','Tafsili') IS NULL
BEGIN
	ALTER TABLE dbo.tPartitions
	ADD Tafsili INT NULL
END

GO
IF COL_LENGTH('tStations','PartitionId') IS NULL
BEGIN
	ALTER TABLE dbo.tStations
	ADD PartitionId INT NOT NULL DEFAULT(1)
END

GO


UPDATE dbo.tStations SET PartitionId = 1 WHERE PartitionId IS NULL 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_SaleSummaryCustom]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_SaleSummaryCustom]
GO

CREATE PROCEDURE [dbo].[Get_SaleSummaryCustom]
(
@Branch INT ,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT = 0
)

 AS
BEGIN

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت صندوق' + '  ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date] AS [Name] ,  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
 INNER JOIN dbo.tUser TU ON TU.UID = TF.[User] AND TU.Branch = TF.Branch
 INNER JOIN dbo.tPer TP ON TP.pPno = TU.pPno AND TP.Branch = TU.Branch  
 INNER JOIN dbo.tFacCash TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 7
GROUP BY TF.[User] , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت کارت' + N' بانک ' + MIN(TPP.nvcBankName) + N' شماره ' + MIN(TPP.NvcPosNo) + N' درتاریخ ' +  TF.[Date]  AS [Name],  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TPP.AccountId) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
 INNER JOIN dbo.tFacCard TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
INNER JOIN dbo.tblPub_Pos TPP on TPP.PosId = TFC.PosId
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 22
GROUP BY TFC.PosId , TF.[Date]


UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت پیک' + ' ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date]  AS [Name],  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)
     		        AND InCharge > 0 
     		        AND Balance = 0 and FacPayment = 0) TF
 INNER JOIN dbo.tPer TP ON TP.pPno = TF.InCharge AND TP.Branch = TF.Branch  and Job = 3
 INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 6
GROUP BY TP.pPno , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت گارسون' + ' ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date]  AS [Name],  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)
     		        AND InCharge > 0 
     		        AND Balance = 0 and FacPayment = 0) TF
 INNER JOIN dbo.tPer TP ON TP.pPno = TF.InCharge AND TP.Branch = TF.Branch  and Job = 9
 INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 6
GROUP BY TP.pPno , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' بدهکاری مشتری' + ' ' + MIN(TC.Name + '' + TC.Family) + N' در تاریخ ' + TF.[Date]  AS [Name] ,  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TC.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
					AND (InCharge IS NULL OR (InCharge > 0 AND FacPayment = 1)) 
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer > 0 AND Credit > 0)
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 7
GROUP BY TC.Code , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' تخفیفات فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , SUM(TF.DiscountTotal) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  2
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'  فروش ' + ' ' + MIN(Ts.[Description]) + N' در تاریخ  ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(Tf.Amount * Tf.FeeUnit) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TS.Tafsili) AS Tafsili FROM 
(SELECT tFacM.* , Amount , FeeUnit , intInventoryNo FROM dbo.tFacM INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tInventory TS ON TS.InventoryNo = TF.intInventoryNo
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  1
GROUP BY TS.InventoryNo , TF.[Date]

--SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' فروش ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice + TF.DiscountTotal - TF.CarryFeeTotal - PackingTotal - ServiceTotal - TaxTotal - DutyTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
--(SELECT * FROM dbo.tFacM
--                    where [Date] >= @DateBefore
--                    AND [Date] <= @DateAfter
--                    AND Recursive = 0
--                    AND Status = 2
--                    AND transferAccounting = 0
--                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
--INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
--INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
--INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  1
--GROUP BY TP.PartitionID , TF.[Date]

--UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' بستانکاری مشتری' + ' ' + MIN(TC.Name + '' + TC.Family) + N' در تاریخ ' + TF.[Date]  AS [Name] ,  0 AS SumBedehKar , SUM(ISNULL(TFC.intAmount,0)+ISNULL(TFCA.intAmount,0)) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TC.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 2
----                    AND transferAccounting = 0
----     		        AND (InCharge = NULL OR (InCharge > 0 AND FacPayment = 1)) 
----     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
----INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer > 0 AND Credit > 0)
----LEFT JOIN dbo.tFacCash TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
----LEFT JOIN dbo.tFacCard TFCA ON TFCA.Branch = TF.Branch AND TFCA.intSerialNo = TF.intSerialNo
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TC.Code , TF.[Date]
----HAVING SUM(ISNULL(TFC.intAmount,0)+ISNULL(TFCA.intAmount,0)) > 0

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'موجودي مواد و کالا' AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM INNER JOIN 
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 16
----GROUP BY TF.[Date]


----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , MIN(TS.Name + ' ' + TS.Family + ' ' + TS.WorkName) AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , TS.Code AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TS.Code , TF.[Date]


----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'برگشت از خريد' AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 4
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 17
----GROUP BY TF.[Date]

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , MIN(TS.Name + ' ' + TS.Family + ' ' + TS.WorkName) AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , TS.Code AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TS.Code , TF.[Date]

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'برگشت از فروش' AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 5
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 18
----GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' عوارض فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.DutyTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  24
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' مالیات فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(TF.TaxTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  26
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' درآمد سرویس   ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.ServiceTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  38
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'  درآمد بسته بندی   ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.PackingTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  3
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' کرایه حمل فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.CarryFeeTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  4
GROUP BY TF.[Date]


END

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER    PROCEDURE [dbo].[Get_SaleSummary]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
SELECT SUM(SumPrice)AS SumPriceTotal ,
        tvw.[Date] ,
        Branch ,
        Tafsili ,
        InventoryName

FROM 
(
SELECT DISTINCT dbo.tFacM.Branch  ,--NO ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacD.intRow ,
                    tfacd.Amount ,
                    tfacd.Feeunit ,
                    ( tfacd.Amount * tfacd.Feeunit ) AS SumPrice ,
                    dbo.tInventory.Tafsili ,
                    dbo.tInventory.Description AS InventoryName
          FROM      dbo.tFacM
                    INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                                        AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
					INNER JOIN dbo.tInventory ON dbo.tInventory.InventoryNo = dbo.tFacD.intInventoryNo
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND (dbo.tCust.Tafsili = 0 OR dbo.tCust.Tafsili IS NULL) ))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch , tvw.Tafsili , InventoryName
 ORDER BY tvw.[Date] 
 
 
END

GO



SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO


ALTER  Function [dbo].Fn_SoodZian

(
  @DateBefore INT  ,
  @DateAfter INT  ,
  @AccountYear SMALLINT ,
  @Branch INT ,
  @MarkazHazineh INT 
)

RETURNS  @ReturnTable TABLE(
 TotalSellAmount BIGINT ,
 TotalSellReturnAmount BIGINT ,
 TotalFirstPrice BIGINT ,
 TotalBuyAmount BIGINT ,
 TotalBuyReturnAmount BIGINT ,
 TotalSaleDiscount BIGINT ,
 TotalBuyDiscount BIGINT ,

 TotalCareeFee BIGINT ,
 TotalPacking BIGINT ,
 TotalService BIGINT ,

 TotalLosses BIGINT ,
 TotalHoghough BIGINT ,
 TotalHazine BIGINT ,
 TotalHazineMali BIGINT ,
 TotalHazineTozie BIGINT 
)	
As

BEGIN


	DECLARE @TotalSellAmount BIGINT
	DECLARE @TotalSellReturnAmount BIGINT
	DECLARE @TotalFirstPrice BIGINT
	DECLARE @TotalBuyAmount BIGINT
	DECLARE @TotalBuyReturnAmount BIGINT
	DECLARE @TotalSaleDiscount BIGINT
	DECLARE @TotalBuyDiscount BIGINT

	DECLARE @TotalCareeFee BIGINT
	DECLARE @TotalPacking BIGINT
	DECLARE @TotalService BIGINT

	DECLARE @TotalLosses BIGINT
	DECLARE @TotalHoghough BIGINT
	DECLARE @TotalHazine BIGINT
	DECLARE @TotalHazineMali BIGINT
	DECLARE @TotalHazineTozie BIGINT
	


		Select @TotalSellAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalSellReturnAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalFirstPrice = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalBuyAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalBuyReturnAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalSaleDiscount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalBuyDiscount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29)
		AND TafsiliId = @MarkazHazineh

		Select @TotalLosses = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHazine = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHazineMali = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 36  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalHazineTozie = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 37  )
		AND MoeinId <> (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32) --Losses  moein code calculated in totallosses
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalCareeFee = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4)
		AND TafsiliId = @MarkazHazineh
		
		Select @TotalPacking = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3)
		AND TafsiliId = @MarkazHazineh
		
		Select @Totalservice = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 38  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 38)
		AND TafsiliId = @MarkazHazineh
		
		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,  Totalservice ,
			 TotalLosses  , TotalHoghough  , TotalHazine , TotalHazineMali , TotalHazineTozie )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  , @Totalservice ,
			 @TotalLosses  , @TotalHoghough  , @TotalHazine , @TotalHazineMali , @TotalHazineTozie)
		            


RETURN 


End

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Get_TarazSoodZian]
    (
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT ,
      @MarkazHazineh INT 
    )
AS 


SELECT ISNULL(TotalSellAmount , 0) AS TotalSellAmount ,
       ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
       ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
       ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
       ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
       ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
       ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
       ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
       ISNULL(TotalPacking , 0) AS TotalPacking ,
       ISNULL(TotalService , 0) AS TotalService ,
       ISNULL(TotalLosses , 0) AS TotalLosses ,
       ISNULL(TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(TotalHazine , 0) AS TotalHazine ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
       
	FROM DBO.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch , @MarkazHazineh )
--===============================================


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Rep_TarazSoodZian]
    (
      @SystemDate NVARCHAR(20) ,
      @SystemDay NVARCHAR(20) ,
      @SystemTime NVARCHAR(20) ,
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT ,
      @MarkazHazineh INT ,
      @MojodiPrice BIGINT 
    )
AS 

    DECLARE @TimeTitle NVARCHAR(10)      
    SET @TimeTitle = N' ساعت : '   

SELECT @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS SysDay  ,
		SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,1 ,4) + N'/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,5,2) + N'/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,7,2) AS FromDate ,
		SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,1 ,4) + N'/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,5,2) + N'/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,7,2) AS ToDate ,
		@MojodiPrice AS MojodiPrice ,
		ISNULL(TotalSellAmount , 0) AS TotalSellAmount ,
		ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
		ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
		ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
		ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
		ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
		ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
		ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
		ISNULL(TotalPacking , 0) AS TotalPacking ,
		ISNULL(TotalService , 0) AS TotalService ,
		ISNULL(TotalLosses , 0) AS TotalLosses ,
		ISNULL(TotalHoghough , 0) AS TotalHoghough ,
		ISNULL(TotalHazine , 0) AS TotalHazine  ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
	FROM dbo.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch , @MarkazHazineh)
--===============================================

GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER    PROCEDURE [dbo].[Get_SaleSummary_Added]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
 
 SELECT 
        tvw.[Date] ,
        Branch ,
        SUM(DiscountTotal) AS Sumdiscount ,
        SUM(PackingTotal) AS SumPacking ,
        SUM(CarryFeeTotal) AS SumCarryFee ,
        SUM(ServiceTotal) AS SumService ,
        SUM(DutyTotal) AS DutyTotal ,
        SUM(TaxTotal) AS TaxTotal ,
        0 AS tafsili
 FROM   ( SELECT DISTINCT dbo.tFacM.Branch ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacM.[Time] ,
                    dbo.tFacM.[User] ,
                    CarryFeeTotal ,
                    DiscountTotal ,
                    StationID ,
                    ServiceTotal ,
                    PackingTotal ,
                    TaxTotal ,
                    DutyTotal ,
                    FacPayment ,
                    Balance ,
                    --( tfacd.Amount * tfacd.Feeunit ) AS SumPrice
                    dbo.tFacM.SumPrice
          FROM      dbo.tFacM
                    --INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                    --                    AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch
 ORDER BY tvw.[Date] 
 
 
END

GO




--این اسکریپت فقط یک رکورد برای اصلاحی ها با زمان و مبلغ اولی و آخری بر می گرداند
--Script_V26_16_Fix10_1_EditedFich
--93/10/21


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO

ALTER  PROCEDURE [dbo].[Get_EditedFactors_Print] (
@SystemDate  	NVARCHAR(20),
@SystemDay   	NVARCHAR(20),
@SystemTime  	NVARCHAR(20),
@DateAfter Nvarchar(20) , 
@DateBefore Nvarchar(20)

)
 AS

SELECT    
		      @DateBefore  AS DateBefore, @DateAfter AS DateAfter ,
	   	      @SystemDay + ' ' + @SystemDate +' '+N' ساعت : ' + @SystemTime AS Sysdate  ,
		      dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tFacM.Status, 
                      dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.OrderType, 
                      dbo.tFacM.ServePlace, dbo.tFacM.StationID,  
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName + ' ' +  dbo.tPer.nvcSurName As FullName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description ,
		      ISNULL(T.Time , 0) AS Time1 , ISNULL(T.SumPrice , 0) AS Price1 , ISNULL(T.intSerialNo ,0) AS MinCode  -- , dbo.tRepFacEditM.SumPrice As Price1
FROM         dbo.tFacM  INNER JOIN          
			 dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
             dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
             dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  and  dbo.tFacM.Branch = dbo.tShift.Branch
				LEFT OUTER JOIN (SELECT code , Branch , intSerialNo , SumPrice , Time FROM dbo.tRepFacEditM )T 
						ON T.Branch = dbo.tFacM.Branch AND T.intSerialNo = tFacM.intSerialNo AND T.Code = (Select Min(Code) FROM dbo.tRepFacEditM WHERE T.Branch = dbo.tRepFacEditM.Branch AND T.intSerialNo = dbo.tRepFacEditM.intSerialNo) 									
WHERE      ISNULL(T.intSerialNo ,0) > 0  AND
			 dbo.tFacm.[Date] >= @DateAfter And dbo.tFacm.[Date] <= @DateBefore 
			And dbo.tFacm.Status =2


order By dbo.tFacM.intSerialNo desc
GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  PROCEDURE [dbo].[InsertPersonel]( 
	@PersonnelNumber nvarchar(50),
	@nvcFirstName nvarchar(50),
	@nvcSurName nvarchar(50),
	@Gender bit,
	@IdNumber nvarchar(50),
	@Job int,
	@InsuranceNo nvarchar(50) ,
	@Address nvarchar(300),
	@Tel nvarchar(30),
	@User int , 
	@UserName nvarchar(50) ,
	@Password nvarchar(50) ,
	@intAccessLevel int ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno int out

	)
 AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

set @Time = dbo.SetTimeFormat(getdate())

select @pPno = isnull(max(Ppno),0) + 1 from tper --Where Branch = @Branch 
If @pPno < (@Branch * 1000 ) Set @pPno = (@Branch * 1000 )


begin Tran
insert into dbo.tper (
	pPno ,
	PersonnelNumber,
	nvcFirstName,
	nvcSurName,
	Gender ,
	IdNumber,
	Job ,
	InsuranceNo  ,
	Address ,
	Tel ,
	[Date] ,
	[Time] ,			
	[User] ,
	Branch,
	MaxCredit,
	ActDeAct

)
values(
	@pPno ,
	@PersonnelNumber,
	@nvcFirstName,
	@nvcSurName ,
	@Gender ,
	@IdNumber ,
	@Job ,
	@InsuranceNo ,
	@Address ,
	@Tel ,
	@Date,
	@Time ,
	@User ,
	@Branch,
	@MaxCredit,
	@ActDeAct
)
if @@Error <> 0 
		GOTO EventHandler	


--set @pPno=@@identity
DECLARE @UID INT
if @intAccessLevel<>0 and @UserName <> '' and @Password<>''

BEGIN

	select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch 
	If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )

	insert into dbo.tUser 
	(
		[Uid] ,
	 UserName ,
	 [Password] ,
	 intAccessLevel ,
	 pPno ,
	 addUser , 
	 Branch, 
	 CountRePrint, 
	 CountInvoicePrint,
	 CountInvoiceEditable,
	 CountInvoiceRefferable
	)
 values (
	@UID ,					
	@UserName  ,
	@Password  ,
	@intAccessLevel ,
	@pPno , 
	@User ,
	@Branch,
	@CountRePrint,
	@CountInvoicePrint,
	@CountInvoiceEditable,
	@CountInvoiceRefferable
	)

if @@Error <> 0 
		GOTO EventHandler	
--SET @UID = @@IDENTITY		
END	



commit Tran



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1




GO







--script V26_16_Fix9_GoodOptions
--اضافه کردن مبلغ آپشن ها به قیمت کالا
-- 93/09/12

UPDATE dbo.tObjects SET ObjectName = N'آپشن های کالاها' WHERE intObjectCode = 213
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_FacMD_Good] (@No Bigint , @Status int , @intLanguage int , @AccountYear Smallint  , @Branch INT ) 

AS
DECLARE @intSerialNo INT 
SELECT @intSerialNo = intSerialNo FROM tfacM where No = @No  And  Status = @Status And  AccountYear =  @AccountYear AND Branch = @Branch


Select Sum(vw_FacMD_Good.Amount)As Amount  , vw_FacMD_Good.GoodCode    , Max(vw_FacMD_Good.ServePlace) As Serveplace , Max(vw_FacMD_Good.DifferencesCodes) As DifferencesCodes  ,
	Max(vw_FacMD_Good.DifferencesDescription) As DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]  ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	MAX(vw_FacMD_Good.Rate) as rate ,vw_FacMD_Good.ChairName  , Max(vw_FacMD_Good.FeeUnit)As FeeUnit ,
	Max(vw_FacMD_Good.intinventoryNo)As intinventoryNo ,Max(vw_FacMD_Good.DestInventoryNo)As DestInventoryNo,Max(vw_FacMD_Good.[ExpireDate])As [ExpireDate] , Max( IsNull(vw_FacMD_Good.[NvcDescription], ''))As [NvcDescription]
        , case @intLanguage when 0 then Name 
			    when 1 then LatinName end as nvcName , vw_FacMD_Good.intRow
	,Max(vw_FacMD_Good.NumberOfUnit) As NumberOfUnit , vw_FacMD_Good.maintype,Max( IsNull(vw_FacMD_Good.[TempAddress], ''))As [TempAddress]
	, Max(vw_FacMD_Good.[Description]) AS UnitDescription , ISNULL(T.Mojodi , 0 ) AS Mojodi
	 , TaxBuy , TaxSale , DutyBuy , DutySale , ISNULL(DestinationId , 0) AS DestinationId
	 , (SELECT SUM(Amount) FROM dbo.tFacD WHERE intSerialNo = @intSerialNo AND Branch = @Branch) AS SumAmount

 	from vw_FacMD_Good 
	LEFT OUTER JOIN
	(SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
	AND t.Branch = @Branch AND t.GoodCode = vw_FacMD_Good.GoodCode AND t.InventoryNo = vw_FacMD_Good.intinventoryNo
 	 
	where No = @No  And  Status = @Status And  vw_FacMD_Good.AccountYear =  @AccountYear AND vw_FacMD_Good.Branch = @Branch
	 Group By vw_FacMD_Good.GoodCode     , vw_FacMD_Good.DifferencesCodes ,vw_FacMD_Good.DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]    ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	vw_FacMD_Good.ChairName   , vw_FacMD_Good.ServePlace , vw_FacMD_Good.FeeUnit  , vw_FacMD_Good.intRow , vw_FacMD_Good.MainType
	,T.Mojodi  , TaxBuy , TaxSale , DutyBuy , DutySale ,DestinationId
Order By  vw_FacMD_Good.intRow




GO

--exec Get_FacMD_Good 350, 2, 0, 1393, 1


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GoodLable]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GoodLable]
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE   PROCEDURE [dbo].[GoodLable]
    (
      @FichNo  INT  ,
      @StrGood  NVARCHAR(50) ,
      @StrDescription NVARCHAR(50) 
    )
AS 
    SELECT  
      @FichNo AS FichNo , 
      @StrGood AS StrGood , 
      @StrDescription AS StrDescription
    
--===============================================



GO

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO



ALTER    Function Split

(
    @nvcMainString nText
)

RETURNS  @ReturnTable TABLE(
	Row int IDENTITY (1, 1) NOT NULL ,
	Amount FLOAT , 
	GoodCode INT , 
	FeeUnit Float , 
	Discount Float ,
	Rate Int ,
	ChairName nvarchar(50),
	[ExpireDate] Int,
	intInventoryNo Int ,
	DestInventoryNo INT ,
	ServePlace INT , 
	DifferencesCode NVARCHAR(50) , 
	DifferencesDescription NVARCHAR(500))
	
As

BEGIN

IF @nvcMainString IS NOT  NULL
BEGIN
    DECLARE @intDelimiterPosField  INT
    DECLARE @intDelimiterPosRecord INT

    DECLARE @Amount FLOAT
    DECLARE @GoodCode INT
    DECLARE @FeeUnit Float
    Declare @Discount Float	
    Declare @Rate Int	
    DECLARE @ChairName  NVARCHAR(50)
    DECLARE @ExpireDate  INT 
    DECLARE @intInventoryNo INT
    DECLARE @DestInventoryNo INT
    DECLARE @ServePlace INT

    DECLARE @DifferencesCode NVARCHAR(50)
    DECLARE @DifferencesDescription NVARCHAR(500)

    DECLARE @TempDifference int 
    DECLARE @TempTable Table (nvcMainString nText)
    

    insert into @TempTable values (@nvcMainString)
   

    SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
    SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)

    WHILE @intDelimiterPosRecord <> 0
    BEGIN
--**********
        	SET @Amount = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS FLOAT)  from @TempTable )

        	SET @Amount =  ROUND(CAST(@Amount AS DECIMAL(15,3)),3)

	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @GoodCode = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )

        	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @FeeUnit = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Float)  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @Discount = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Float)  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @Rate = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Int)  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @ChairName = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS NVARCHAR(50))  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @ExpireDate = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )

        	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @intInventoryNo = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )

        	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )


-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @DestInventoryNo = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
	If @DestInventoryNo = 0 SET @DestInventoryNo = Null
             
	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )


-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
	

	SET @DifferencesCode = ''
	SET @DifferencesDescription = ' '

	IF @intDelimiterPosField < @intDelimiterPosRecord  and  @intDelimiterPosField > 0
		Begin
			SET @ServePlace = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS INT)  from @TempTable )
		
		        Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField ) from @TempTable )
			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
	
			--Set @DifferencesCode =( select  LTrim(RTrim(SUBSTRING(nvcMainString , 1 , @intDelimiterPosRecord - 1)))  from @TempTable )

			WHILE @intDelimiterPosField < @intDelimiterPosRecord  and  @intDelimiterPosField > 0
				BEGIN
					SET @TempDifference  = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS INT)  from @TempTable )
					SET @DifferencesCode = @DifferencesCode + ';' + CAST (@TempDifference AS nvarchar(50))
					SET @DifferencesDescription = @DifferencesDescription + ' , ' + (SELECT RTRIM(LTRIM([Difference])) FROM tDifferences WHERE Code = @TempDifference)
		        		
					Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
		
					SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
					SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)

				        	
				END
			SET @TempDifference = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1 , @intDelimiterPosRecord - 1))) AS INT)  from @TempTable )
			SET @DifferencesCode = @DifferencesCode + ';' + CAST (@TempDifference AS nvarchar(50))
			SET @DifferencesDescription = @DifferencesDescription + ' , ' + (SELECT RTRIM(LTRIM([Difference])) FROM tDifferences WHERE Code = @TempDifference)
		        
			Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosRecord + 1, DataLength(nvcMainString) - @intDelimiterPosRecord  ) from @TempTable )
			IF @DifferencesCode <> ''
				BEGIN
					Set @DifferencesCode = RIGHT (@DifferencesCode , LEN(@DifferencesCode) - 1)
					Set @DifferencesDescription = RIGHT (@DifferencesDescription , LEN(@DifferencesDescription) - 3)				
				End					
		END        
	ELSE		
		BEGIN
			SET @ServePlace = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString , 1 , @intDelimiterPosRecord - 1))) AS INT)  from @TempTable )
		
		      	Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosRecord + 1, DataLength(nvcMainString) - @intDelimiterPosRecord  ) from @TempTable )
	
		END

        INSERT INTO @ReturnTable(Amount , GoodCode , FeeUnit , Discount, Rate , ServePlace,ChairName ,[ExpireDate] , intInventoryNo ,DestInventoryNo ,  DifferencesCode ,DifferencesDescription) VALUES(@Amount, @GoodCode, @FeeUnit, @Discount , @Rate ,@ServePlace,@ChairName ,@ExpireDate ,@intInventoryNo , @DestInventoryNo , @DifferencesCode , @DifferencesDescription )
                
        SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable )
        SET @intDelimiterPosRecord = ( Select patindex('%/%' , nvcMainString)  from @TempTable )

    End

End

Return


End


GO






SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Delete_Inventory
(
	@InventoryNo	INT ,
	@Branch INT

)
AS

	DELETE FROM dbo.tblAcc_Tafsili WHERE TafsiliId = (SELECT Tafsili FROM tInventory WHERE dbo.tInventory.InventoryNo = @InventoryNo AND [Branch] = @Branch)
	DELETE from dbo.tInventory
	WHERE dbo.tInventory.InventoryNo = @InventoryNo AND [Branch] = @Branch


GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER PROCEDURE [dbo].[GetInventory_Branch] 
(@intLanguage int ,
@Branch int)
AS

 SELECT    Branch ,  InventoryNo, case @intLanguage  when 0 then  [Description]
					when 1 then IsNull(LatinDescription , ' ' )
		end as [Description] , Active , ISNULL(Tafsili ,0) AS Tafsili

 FROM         dbo.tInventory
 Where Branch =  @Branch


GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Insert_tinventory] (
					@Description nvarchar(50) , 
					@Active bit ,
					@Branch int , 
					@Account INT ,
					@InventoryNo int out )

AS

Begin Tran
set @InventoryNo=-1
Set @InventoryNo = (Select isnull(Max(InventoryNo) , 0) + 1 as InventoryNo from dbo.tinventory  
	WHERE    Branch  = @Branch )
IF @InventoryNo < @Branch * 100 SET @InventoryNo = @Branch * 100

declare @MasterCode int
select @MasterCode=InventoryNo from tinventory where branch=@Branch  and MasterCode is null 
--if  ( @MasterCode is null) or ( @MasterCode  is not null)
--	Goto ErrHandler

Insert Into dbo.tinventory
(InventoryNo , [Description] ,MasterCode,  Active , Branch)
values
( @InventoryNo , @Description ,@MasterCode,  @Active , @Branch)
 --set @InventoryNo=@@identity
if @@Error <> 0 
	Goto ErrHandler

IF @Account = 1
BEGIN 
	DECLARE @TafsiliId INT 
	SELECT @TafsiliId = ISNULL(MAX(TafsiliId) ,0) + 1 FROM dbo.tblAcc_Tafsili

	EXEC Insert_tblAcc_Tafsili @Branch ,@TafsiliId ,@Description , @Active , 4
	if @@Error <> 0 
		Goto ErrHandler
	UPDATE dbo.tInventory SET Tafsili = @TafsiliId WHERE InventoryNo = @InventoryNo And Branch = @Branch
END 

Commit Tran


Return

ErrHandler:
RollBack Tran
Set @InventoryNo = -1
Return




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Update_tinventory] (
					@Description nvarchar(50) , 
					@Active bit , 
					@Branch int ,
					@Account INT ,
					@InventoryNo int OUTPUT )

AS

Begin Tran


Update dbo.tinventory
set 	 [Description] = @Description , 
	Active = @Active
Where 	InventoryNo = @InventoryNo And Branch = @Branch

if @@Error <> 0 
	Goto ErrHandler
IF @Account = 1
BEGIN 
	
	DECLARE @TafsiliId INT 
	SELECT @TafsiliId = ISNULL(Tafsili, 0) FROM dbo.tInventory Where InventoryNo = @InventoryNo And Branch = @Branch
	PRINT @TafsiliId
	IF @TafsiliId = 0
		BEGIN 
		SELECT @TafsiliId = ISNULL(MAX(TafsiliId) ,0) + 1 FROM dbo.tblAcc_Tafsili
		PRINT @TafsiliId
		EXEC Insert_tblAcc_Tafsili @Branch ,@TafsiliId ,@Description , @Active , 4
		if @@Error <> 0 
			Goto ErrHandler
		
		UPDATE dbo.tInventory SET Tafsili = @TafsiliId WHERE InventoryNo = @InventoryNo And Branch = @Branch
		END 
END 

Commit Tran


Return

ErrHandler:
RollBack Tran
Set @InventoryNo = -1
Return


GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[GetInventory] 
(@intLanguage int ,
@Type int)
AS

	If @type = 0 
	Begin
		 SELECT     InventoryNo, case @intLanguage  when 0 then  [Description]
							when 1 then LatinDescription
				end as [Description] , ISNULL(Tafsili ,0) AS Tafsili
		
		 FROM         dbo.tInventory
		--Where Branch = dbo.Get_Current_Branch() 
	End
	Else If @type = 1 
	Begin
		 SELECT     InventoryNo, case @intLanguage  when 0 then  [Description]
							when 1 then LatinDescription
				end as [Description] , ISNULL(Tafsili ,0) AS Tafsili
		
		 FROM         dbo.tInventory
		  --Where MasterCode is  null OR  Branch =  dbo.Get_Current_Branch() 
	End




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    proc Get_Good_Code (@Code int , @intLanguage int , @StationId INT , @Flag Bit , @AccountYear Smallint)


as

DECLARE @Branch INT 
SET @Branch = (SELECT TOP 1 Branch FROM dbo.tStations WHERE StationID = @StationId )

DECLARE @GoodFirstCode INT
DECLARE @Mojodi AS INT

If @Flag = 1
BEGIN


SET @GoodFirstCode=(SELECT TOP 1 [GoodFirstCode] FROM [dbo].[tUsePercent]
			WHERE [GoodCode]=@Code
			AND  [GoodFirstCode] IN		
					(SELECT [Code] FROM [dbo].[tGood] WHERE [GoodType] = 4))
IF @GoodFirstCode IS NULL 
	BEGIN 
	SELECT @Mojodi = [Mojodi] FROM [dbo].[tInventory_Good]
	INNER JOIN dbo.tStation_Inventory_Good ON tInventory_Good.GoodCode = dbo.tStation_Inventory_Good.GoodCode
	AND tInventory_Good.InventoryNo = dbo.tStation_Inventory_Good.InventoryNo AND tInventory_Good.Branch = dbo.tStation_Inventory_Good.Branch
	AND dbo.tStation_Inventory_Good.StationID = @StationId 
	AND dbo.tStation_Inventory_Good.GoodCode = @Code 
	AND dbo.tStation_Inventory_Good.Branch = @Branch
	AND dbo.tStation_Inventory_Good.AccountYear = @AccountYear
	 
	WHERE tInventory_Good.[GoodCode]=@Code
--	AND [InventoryNo]=@InventoryNo
	AND tInventory_Good.[Branch]=@Branch
	AND tInventory_Good.[AccountYear]=@AccountYear
--	IF @Mojodi <= 0 AND  @GoodCode IN		
--					(SELECT [Code] FROM [dbo].[tGood] WHERE [GoodType] = 2) 
--		SET @Mojodi = 1 
--	SELECT @Mojodi AS Mojodi
	END 
ELSE
	BEGIN 
	SELECT @Mojodi = [Mojodi] FROM [dbo].[tInventory_Good]
	INNER JOIN dbo.tStation_Inventory_Good ON tInventory_Good.GoodCode = dbo.tStation_Inventory_Good.GoodCode
	AND tInventory_Good.InventoryNo = dbo.tStation_Inventory_Good.InventoryNo AND tInventory_Good.Branch = dbo.tStation_Inventory_Good.Branch
	AND dbo.tStation_Inventory_Good.StationID = @StationId 
	AND dbo.tStation_Inventory_Good.GoodCode = @GoodFirstCode 
	AND dbo.tStation_Inventory_Good.Branch = @Branch
	AND dbo.tStation_Inventory_Good.AccountYear = @AccountYear
	WHERE tInventory_Good.[GoodCode]=@GoodFirstCode
--	AND tInventory_Good.[InventoryNo]=@InventoryNo
	AND tInventory_Good.[Branch]=@Branch
	AND tInventory_Good.[AccountYear]=@AccountYear

	END 
  Select vw_Good.* , tInventory.InventoryNo , CASE @intLanguage WHEN 0 THEN  [Name]
		when 1 then LatinName
	end as [Name], ISNULL(@Mojodi , 0 ) AS Mojodi
	, ISNULL(Tafsili , 0) AS Tafsili
   FROM [dbo].[vw_Good]
   Inner Join tInventory_Level1 On tInventory_Level1.Level1 = vw_Good.Level1 
   Inner Join tInventory On tInventory_Level1.Branch = tInventory.Branch And  tInventory_Level1.InventoryNo = tInventory.InventoryNo
   Inner Join tStation_Inventory_Good On tStation_Inventory_Good.InventoryNo = tInventory.InventoryNo 
	And tStation_Inventory_Good.StationId = @StationId And tStation_Inventory_Good.Branch = @Branch
	And tStation_Inventory_Good.AccountYear = @AccountYear 
	And tStation_Inventory_Good.GoodCode = vw_Good.Code
   where vw_Good.Code = @Code And tStation_Inventory_Good.Active = 1
End
Else
Begin

  Select vw_Good.* ,  CASE @intLanguage WHEN 0 THEN  [Name]
		when 1 then LatinName
	end as [Name]  ,tInventory.InventoryNo , ISNULL(T.Mojodi , 0 ) AS Mojodi
	, ISNULL(Tafsili , 0) AS Tafsili
   FROM [dbo].[vw_Good]
   Inner Join tInventory_Level1 On tInventory_Level1.Level1 = vw_Good.Level1 
   Inner Join tInventory On tInventory_Level1.Branch = tInventory.Branch And  tInventory_Level1.InventoryNo = tInventory.InventoryNo
--   Inner Join tInventory On tInventoryType.Type = tInventory.Type  Or tInventory.Type = 1
   LEFT OUTER JOIN
   (SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
   AND t.Branch = @Branch AND t.GoodCode = vw_Good.Code AND t.InventoryNo = tInventory.InventoryNo

   where vw_Good.Code = @Code 


End





GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  proc Get_Good_Barcode (@Barcode nvarchar(50) , @StationId INT , @Flag Bit , @AccountYear Smallint )

AS

DECLARE @Code INT 
SELECT @Code = Code FROM dbo.tGood WHERE dbo.tGood.BarCode = @Barcode
IF @Code IS NULL RETURN 
DECLARE @GoodFirstCode INT
DECLARE @Mojodi AS INT
DECLARE @Branch INT 
SET @Branch = (SELECT TOP 1 Branch FROM dbo.tStations WHERE StationID = @StationId )

If @Flag = 1
BEGIN


SET @GoodFirstCode=(SELECT TOP 1 [GoodFirstCode] FROM [dbo].[tUsePercent]
			WHERE [GoodCode]=@Code
			AND  [GoodFirstCode] IN		
					(SELECT [Code] FROM [dbo].[tGood] WHERE [GoodType] = 4))
IF @GoodFirstCode IS NULL 
	BEGIN 
	SELECT @Mojodi = [Mojodi] FROM [dbo].[tInventory_Good]
	INNER JOIN dbo.tStation_Inventory_Good ON tInventory_Good.GoodCode = dbo.tStation_Inventory_Good.GoodCode
	AND tInventory_Good.InventoryNo = dbo.tStation_Inventory_Good.InventoryNo AND tInventory_Good.Branch = dbo.tStation_Inventory_Good.Branch
	AND dbo.tStation_Inventory_Good.StationID = @StationId 
	AND dbo.tStation_Inventory_Good.GoodCode = @Code 
	AND dbo.tStation_Inventory_Good.Branch = @Branch
	AND dbo.tStation_Inventory_Good.AccountYear = @AccountYear
	 
	WHERE tInventory_Good.[GoodCode]=@Code
--	AND [InventoryNo]=@InventoryNo
	AND tInventory_Good.[Branch]=@Branch
	AND tInventory_Good.[AccountYear]=@AccountYear
--	IF @Mojodi <= 0 AND  @GoodCode IN		
--					(SELECT [Code] FROM [dbo].[tGood] WHERE [GoodType] = 2) 
--		SET @Mojodi = 1 
--	SELECT @Mojodi AS Mojodi
	END 
ELSE
	BEGIN 
	SELECT @Mojodi = [Mojodi] FROM [dbo].[tInventory_Good]
	INNER JOIN dbo.tStation_Inventory_Good ON tInventory_Good.GoodCode = dbo.tStation_Inventory_Good.GoodCode
	AND tInventory_Good.InventoryNo = dbo.tStation_Inventory_Good.InventoryNo AND tInventory_Good.Branch = dbo.tStation_Inventory_Good.Branch
	AND dbo.tStation_Inventory_Good.StationID = @StationId 
	AND dbo.tStation_Inventory_Good.GoodCode = @GoodFirstCode 
	AND dbo.tStation_Inventory_Good.Branch = @Branch
	AND dbo.tStation_Inventory_Good.AccountYear = @AccountYear
	WHERE tInventory_Good.[GoodCode]=@GoodFirstCode
--	AND tInventory_Good.[InventoryNo]=@InventoryNo
	AND tInventory_Good.[Branch]=@Branch
	AND tInventory_Good.[AccountYear]=@AccountYear
	END 
 
  Select vw_Good.* , tInventory.InventoryNo , ISNULL(@Mojodi , 0 ) AS Mojodi
  , ISNULL(Tafsili ,0 ) AS Tafsili
   FROM [dbo].[vw_Good]
   Inner Join tInventory_Level1 On tInventory_Level1.Level1 = vw_Good.Level1 
   Inner Join tInventory On tInventory_Level1.Branch = tInventory.Branch And  tInventory_Level1.InventoryNo = tInventory.InventoryNo
   Inner Join tStation_Inventory_Good On tStation_Inventory_Good.InventoryNo = tInventory.InventoryNo 
	And tStation_Inventory_Good.StationId = @StationId And tStation_Inventory_Good.Branch = @Branch
	And tStation_Inventory_Good.AccountYear = @AccountYear 
	And tStation_Inventory_Good.GoodCode = vw_Good.Code
--   LEFT OUTER JOIN
--   (SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
--   AND t.Branch = dbo.Get_Current_Branch() AND t.GoodCode = vw_Good.Code AND t.InventoryNo = tInventory.InventoryNo
   where vw_Good.Code = @Code And tStation_Inventory_Good.Active = 1 
End

--If @Flag = 1
--Begin
--
--   Select vw_Good.* , tInventory.InventoryNo , ISNULL(T.Mojodi , 0 ) AS Mojodi  FROM [dbo].[vw_Good]
--   Inner Join tInventory_Level1 On tInventory_Level1.Level1 = vw_Good.Level1 
--   Inner Join tInventory On tInventory_Level1.Branch = tInventory.Branch And  tInventory_Level1.InventoryNo = tInventory.InventoryNo
--   Inner Join tStation_Inventory_Good On tStation_Inventory_Good.InventoryNo = tInventory.InventoryNo 
--	And tStation_Inventory_Good.StationId = @StationId 
--	And tStation_Inventory_Good.GoodCode = vw_Good.Code 
--	And tStation_Inventory_Good.Branch = dbo.Get_Current_Branch()
--	And tStation_Inventory_Good.AccountYear = @AccountYear 
--	LEFT OUTER JOIN
--	(SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
--	AND t.Branch = dbo.Get_Current_Branch() AND t.GoodCode = vw_Good.Code AND t.InventoryNo = tInventory.InventoryNo
--	AND vw_Good.BarCode = @Barcode
--   where BarCode = @Barcode  And Len(Barcode) > 0 And tStation_Inventory_Good.Active = 1
--End
Else
Begin

   Select vw_Good.* , tInventory.InventoryNo , ISNULL(T.Mojodi , 0 ) AS Mojodi 
   , ISNULL(Tafsili ,0 ) AS Tafsili
   FROM [dbo].[vw_Good]
   Inner Join tInventory_Level1 On tInventory_Level1.Level1 = vw_Good.Level1 
   Inner Join tInventory On tInventory_Level1.Branch = tInventory.Branch And  tInventory_Level1.InventoryNo = tInventory.InventoryNo
 --  Inner Join tInventory On tInventoryType.Type = tInventory.Type    Or tInventory.Type = 1
   LEFT OUTER JOIN
   (SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
   AND t.Branch = dbo.Get_Current_Branch() AND t.GoodCode = vw_Good.Code AND t.InventoryNo = tInventory.InventoryNo

   where BarCode = @Barcode   And Len(Barcode)  > 1 
End





GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   PROCEDURE [dbo].[Get_SaleSummary_Added]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
 
 SELECT 
        tvw.[Date] ,
        Branch ,
        SUM(DiscountTotal) AS Sumdiscount ,
        SUM(PackingTotal) AS SumPacking ,
        SUM(CarryFeeTotal) AS SumCarryFee ,
        SUM(ServiceTotal) AS SumService ,
        SUM(DutyTotal) AS DutyTotal ,
        SUM(TaxTotal) AS TaxTotal ,
        0 AS tafsili
 FROM   ( SELECT DISTINCT dbo.tFacM.Branch ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacM.[Time] ,
                    dbo.tFacM.[User] ,
                    CarryFeeTotal ,
                    DiscountTotal ,
                    StationID ,
                    ServiceTotal ,
                    PackingTotal ,
                    TaxTotal ,
                    DutyTotal ,
                    FacPayment ,
                    Balance ,
                    --( tfacd.Amount * tfacd.Feeunit ) AS SumPrice
                    dbo.tFacM.SumPrice
          FROM      dbo.tFacM
                    --INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                    --                    AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch
 ORDER BY tvw.[Date] 
 
 
END

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   PROCEDURE [dbo].[Get_SaleSummary]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
SELECT SUM(SumPrice)AS SumPriceTotal ,
        tvw.[Date] ,
        Branch ,
        Tafsili ,
        InventoryName

FROM 
(
SELECT DISTINCT dbo.tFacM.Branch  ,--NO ,
                    dbo.tFacM.[Date] ,
                    tfacd.Amount ,
                    tfacd.Feeunit ,
                    ( tfacd.Amount * tfacd.Feeunit ) AS SumPrice ,
                    dbo.tInventory.Tafsili ,
                    dbo.tInventory.Description AS InventoryName
          FROM      dbo.tFacM
                    INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                                        AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
					INNER JOIN dbo.tInventory ON dbo.tInventory.InventoryNo = dbo.tFacD.intInventoryNo
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND (dbo.tCust.Tafsili = 0 OR dbo.tCust.Tafsili IS NULL) ))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch , tvw.Tafsili , InventoryName
 ORDER BY tvw.[Date] 
 
 
END

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--exec Get_All_Factors 2, 1, 1391, 3, N'91/06/27', N'91/06/30'
--GO 



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tFacM_Description]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_tFacM_Description
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE  Proc Get_tFacM_Description
@Status INT ,
@AccountYear INT ,
@Branch INT ,
@nvcDescription Nvarchar(255)     
as    

Set @nvcDescription = Replace(  @nvcDescription  , N'ک' , N'ك' ) 
Set @nvcDescription = Replace(  @nvcDescription  , N'ي' , N'ی' )
--UPDATE tfacM SET NvcDescription = Replace(  @nvcDescription  ,N'ک' , N'ك' ), NvcDescription = Replace(  @nvcDescription  , N'ي' , N'ی' )

SELECT 		dbo.tFacM.intSerialNo, [No],tfacm.[Date],tfacm.[Time], SumPrice, isnull( NvcDescription ,N'') as NvcDescription ,
            dbo.tPer.nvcFirstName ,
            dbo.tPer.nvcSurName 
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno
WHERE tfacm.Status = @Status AND AccountYear = @AccountYear AND dbo.tFacM.Branch = @Branch
  AND CHARINDEX ( @nvcDescription , NvcDescription ) > 0 
Order By intSerialNo


GO





IF COL_LENGTH('tblTotal_PrintFich','nvcPrintDate') IS NULL
BEGIN
	ALTER TABLE dbo.tblTotal_PrintFich
	ADD nvcPrintDate DATETIME NULL 
END

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Update_tblTotal_printFich
	(
	 @intPrintFichNo INT ,
	 @TypeFlag BIT ,
	 @nvcError NVARCHAR(255)
	 )

AS


IF @TypeFlag = 0
	UPDATE [dbo].[tblTotal_PrintFich]  
		SET IsPrinted=1 , nvcError = @nvcError , nvcPrintDate = GETDATE()
	 	WHERE intPrintFichNo=@intPrintFichNo 

ELSE
	DELETE FROM  [dbo].[tblTotal_PrintFich]  
	 	WHERE intPrintFichNo=@intPrintFichNo 
	

Return 1



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

---------------------------------------------------------
-----------------گزارش فروش درصدی  و ساعتی


ALTER   PROCEDURE [dbo].[GetPercentInvoicePerHourInfo]
    (
      @intLanguage INT = 0,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50),
      @Branch1 INT,
      @Branch2 INT
    )
AS 
    DECLARE @tmp NVARCHAR(50)
    declare @Time3 NVARCHAR(50)
    declare @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @Time2
	

    IF @Time2 < @Time1 
        BEGIN
		/*SET @tmp        = @Time2
		SET @Time2  = @Time1
		SET @Time1 = @tmp*/
            SET @Time3 = '00'
            SET @Time4 = '24'
        END

    IF @Date2 < @Date1 
        BEGIN
            SET @tmp = @Date2
            SET @Date2 = @Date1
            SET @Date1 = @tmp
		
        END

    SET @Time1 = LTRIM(LEFT(@Time1, 2))
    SET @Time2 = LTRIM(LEFT(@Time2, 2))
    SET @Time3 = LTRIM(LEFT(@Time3, 2))
    SET @Time4 = LTRIM(LEFT(@Time4, 2))

    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage = 0 
        SET @TimeTitle = N' ساعت : '
    ELSE 
        SET @TimeTitle = N'Time: '

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
    SELECT  dbo.VwPercentInvoicePerHour.FactorCount,
            dbo.VwPercentInvoicePerHour.SalePriceTotal,
            dbo.VwPercentInvoicePerHour.[Date],
            dbo.VwPercentInvoicePerHour.FromTime,
            dbo.VwPercentInvoicePerHour.ToTime,
            dbo.VwPercentInvoicePerHour.Branch,
            CAST(( dbo.VwPercentInvoicePerHour.SalePriceTotal
                   / t.MySalePriceTotal ) * 100 AS DECIMAL(6, 3)) AS TotalPercent,
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @Time1 AS TimeBefore,
            @Time2 AS TimeAfter,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate
            , @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2 , Branch
    FROM    dbo.VwPercentInvoicePerHour
            INNER JOIN ( SELECT dbo.VwPercentInvoicePerHour.[Date] AS myDate,
                                SUM(dbo.VwPercentInvoicePerHour.SalePriceTotal) AS MySalePriceTotal
                         FROM   dbo.VwPercentInvoicePerHour
                         WHERE  dbo.VwPercentInvoicePerHour.[date] >= @Date1
                                AND dbo.VwPercentInvoicePerHour.[date] <= @Date2
                                AND ( ( dbo.VwPercentInvoicePerHour.FromTime >= @Time1
                                        AND dbo.VwPercentInvoicePerHour.ToTime <= @Time4
                                      )
                                      OR ( dbo.VwPercentInvoicePerHour.ToTime <= @Time2
                                           AND dbo.VwPercentInvoicePerHour.FromTime >= @Time3
                                         )
                                    )
                         GROUP BY dbo.VwPercentInvoicePerHour.[date]
                       ) t ON t.MyDate = dbo.VwPercentInvoicePerHour.[date]
    WHERE   dbo.VwPercentInvoicePerHour.[date] >= @Date1
            AND dbo.VwPercentInvoicePerHour.[date] <= @Date2
            AND ( ( dbo.VwPercentInvoicePerHour.FromTime >= @Time1
                    AND dbo.VwPercentInvoicePerHour.ToTime <= @Time4
                  )
                  OR ( dbo.VwPercentInvoicePerHour.ToTime <= @Time2
                       AND dbo.VwPercentInvoicePerHour.FromTime >= @Time3
                     )
                )
	--AND dbo.VwPercentInvoicePerHour.FromTime >= @Time1 
	--AND dbo.VwPercentInvoicePerHour.ToTime   <=  @Time2
            AND dbo.VwPercentInvoicePerHour.Branch >= @Branch1
            AND dbo.VwPercentInvoicePerHour.Branch <= @Branch2
    ORDER BY dbo.VwPercentInvoicePerHour.[Date] , dbo.VwPercentInvoicePerHour.FromTime


GO




------------------------------------------------------------------------------
--------------------------------------------------------------------------





--فلگ برای رسیدهای موقت
--رسید موقت فقط یکبار به رسید دائم تبدیل شود
--دسترسی برای دائمی کردن رسید موقت
--93/09/17


IF COL_LENGTH('tFacM','BitTempReceived') IS NULL
BEGIN
	ALTER TABLE tFacM
	ADD BitTempReceived BIT NULL 
END

GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 327 , -- intObjectCode - int
          N'frmSaveTempReceived' , -- ObjectId - nvarchar(50)
          N' دائمی کردن رسید موقت' , -- ObjectName - nvarchar(50)
          N'frmSaveTempReceived' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
GO
        
        
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          327  -- intObjectCode - int
          )
GO
        


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Get_All_Factors]
    (
      @Status INT ,
      @User INT ,
      @AccountYear SMALLINT ,
      @Branch INT,
	  @DateAfter Nvarchar(8) , 
      @DateBefore Nvarchar(8)
    )
AS 
    DECLARE @AccessLevel INT
    DECLARE @LastfacmNo INT
    DECLARE @Date NVARCHAR(50)
    SET @Date = dbo.[Get_ShamsiDate_For_Current_Shift](GETDATE())
    DECLARE @ShiftNo INT
    SET @ShiftNo = dbo.Get_Shift(dbo.SetTimeFormat(GETDATE()))
--    PRINT @ShiftNo

    SET @AccessLevel = ISNULL(( SELECT MIN(AccessLevel)
                                FROM    ( SELECT TOP 100 PERCENT
                                                    CASE WHEN [ObjectId] LIKE N'viewallstationsfactors'
                                                         THEN 1
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                         THEN 2
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                         THEN 3
                                                         ELSE 4
                                                    END AS AccessLevel
                                          FROM      dbo.tUser
                                                    INNER JOIN dbo.tAccess_Object ON dbo.tUser.intAccessLevel = dbo.tAccess_Object.intAccessLevel
                                                    INNER JOIN dbo.tObjects ON dbo.tAccess_Object.intObjectCode = dbo.tObjects.intObjectCode
                                          WHERE     --tObjects.ObjectId LIKE 'viewallstationsfactors' AND
                                                    UID = @User
                                                    --AND dbo.tUser.Branch = @Branch
                                                    AND ( [dbo].[tObjects].[ObjectId] LIKE N'viewallstationsfactors'
                                                    OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                    OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                        )
                                          ORDER BY  [dbo].[tObjects].[intObjectCode] DESC
                                        ) T1
                              ), 4)

    DECLARE @intAccessLevel INT
    SELECT  @intAccessLevel = intAccessLevel
    FROM    [dbo].[tUser]
    WHERE   uid = @User
            --AND [Branch] = @Branch
    IF @intAccessLevel = 1 
        SET @AccessLevel = @intAccessLevel


    SET @LastfacmNo = ( SELECT  ISNULL(MAX([NO]), 0) + 1
                        FROM    tFacM
                        WHERE   Status = @Status
                                AND Branch = @Branch
                                AND dbo.tFacM.AccountYear = @AccountYear
                      )   
--    IF @LastfacmNo < 1000 
--        SET @LastfacmNo = 0 
--    ELSE 
--        IF @LastfacmNo > 1000 
--            SET @LastfacmNo = @LastfacmNo - 1000

    SELECT  dbo.tFacM.intSerialNo, [No],tfacm.[Date],tfacm.[Time], SumPrice, Balance, Recursive, ServiceTotal, CarryFeeTotal, DiscountTotal,isnull( NvcDescription ,N'') as NvcDescription ,
            dbo.tPer.nvcFirstName ,
            dbo.tPer.nvcSurName ,
            ISNULL(tcust.WorkName + dbo.tCust.Family , '') + ISNULL(tSupplier.WorkName + dbo.tSupplier.Family , '') AS CustomerName
            , ISNULL(tfacm.GuestNo ,'') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.[No]) AS TempNo
            , tshift.Description AS ShiftDescription , ISNULL(tfacm.BitTempReceived , 0) AS BitTempReceived
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno
            INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code
            LEFT OUTER JOIN dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code 
            LEFT OUTER JOIN dbo.tSupplier ON dbo.tFacM.Owner = dbo.tSupplier.Code 
    WHERE   ( @AccessLevel = 1
              OR ( @AccessLevel = 2
                   AND [dbo].[tFacM].[Date] = @Date
                 )
              OR ( @AccessLevel = 3
                   AND [ShiftNo] = @ShiftNo
                   AND [dbo].[tFacM].[Date] = @Date
                 )
	      OR ( @AccessLevel = 4
	           AND dbo.tFacM.[ShiftNo] = @ShiftNo
	           AND dbo.tFacM.[Date] = @Date
	           AND dbo.tFacM.[User] = @User
	         )
            )
            AND dbo.tFacM.Status = @Status
           -- AND dbo.tFacM.[No] > @LastfacmNo
            AND dbo.tFacM.AccountYear = @AccountYear
			AND  dbo.tFacm.[Date] >= @DateAfter 
			And dbo.tFacm.[Date] <= @DateBefore
			AND dbo.tFacM.Branch = @Branch
    ORDER BY No DESC



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   PROCEDURE [dbo].[Get_Define_Factors]
    (
      @Status INT ,
      @User INT ,
      @No BIGINT ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS 
--    DECLARE @Branch INT
--    SELECT  @Branch = [dbo].[Get_Current_Branch]()
    DECLARE @AccessLevel INT
    DECLARE @ShiftNo INT
    DECLARE @Date AS NVARCHAR(10)
    SET @ShiftNo = dbo.Get_Shift(dbo.SetTimeFormat(GETDATE()))
    SET @Date = [dbo].[Get_ShamsiDate_For_Current_Shift](GETDATE())

    SET @AccessLevel = ISNULL(( SELECT MIN(AccessLevel)
                                FROM    ( SELECT TOP 100 PERCENT
                                                    CASE WHEN [ObjectId] LIKE N'viewallstationsfactors'
                                                         THEN 1
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                         THEN 2
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                         THEN 3
                                                         ELSE 4
                                                    END AS AccessLevel
                                          FROM      dbo.tUser
                                                    INNER JOIN dbo.tAccess_Object ON dbo.tUser.intAccessLevel = dbo.tAccess_Object.intAccessLevel
                                                    INNER JOIN dbo.tObjects ON dbo.tAccess_Object.intObjectCode = dbo.tObjects.intObjectCode
                                          WHERE     --tObjects.ObjectId LIKE 'viewallstationsfactors' AND
                                                    UID = @User
                                                    --AND dbo.tUser.Branch = @Branch
                                                    AND ( [dbo].[tObjects].[ObjectId] LIKE N'viewallstationsfactors'
                                                          OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                          OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                        )
                                          ORDER BY  [dbo].[tObjects].[intObjectCode] DESC
                                        ) T1
                              ), 4)
    
    DECLARE @intAccessLevel INT
    SELECT  @intAccessLevel = intAccessLevel
    FROM    [dbo].[tUser]
    WHERE   uid = @User
           -- AND [Branch] = @Branch
    IF @intAccessLevel = 1 
        SET @AccessLevel = @intAccessLevel

    SELECT  dbo.tFacM.* ,
            dbo.tPer.nvcFirstName ,
            dbo.tPer.nvcSurName ,
            ISNULL(tcust.WorkName + dbo.tCust.Family , '') + ISNULL(tSupplier.WorkName + dbo.tSupplier.Family , '') AS CustomerName
            ,  ISNULL(tfacm.GuestNo ,'') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.[No]) AS TempNo
            , tshift.Description AS ShiftDescription , ISNULL(tfacm.BitTempReceived ,0) AS BitTempReceived
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno
            INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code
            LEFT OUTER JOIN dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code 
            LEFT OUTER JOIN dbo.tSupplier ON dbo.tFacM.Owner = dbo.tSupplier.Code 
    WHERE   ( @AccessLevel = 1
              OR ( @AccessLevel = 2
                   AND [dbo].[tFacM].[Date] = @Date
                 )
              OR ( @AccessLevel = 3
                   AND [ShiftNo] = @ShiftNo
                   AND [dbo].[tFacM].[Date] = @Date
                 )
	      OR ( @AccessLevel = 4
	           AND dbo.tFacM.[ShiftNo] = @ShiftNo
	           AND dbo.tFacM.[Date] = @Date
	           AND dbo.tFacM.[User] = @User
	         )
            )
            AND dbo.tFacM.Status = @Status
            AND dbo.tFacM.[No] = @No
            AND AccountYear = @AccountYear
            AND dbo.tFacM.Branch = @Branch
--===============================================




GO



ALTER    VIEW dbo.vw_FacM_Per
AS
SELECT  dbo.tFacM.StationID,
		dbo.tFacM.RegDate, 
		ISNULL(dbo.tFacM.InCharge, 0) AS InCharge, 
		ISNULL(dbo.tFacM.TableNo, 0) AS TableNo, 
		dbo.tFacM.[Time], 
        dbo.tPer.nvcFirstName, 
		dbo.tPer.nvcSurName, 
		dbo.tFacM.[No], 
		dbo.tFacM.Status, 
		dbo.tFacM.[User], 
		dbo.tFacM.intSerialNo, 
        dbo.tShift.Description AS ShiftDescription, 
		dbo.tShift.Code AS ShiftNo, 
		dbo.tFacM.Balance, 
		dbo.tFacM.FacPayment, 
		dbo.tFacM.ServePlace , 
		dbo.tFacM.AccountYear
		, CASE DeliveryPer.job WHEN 3 THEN ISNULL(DeliveryPer.nvcFirstName,'-') +' '+ISNULL(DeliveryPer.nvcSurName,'-') ELSE N'--' END AS DeliveryFullName 
		,dbo.tFacM.Branch
		, dbo.tFacM.BitHavaleResid
		,dbo.tFacM.transferAccounting 
		, tfacm.BitLock , tfacm.GuestNo , tfacm.TempNo , Refrence_Acc , ISNULL(BitTempReceived ,0) AS BitTempReceived
FROM    dbo.tFacM 
		INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID 
							--AND dbo.tFacM.Branch = dbo.tUser.Branch 
		INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno 
							--AND dbo.tUser.Branch = dbo.tPer.Branch 
		INNER JOIN dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code 
							--AND dbo.tFacM.Branch = dbo.tShift.Branch
		LEFT OUTER JOIN dbo.tPer AS DeliveryPer ON tFacM.InCharge = DeliveryPer.pPno 
							--AND tFacM.Branch = DeliveryPer.Branch 
                      
--WHERE     (dbo.tFacM.Branch = dbo.Get_Current_Branch()) 


GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_BitTempReceived]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_BitTempReceived
GO


CREATE PROCEDURE [dbo].Update_BitTempReceived (
	@intSerialNo BIGINT  ,
	@Branch INT 
	)

AS

	UPDATE dbo.tFacM
		SET BitTempReceived = 1 WHERE intSerialNo = @intSerialNo AND Branch = @Branch
		
		
GO

--SELECT * FROM dbo.tFacM ORDER BY intSerialNo DESC


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO


ALTER  PROCEDURE Get_CurrentEditTime
@intserialNo INT ,
@Baranch INT 
 
AS

DECLARE @MinuteUseDiff INT
SELECT @MinuteUseDiff = 
( (CAST(SUBSTRING(dbo.shamsi(GETDATE()), 4, 2) AS INT) - 1 ) * 30
                + CAST(SUBSTRING(dbo.shamsi(GETDATE()), 7, 2) AS INT) ) * 1440
+ ( DATEPART(HOUR, GETDATE()) * 60 + DATEPART(minute, GETDATE()) ) 
 - ( (CAST(SUBSTRING(T.nvcDate, 4, 2) AS INT) - 1) * 30
                + CAST(SUBSTRING(T.nvcDate, 7, 2) AS INT) )  * 1440
           - ( CAST(SUBSTRING(T.nvcTime, 1, 2) AS INT) * 60
                + CAST(SUBSTRING(T.nvcTime, 4, 2) AS INT) ) 
from 
(
SELECT TOP 1 ISNULL(dbo.tRepFacEditM.Time , dbo.tFacM.Time) AS nvcTime , 
	         ISNULL(dbo.tRepFacEditM.RegDate , dbo.tFacM.RegDate) AS nvcDate 
 FROM dbo.tFacM LEFT OUTER JOIN dbo.tRepFacEditM ON dbo.tRepFacEditM.Branch = dbo.tFacM.Branch AND dbo.tRepFacEditM.intSerialNo = dbo.tFacM.intSerialNo
WHERE dbo.tFacM.intSerialNo = @intserialNo AND dbo.tFacM.Branch = @Baranch	
	And Code = (Select MIN(Code) from tRepFacEditM where intSerialNo = @intserialNo AND Branch = @Baranch)
) T

--SET @MinuteUseDiff = 25
SELECT @MinuteUseDiff AS  MinuteUseDiff,
 CASE WHEN @MinuteUseDiff < 0 THEN 0 WHEN @MinuteUseDiff > dbo.[Get_UserEditTime]() THEN 0 ELSE 1 END AS UserDiffTme ,
 CASE WHEN @MinuteUseDiff < 0 THEN 0 WHEN @MinuteUseDiff > dbo.[Get_ManagerEditTime]() THEN 0 ELSE 1 END AS ManagerDiffTme


GO


--امکان جابجایی پرسنل در شعبات
--93/10/01


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tPer]') and OBJECTPROPERTY(id, N'IsPrimaryKey') = 1)
ALTER TABLE [dbo].[tPer] DROP CONSTRAINT [PK_tPer]
GO

ALTER TABLE [dbo].[tPer] ADD 
	CONSTRAINT [PK_tPer] PRIMARY KEY  NONCLUSTERED 
	(
		[pPno]
	)  ON [PRIMARY] 
GO



ALTER   PROCEDURE [dbo].[UpdatePersonel]( 
	@CurrentPPNO 		INT,
	@PersonnelNumber 	NVARCHAR(50),
	@nvcFirstName 		NVARCHAR(50),
	@nvcSurName	 	NVARCHAR(50),
	@Gender 		BIT,
	@IdNumber 		NVARCHAR(50),
	@Job 			INT,
	@InsuranceNo 		NVARCHAR(50) ,
	@Address 		NVARCHAR(300),
	@Tel 			NVARCHAR(30),
	@User 			INT , 
	@UID 			INT ,
	@UserName 		NVARCHAR(50) ,
	@Password 		NVARCHAR(50) ,
	@intAccessLevel 	INT ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno 			INT OUT
	       )
AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

SET @Time= dbo.SetTimeFormat(getdate())

BEGIN TRANSACTION

	UPDATE tPer
		SET PersonnelNumber 	= @PersonnelNumber,
		    nvcFirstName    	= @nvcFirstName,
		    nvcSurName	    	= @nvcSurName,
		    Gender	    	= @Gender,
		    IdNumber       	= @IdNumber,
		    Job		    	= @Job,
		    InsuranceNo     	= @InsuranceNo,
		    Address	    	= @Address,
		    Tel   	    	= @Tel,
		    [Date]	    	= @Date,
		    [Time]	    	= @Time,
		    [User]	    	= @User,
		    MaxCredit		=@MaxCredit,
		    ActDeAct 		=@ActDeAct ,
		    Branch			= @Branch
	WHERE       pPNO = @CurrentPPNO  


	IF @@ERROR <> 0 
		GOTO EventHandler	

	set @pPno = @CurrentPPNO

	IF @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID<>0
		UPDATE tUser
			SET 		UserName       	= @UserName,
	        	   		 	[Password]     	= @Password,
			    		intAccessLevel 	= @intAccessLevel,
			    		pPno           	= @pPno,
			    		addUser        	= @User,
					 CountRePrint		=@CountRePrint,
		  			 CountInvoicePrint	=@CountInvoicePrint,
					 CountInvoiceEditable		=@CountInvoiceEditable,
		  			 CountInvoiceRefferable	=@CountInvoiceRefferable ,
		  			 Branch					= @Branch
			WHERE   UID = @UID    
	else 
		if @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID=0
		BEGIN 
			select @Uid = isnull(max(Uid),0) + 1 from tUser Where Branch = @Branch   
			If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )  
			insert into dbo.tUser (
						UID ,
						UserName , 
						[Password] , 
						intAccessLevel , 
						pPno , 
						addUser , 
						Branch,
						CountRePrint,
						CountInvoicePrint,
						CountInvoiceEditable,
		  			 	CountInvoiceRefferable
			) values (	
						@UID ,				
						@UserName  ,
						@Password  ,
						@intAccessLevel ,
						@pPno , 
						@User , 
						@Branch,
						@CountRePrint,
						@CountInvoicePrint,
						@CountInvoiceEditable,
		  			 	@CountInvoiceRefferable
			)			
			END 
	IF @@ERROR <> 0 
		GOTO EventHandler	



COMMIT TRANSACTION



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1



GO




--این اسکریپت فقط یک رکورد برای اصلاحی ها با زمان و مبلغ اولی و آخری بر می گرداند
--Script_V26_16_Fix10_1_EditedFich
--93/10/21


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO

ALTER  PROCEDURE [dbo].[Get_EditedFactors_Print] (
@SystemDate  	NVARCHAR(20),
@SystemDay   	NVARCHAR(20),
@SystemTime  	NVARCHAR(20),
@DateAfter Nvarchar(20) , 
@DateBefore Nvarchar(20)

)
 AS

SELECT    
		      @DateBefore  AS DateBefore, @DateAfter AS DateAfter ,
	   	      @SystemDay + ' ' + @SystemDate +' '+N' ساعت : ' + @SystemTime AS Sysdate  ,
		      dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tFacM.Status, 
                      dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.OrderType, 
                      dbo.tFacM.ServePlace, dbo.tFacM.StationID,  
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName + ' ' +  dbo.tPer.nvcSurName As FullName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description ,
		      ISNULL(T.Time , 0) AS Time1 , ISNULL(T.SumPrice , 0) AS Price1 , ISNULL(T.intSerialNo ,0) AS MinCode  -- , dbo.tRepFacEditM.SumPrice As Price1
FROM         dbo.tFacM  INNER JOIN          
			 dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
             dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
             dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  and  dbo.tFacM.Branch = dbo.tShift.Branch
				LEFT OUTER JOIN (SELECT code , Branch , intSerialNo , SumPrice , Time FROM dbo.tRepFacEditM )T 
						ON T.Branch = dbo.tFacM.Branch AND T.intSerialNo = tFacM.intSerialNo AND T.Code = (Select Min(Code) FROM dbo.tRepFacEditM WHERE T.Branch = dbo.tRepFacEditM.Branch AND T.intSerialNo = dbo.tRepFacEditM.intSerialNo) 									
WHERE      ISNULL(T.intSerialNo ,0) > 0  AND
			 dbo.tFacm.[Date] >= @DateAfter And dbo.tFacm.[Date] <= @DateBefore 
			And dbo.tFacm.Status =2


order By dbo.tFacM.intSerialNo desc
GO





SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE dbo.Insert_CustomerFast


	@MembershipId BIGINT  ,     
	@Name nVarChar(50),   
	@Family nVarChar(50),    
	@Address nvarchar(150),  
	@Tel1 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Description nVarChar(200), 
	@User int ,   
	@Code Bigint out 

as  

 BEGIN TRAN  

	declare @MasterCode int
	set @MasterCode=null

	declare @Owner int    
	set @Owner=0

	declare @Sex int
	set @Sex=1 

	declare @WorkName nVarChar(50)
	set @WorkName=N''

	declare @InternalNo nVarChar(50)
	set @InternalNo=N''
 
	declare @Unit nVarChar(50)
	set @Unit=N''
  
	declare @City int
	set @City=1
	
	declare @ActKind int
	set @ActKind=1
  
	declare @ActDeAct bit
	set @ActDeAct=1
	
	declare @Prefix int
	set @Prefix=1
	  
	declare @Assansor bit   
	set @Assansor=0

	declare @PostalCode nVarChar(50)   
	set @PostalCode=N''

	declare @Tel2 nVarChar(50) 
	set @Tel2=N'' 
	declare @Tel3 nVarChar(50) 
	set @Tel3=N'' 
	declare @Tel4 nVarChar(50) 
	set @Tel4=N'' 

	declare @Fax nVarChar(50)
	set @Fax=N''  

	declare @Email nVarChar(50)
	set @Email=N''
  
	declare @Flour nVarChar(50)
	set @Flour=N''

	declare @CarryFee Float
	set  @CarryFee=0
  
	declare @PaykFee Float
	set  @PaykFee=0
  
	declare @Distance int
	set @Distance=1
   
	declare @Credit Float
	set @Credit=0
   
	declare @Discount Float
	set  @Discount=0
 
	declare @BuyState int
	set @BuyState=15
   
	declare @FamilyNo int 
	set  @FamilyNo=0 
	declare @Member Bit 
	set @Member=1
  
	declare @State int 
	set @State=1
  
	declare @Central BIT 
	set   @Central=1
	
	declare @Sellprice smallint
	set   @Sellprice=1

	declare @EconomicCode NVARCHAR(20) 
	set @EconomicCode=N''

	declare @nvcRFID NVARCHAR(20)
	set @nvcRFID=N''

	declare @nvcBirthDate NVARCHAR(10)
	set @nvcBirthDate=N''


Declare @Branch Int  
Set @Branch = dbo.Get_Current_Branch()  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  
 
Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  ( Branch = @Branch  Or Branch Is NULL ) )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

insert Into dbo.tCust  
(   
	Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Assansor,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	CarryFee,   
	PaykFee,   
	Distance,   
	Credit,   
	Discount,   
	BuyState,   
	[Description],   
	[Date],   
	[Time],   
	[User],  
	FamilyNo ,  
	Member ,  
	State ,  
	Central ,  
	Branch,  
	nvcRFID,  
	sellprice ,
	EconomicCode ,
	nvcBirthDate
	
)  
values  
(   
	@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName ,   
	@InternalNo,   
	@Unit,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Assansor,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@CarryFee,   
	@PaykFee,   
	@Distance,   
	@Credit,   
	@Discount,   
	@BuyState,   
	@Description,   
	@Date,   
	@Time,   
	@User ,  
	@FamilyNo ,  
	@Member ,  
	@State ,  
	@Central ,  
	@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate
	
)  

if @@Error <> 0   
 goto ErrHandler  

--SET @Code = @@IDENTITY
--SET @Code = 200


Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code


GO





--اضافه شدن دسته بندی به آپشن های کالاها
--برای یوگوبری و مشابه
--Script_V26_16_Fix10_4_OptionCategory
--93/10/23


IF COL_LENGTH('tDifferences','CategoryType') IS NULL
BEGIN
	ALTER TABLE dbo.tDifferences
	ADD CategoryType INT NULL
END

GO





SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    Procedure dbo.Insert_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int, 
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Code Bigint out 

)  

as  

Begin Tran  

--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  
if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  )  --AND (Branch = @Branch )
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode   )--AND (Branch = @Branch )
 end   
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId ) <> 0 --AND Branch = @Branch)   
  Goto ErrHandler   

Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  Branch = @Branch )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

if @nvcRFID = N''  
  SET @nvcRFID=N'-999'  

insert Into dbo.tCust  
(   
	Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Assansor,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	CarryFee,   
	PaykFee,   
	Distance,   
	Credit,   
	Discount,   
	BuyState,   
	[Description],   
	[Date],   
	[Time],   
	[User],  
	FamilyNo ,  
	Member ,  
	State ,  
	Central ,  
	Branch,  
	nvcRFID,  
	sellprice ,
	EconomicCode ,
	nvcBirthDate ,
	TotalRemainingAmount
	
)  
values  
(   
	@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Assansor,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@CarryFee,   
	@PaykFee,   
	@Distance,   
	@Credit,   
	@Discount,   
	@BuyState,   
	@Description,   
	@Date,   
	@Time,   
	@User ,  
	@FamilyNo ,  
	@Member ,  
	@State ,  
	@Central ,  
	@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate ,
	@TotalRemainingAmount
	
)  
if @@Error <> 0   
 goto ErrHandler  

--Set @Code = @@Identity  
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
  and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address)  
 , nvcRFID=CAST(Branch AS NVARCHAR(1))+CAST(Code AS NVARCHAR(8))  
  where code=@code  AND Branch = @Branch 


Update [tCust]
Set [Name] = Replace(  [Name] , N'ك' , N'ک'  ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي' , N'ی' ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ي' , N'ی' ) 



Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO



--Script_V26_16_Fix10_6_UsedGood
--چاپ لیبل خاص برای کالاها
--ساخت جدولی که عناصر تشکیل دهنده کالا می باشد
--tGood_Used  نام جدول جدید
-- ممکن است یک کالا از چند کالای دیگر تشکیل شده باشد
-- و در اینصورت این اقلام مصرفی در پرینت لیبل شرکت می کنند
--برای جلوگیری از اشتباه از دستکاری جدول ضریب مصرف پرهیز می گردد
--درسایر تنظیمات استفاده از جدول مصرف برای لیبل تیک بخورد
-- اگر کالایی در جدول مصرف تعیین نشود همان روال قبلی پابرجاست و فقط آن کالا لیبلش چاپ می شود
--93/10/26


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tGood_tGood_Used]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tGood_Used] DROP CONSTRAINT FK_tGood_tGood_Used
GO

/****** Object:  Table [dbo].[tGood]    Script Date: 01/14/2015 10:27:55 ******/

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tGood_Used]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tGood_Used]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE TABLE [dbo].[tGood_Used] (
	[GoodCode] [int] NOT NULL ,
	[GoodFirstCode] [int] NOT NULL ,
	[Auto_Id] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tGood_Used] ON [dbo].[tGood_Used]([GoodCode]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tGood_Used] ADD 
	CONSTRAINT [FK_tGood_tGood_Used] FOREIGN KEY 
	(
		[GoodCode]
	) REFERENCES [dbo].[tGood] (
		[Code]
	) ON UPDATE CASCADE 
GO


--INSERT INTO dbo.tGood_Used
--        ( GoodCode , GoodFirstCode )
--SELECT Code , Code FROM dbo.tGood
--GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Good_Used]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Good_Used]
GO

CREATE PROCEDURE Get_Good_Used
@GoodCode INT 

AS

SELECT RTRIM(lTrim(tGood.Name)) AS nvcName, (SELECT COUNT(*) FROM dbo.tGood_Used WHERE GoodCode = @GoodCode)  AS UsedCount  FROM dbo.tGood_Used INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tGood_Used.GoodFirstCode
WHERE GoodCode = @GoodCode

GO


ALTER  PROCEDURE [dbo].[Get_FacMD_Good] (@No Bigint , @Status int , @intLanguage int , @AccountYear Smallint  , @Branch INT ) 

AS
DECLARE @intSerialNo INT 
SELECT @intSerialNo = intSerialNo FROM tfacM where No = @No  And  Status = @Status And  AccountYear =  @AccountYear AND Branch = @Branch


Select Sum(vw_FacMD_Good.Amount)As Amount  , vw_FacMD_Good.GoodCode    , Max(vw_FacMD_Good.ServePlace) As Serveplace , Max(vw_FacMD_Good.DifferencesCodes) As DifferencesCodes  ,
	Max(vw_FacMD_Good.DifferencesDescription) As DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]  ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	MAX(vw_FacMD_Good.Rate) as rate ,vw_FacMD_Good.ChairName  , Max(vw_FacMD_Good.FeeUnit)As FeeUnit ,
	Max(vw_FacMD_Good.intinventoryNo)As intinventoryNo ,Max(vw_FacMD_Good.DestInventoryNo)As DestInventoryNo,Max(vw_FacMD_Good.[ExpireDate])As [ExpireDate] , Max( IsNull(vw_FacMD_Good.[NvcDescription], ''))As [NvcDescription]
        , case @intLanguage when 0 then Name 
			    when 1 then LatinName end as nvcName , vw_FacMD_Good.intRow
	,Max(vw_FacMD_Good.NumberOfUnit) As NumberOfUnit , vw_FacMD_Good.maintype,Max( IsNull(vw_FacMD_Good.[TempAddress], ''))As [TempAddress]
	, Max(vw_FacMD_Good.[Description]) AS UnitDescription , ISNULL(T.Mojodi , 0 ) AS Mojodi
	 , TaxBuy , TaxSale , DutyBuy , DutySale , ISNULL(DestinationId , 0) AS DestinationId
	 , (SELECT SUM(Amount) FROM dbo.tFacD WHERE intSerialNo = @intSerialNo AND Branch = @Branch) AS SumAmount

 	from vw_FacMD_Good 
	LEFT OUTER JOIN
	(SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
	AND t.Branch = @Branch AND t.GoodCode = vw_FacMD_Good.GoodCode AND t.InventoryNo = vw_FacMD_Good.intinventoryNo
 	 
	where No = @No  And  Status = @Status And  vw_FacMD_Good.AccountYear =  @AccountYear AND vw_FacMD_Good.Branch = @Branch
	 Group By vw_FacMD_Good.GoodCode     , vw_FacMD_Good.DifferencesCodes ,vw_FacMD_Good.DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]    ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	vw_FacMD_Good.ChairName   , vw_FacMD_Good.ServePlace , vw_FacMD_Good.FeeUnit  , vw_FacMD_Good.intRow , vw_FacMD_Good.MainType
	,T.Mojodi  , TaxBuy , TaxSale , DutyBuy , DutySale , DestinationId
Order By  vw_FacMD_Good.intRow




GO




ALTER    PROCEDURE [dbo].[UpdatePersonel]( 
	@CurrentPPNO 		INT,
	@PersonnelNumber 	NVARCHAR(50),
	@nvcFirstName 		NVARCHAR(50),
	@nvcSurName	 	NVARCHAR(50),
	@Gender 		BIT,
	@IdNumber 		NVARCHAR(50),
	@Job 			INT,
	@InsuranceNo 		NVARCHAR(50) ,
	@Address 		NVARCHAR(300),
	@Tel 			NVARCHAR(30),
	@User 			INT , 
	@UID 			INT ,
	@UserName 		NVARCHAR(50) ,
	@Password 		NVARCHAR(50) ,
	@intAccessLevel 	INT ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno 			INT OUT
	       )
AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

SET @Time= dbo.SetTimeFormat(getdate())

BEGIN TRANSACTION

	UPDATE tPer
		SET PersonnelNumber 	= @PersonnelNumber,
		    nvcFirstName    	= @nvcFirstName,
		    nvcSurName	    	= @nvcSurName,
		    Gender	    	= @Gender,
		    IdNumber       	= @IdNumber,
		    Job		    	= @Job,
		    InsuranceNo     	= @InsuranceNo,
		    Address	    	= @Address,
		    Tel   	    	= @Tel,
		    [Date]	    	= @Date,
		    [Time]	    	= @Time,
		    [User]	    	= @User,
		    MaxCredit		=@MaxCredit,
		    ActDeAct 		=@ActDeAct ,
		    Branch			= @Branch
	WHERE       pPNO = @CurrentPPNO  


	IF @@ERROR <> 0 
		GOTO EventHandler	

	set @pPno = @CurrentPPNO

	IF @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID<>0
		UPDATE tUser
			SET 		UserName       	= @UserName,
	        	   		 	[Password]     	= @Password,
			    		intAccessLevel 	= @intAccessLevel,
			    		pPno           	= @pPno,
			    		addUser        	= @User,
					 CountRePrint		=@CountRePrint,
		  			 CountInvoicePrint	=@CountInvoicePrint,
					 CountInvoiceEditable		=@CountInvoiceEditable,
		  			 CountInvoiceRefferable	=@CountInvoiceRefferable ,
		  			 Branch					= @Branch
			WHERE   UID = @UID    
	else 
		if @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID=0
		BEGIN 
			select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch   
			If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )  
			insert into dbo.tUser (
						UID ,
						UserName , 
						[Password] , 
						intAccessLevel , 
						pPno , 
						addUser , 
						Branch,
						CountRePrint,
						CountInvoicePrint,
						CountInvoiceEditable,
		  			 	CountInvoiceRefferable
			) values (	
						@UID ,				
						@UserName  ,
						@Password  ,
						@intAccessLevel ,
						@pPno , 
						@User , 
						@Branch,
						@CountRePrint,
						@CountInvoicePrint,
						@CountInvoiceEditable,
		  			 	@CountInvoiceRefferable
			)			
			END 
	IF @@ERROR <> 0 
		GOTO EventHandler	



COMMIT TRANSACTION



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Do_SaveInDetailsFactorReceived](@intSerialNo int, @ds nvarchar(4000)  , @Branch INT = NULL , @Remain INT = NULL , @Mode INT = NULL , @Result INT OUT  ) 
AS
BEGIN TRANSACTION

SET @Result = -1
If @Branch IS NULL
	SET @Branch = dbo.Get_Current_Branch()
IF @Remain IS NULL 
	SET @Remain = 0
DECLARE @SumPrice  float
	SET @SumPrice = (Select Sumprice From tFacm Where intserialno = @intSerialNo and Branch = @Branch )
IF @Mode IS NULL SET @Mode = 2
IF @Mode = 2
	DELETE FROM tFacCash WHERE intSerialNo = @intSerialNo And Branch = @Branch

INSERT INTO tFacCash
                      (Branch , intSerialNo, intAmount)
SELECT @Branch , @intSerialNo AS intSerialNo, c8 
FROM dbo.SplitFactorReceived(@ds)
WHERE c1 = 1
     IF @@ERROR <>0      
        GoTo EventHandler       

IF @Mode = 2
	DELETE FROM tFacCheque WHERE intSerialNo = @intSerialNo And Branch = @Branch
INSERT INTO tFacCheque
                      (Branch ,intSerialNo, intChequeSerial, intChequeAcc, intChequeDate, tintBank, nvcBranch, intChequeAmount)
SELECT @Branch ,@intSerialNo AS intSerialNo, c2, c3, c4, c5, c6, c8
FROM dbo.SplitFactorReceived(@ds)
WHERE c1 = 2
     IF @@ERROR <>0      
        GoTo EventHandler       

IF @Mode = 2
	DELETE FROM tFacCredit WHERE intSerialNo = @intSerialNo And Branch = @Branch
INSERT INTO tFacCredit
                      (Branch ,intSerialNo, intCreditSerial, intAmount)
SELECT @Branch ,@intSerialNo AS intSerialNo, c2, c8
FROM dbo.SplitFactorReceived(@ds)
WHERE c1 = 3
     IF @@ERROR <>0      
        GoTo EventHandler       

IF @Mode = 2
	DELETE FROM tFacLoan WHERE intSerialNo = @intSerialNo And Branch = @Branch
INSERT INTO tFacLoan
                      (Branch ,intSerialNo, intLoanDate, tintCount, intAmount)
SELECT @Branch ,@intSerialNo AS intSerialNo, c4, c7, c8
FROM dbo.SplitFactorReceived(@ds)
WHERE c1 = 4
     IF @@ERROR <>0      
        GoTo EventHandler       


IF @Mode = 2
	DELETE FROM tFacCard WHERE intSerialNo = @intSerialNo And Branch = @Branch
INSERT INTO tFacCard
                      (Branch ,intSerialNo, PosId ,intAmount , NvcTraceNo , CardNumber , TransTime  )
SELECT @Branch ,@intSerialNo , c7 ,c8 , c9 , c10 , dbo.setTimeFormat(getdate())
FROM dbo.SplitFactorReceived(@ds)
WHERE c1 = 5
     IF @@ERROR <>0      
        GoTo EventHandler       

SET @Result = 1

COMMIT TRAN

Return @Result      

EventHandler:      

    ROLLBACK TRAN      
    SET @Result = -1      

    RETURN @Result


GO


--Script_V26_16_Fix10_8_RoundtFacM_D
--روند کردن هر سطر فاکتور خرید برای اینکه مغایرت یک ریالی با روند نرم افزار ویندوزی پیش نیاید
-- قبلا در آخر فاکتور گرد می کرد
--93/10/27


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

            @Status INT ,      
            @Owner INT ,      
            @Customer INT ,      
            @DiscountTotal FLOAT ,      
            @CarryFeeTotal FLOAT ,      
            @Recursive INT ,      
            @InCharge INT ,      
            @FacPayment BIT ,      
            @OrderType INT ,      
            @StationId INT ,      
            @ServiceTotal FLOAT ,      
            @PackingTotal FLOAT ,      
            @TableNo INT ,      
            @User INT ,      
            @Date NVARCHAR(50) ,      
            @DetailsString nText,      
            @ds nText = '',      
            @Balance BIT ,      
            @AccountYear smallint = null  ,       
            @NvcDescription Nvarchar(150) = Null ,      
            @HavaleNo int = Null  ,      
            @TempAddress Nvarchar(255) = '',  
			@GuestNo INT,    
            @lastFacMNo INT OUT  ,
		    @Person INT = NULL     
             )      

AS      

Declare @intserialNo int      
Declare @intserialNo2 int      
--Declare @intserialNo3 Bigint    

SET @intserialNo = 0        
SET @intserialNo2   = 0      
--SET @intserialNo3   = 0      

DECLARE @No1  INT     
DECLARE @No2  INT     
--DECLARE @No3  INT     

DECLARE @SumPrice  float      
Set @SumPrice = 0      

DECLARE @proper_time nvarchar(5)      

IF @AccountYear IS  Null
    Set @AccountYear = dbo.get_AccountYear() 
    
IF  @Owner = 0      
    SET @Owner = NULL      

IF  @TableNo < 1      
    SET @TableNo = NULL      

IF  @Incharge < 1      
    SET @Incharge = NULL      

IF  @Customer=0      
    SET @Customer = NULL      

BEGIN TRAN      

    DECLARE @MasterServePlace INT      
    DECLARE @newtime nvarchar(5)      
    select @newtime=dbo.setTimeFormat(getdate())      
    SELECT @MasterServePlace = SUM(tmpTable.SServePlace)      
    FROM (  SELECT DISTINCT ServePlace As SServePlace      
         FROM Split(@DetailsString)      
           ) tmpTable      

----------------------------------------Date From Server-----------------------------------------------------------------      
If @Status = 2 And dbo.Get_DateFromServer() = 1      
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      

------Start New Line For Avoid Repeat in tFacm------
DECLARE @RepeatNo INT

declare @d1 as datetime

set @d1 = CONVERT(datetime ,@NewTime)
set @d1 = DATEADD(MINUTE,-1,@d1)

--select  CONVERT(VARCHAR(5),@d1,108)

SELECT @RepeatNo = COUNT(*) FROM dbo.tFacM WHERE Status = 2 AND [Date] = @Date 
  --AND [Time] <= @NewTime AND [Time] >= CONVERT(VARCHAR(5),@d1,108) 
    AND TableNo = @TableNo AND Recursive = 0 AND FacPayment = 0 --AND Balance = 0

IF @RepeatNo > 0 
    GOTO EventHandler

----End New Line -----------------------------------------------------------------------------------------------      

 Declare @intBranch  int      
 Declare @ShiftNo int      
 DECLARE @TempNo INT 

 select @intBranch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@DetailsString))      
 IF @intBranch = 0 OR @intBranch IS NULL     SET @intBranch = dbo.Get_Current_Branch()

    DECLARE @IdentityNo INT
    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
    FROM    tFacm
    WHERE   Branch = @intBranch 

    IF @IdentityNo < ( @intBranch * 10000000 ) 
        SET @IdentityNo = ( @intBranch * 10000000 )

 SET @NO1 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND AccountYear = @AccountYear)      

 SET @ShiftNo= dbo.Get_Shift(GETDATE())      
 SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      


     INSERT INTO tFacM (   
		intSerialNo ,   
		[No] ,      
		[Date] ,      
		RegDate ,      
		Status ,      
		Customer ,      
		SumPrice ,      
		OrderType ,      
		ServePlace ,      
		StationId ,      
		ServiceTotal ,      
		Recursive ,      
		CarryFeeTotal ,      
		PackingTotal ,      
		DiscountTotal ,      
		[Time] ,      
		[User] ,      
		TableNo ,      
		shiftNo ,      
		incharge,      
		owner ,      
		FacPayment ,       
		Balance ,       
		Branch,      
		AccountYear ,      
		NvcDescription,      
		TempAddress ,
		GuestNo ,
		TempNo    
		
 )      
     Values       

(	    @IdentityNo ,  
        @NO1 ,      
        @Date ,      
        dbo.Shamsi(GETDATE()) ,      
        @Status,      
        @Customer ,      
        @SumPrice ,      
        @OrderType ,      
        @MasterServePlace ,      
        @StationId ,      
        @ServiceTotal ,      
        @Recursive ,      
        @CarryFeeTotal ,      
        @PackingTotal ,      
        @DiscountTotal ,      
        @newtime,      
        @User ,      
        @TableNo,      
        @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
        @Incharge ,      
        @owner ,      
        @FacPayment ,      
        @Balance ,      
		@intBranch , --dbo.Get_Current_Branch(),      
		@AccountYear ,      
		@NvcDescription,      
		@TempAddress,
		@GuestNo,
		@TempNo  
 )      
     IF @@ERROR <>0      
        GoTo EventHandler       

    SET @intserialNo = @IdentityNo

declare @destbranch  INT 
SET @destbranch = 0
DECLARE @TempNo2 INT 
DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      
 
If @Status = 6 AND @DestinventoryNo > 0  -- And (@destbranch= @intBranch Or dbo.AutoResid() = 1)    
	Begin      
	select @destbranch=  @intBranch --   branch from tInventory where inventoryNo=(SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

	  SET @NO2 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @destbranch AND AccountYear = @AccountYear)      
	  --SET @TempNo2 = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      

     INSERT INTO tFacM ( 
				intSerialNo ,     
                [No] ,      
                [Date] ,      
                RegDate ,      
                Status ,      
                Customer ,      
                SumPrice ,      
                OrderType ,      
                ServePlace ,      
                StationId ,      
                ServiceTotal ,      
                Recursive ,      
                CarryFeeTotal ,      
                PackingTotal ,      
                DiscountTotal ,      
                TaxTotal ,
                DutyTotal ,     
                [Time] ,      
                [User] ,      
                TableNo ,      
                shiftNo ,      
                incharge,      
                owner ,      
                FacPayment ,       
                Balance ,       
                Branch,      
			  AccountYear ,      
			  NvcDescription,      
			  TempAddress,
			  GuestNo ,
			  TempNO     

 )      
     Values      
(				@IdentityNo + 1 ,     
                @NO2 ,      
                @Date ,      
                dbo.Shamsi(GETDATE()) ,      
                7,      
                @Customer ,      
                @SumPrice ,      
                @OrderType ,      
                @MasterServePlace ,      
                @StationId ,      
                @ServiceTotal ,      
                @Recursive ,      
                @CarryFeeTotal ,      
                @PackingTotal ,      
                @DiscountTotal ,      
                0 ,
                0 ,      
                @newtime,      
                @User ,      
                @TableNo,      
                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
                @Incharge ,      
                @owner ,      
                @FacPayment ,      
                @Balance ,      
				@DestBranch ,     
				@AccountYear ,      
				@NvcDescription,      
				@TempAddress,
				@GuestNo ,
				NULL --@TempNo2    
		
 )      
		 IF @@ERROR <>0      
			GoTo EventHandler      
		SET @intserialNo2 = @IdentityNo + 1      

            UPDATE  tfacm
            SET     NvcDescription = @NvcDescription + N' رسيد -   '
                    + CAST(@No2 AS NVARCHAR(8))
            WHERE   intSerialNo = @intserialNo
            UPDATE  tfacm
            SET     RefrenceHavale = @intserialNo2
            WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch

end      


----------------------------------Fill Details Factor  --------------------------------------------------------------      
If @Status = 6 AND @DestinventoryNo > 0 -- AND (@destbranch= @intBranch  Or dbo.AutoResid() = 1)        
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch      
Else       
 exec InsertFactorDetail @DetailsString , @intserialNo , 0, @Customer , @intBranch      

     IF @@ERROR <>0      
        GoTo EventHandler      
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------      

----------------------------------Total SumPrice Calculate  --------------------------------------------------------------      
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(SUM(ROUND( (Amount * FeeUnit * discount/100 ) ,0) ) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(SUM( ROUND((Amount * FeeUnit) * (1 - discount/100),0) )  AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

     IF @@ERROR <>0      
        GoTo EventHandler      
DECLARE @TaxTotal FLOAT  
SET @TaxTotal = 0
DECLARE @ValueGoodsTax FLOAT
SET @ValueGoodsTax = 0
IF @status = 2 OR @status = 5  OR @status = 10
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxBuy = 1
	            )  
SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
IF @@ERROR <> 0 
GOTO EventHandler

DECLARE @DutyTotal INT 
SET @DutyTotal = 0
DECLARE @ValueGoodsDuty FLOAT
SET @ValueGoodsDuty = 0
IF @status = 2 OR @status = 5  OR @status = 10
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutyBuy = 1
	            )  
SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
IF @@ERROR <> 0 
GOTO EventHandler

Declare @SumPrice2 Bigint      
Set @SumPrice2 = (Select Cast(Sum(Amount * FeeUnit) as Bigint) From tFacd Where intSerialNo = @intserialNo2 And Branch = @DestBranch )        
     IF @@ERROR <>0      
        GoTo EventHandler      
----------------------------------ServiceRate Calculate  --------------------------------------------------------------      
Declare @ReserveServiceRate Int      
Set @ReserveServiceRate = 0      

If  @TableNo >0      
Begin      
	Declare @Reserve Bit      
	Set @Reserve = (Select Reserve From tTable Where tTable.[No] = @TableNo)      
	If @Reserve = 1      
		Set @ReserveServiceRate = (Select t.ReserveServiceRate From(Select ReserveServiceRate From tPartitions  inner join tTable        
		On tPartitions.PartitionId = tTable.PartitionId and  tTable.[No] = @TableNo)t )      

        Update dbo.tTable      
           Set   dbo.tTable.Empty  = 0      
                Where dbo.tTable.[No] = @TableNo AND  @Balance = 0    
	If dbo.Get_TableMonitoring() = 1   ---Table Monitoring      
	 	Begin      
		DECLARE @intTableUsedNo INT      
--		SELECT @intTableUsedNo=intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
--		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@intBranch      
		DECLARE @nvcString NVARCHAR(100)      
		SET @nvcString=','+CAST(@TableNo AS NVARCHAR(5))+'/'      
		--IF @intTableUsedNo is NULL      
		EXEC insert_tblSamar_TableUsage @nvcString,1      
--		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.nvcStartTime=  @newtime      
--		FROM    ( SELECT     dbo.vwSamar_TableUsage_BusyTable.intTableUsedNo, dbo.vwSamar_TableUsage_BusyTable.nvcStartTime,       
--		                 dbo.vwSamar_TableUsage_BusyTable.intBranch, dbo.tTable.[No]      
--				FROM         dbo.tTable LEFT OUTER JOIN      
--		                 dbo.vwSamar_TableUsage_BusyTable ON dbo.vwSamar_TableUsage_BusyTable.intTableNo = dbo.tTable.[No] AND       
--		                 dbo.vwSamar_TableUsage_BusyTable.intBranch = dbo.tTable.Branch)t      
--		WHERE  tblSamar_TableUsage.intTableNo=t.[No] and tblSamar_TableUsage.intBranch=t.intBranch      
--		and tblSamar_TableUsage.intTableNo=@TableNo and tblSamar_TableUsage.intBranch= @intBranch     
		END        
End      
     IF @@ERROR <>0      
        GoTo EventHandler      


If @ReserveServiceRate > 0       
 Set @ServiceTotal = @ReserveServiceRate      


 If @ServiceTotal <> 0      
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)      
     IF @@ERROR <>0      
        GoTo EventHandler       
----------------------------------Round Sumprice  --------------------------------------------------------------      
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5  OR @status = 10
 BEGIN 
  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
 END 
    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal     

    Declare @Remain INT
    SET @Remain = 0  
    IF @Status = 2 OR @Status = 10
    BEGIN   
    Set @Remain = dbo.RoundSumPrice(@SumPrice )         
    Set @SumPrice = @SumPrice - @Remain      
    Set @DiscountTotal = @DiscountTotal + @Remain    
    END  
---select @Remain as remain      
----------------------------------Calculate Packing---------------------------------------------------------------      
If dbo.Get_AutoPacking() = 1      
Begin      
    Declare @UserPacking INT      
    SET @UserPacking = ISNULL((Select ISNULL(Sum(Amount),0)  From tGood inner join tFacD On tFacD.GoodCode = tGood.Code       
        where intSerialNo = @intSerialNo and MainType =1 and Branch = @intBranch Group By MainType) ,0)      
    Set @UserPacking = @UserPacking * [dbo].[Get_User_Packing]()      
    Set @SumPrice = @SumPrice + @UserPacking      
    Update tFacm Set PackingTotal = @UserPacking  Where intSerialNo = @intserialNo And Branch = @intBranch       
End      
----------------------------------Net Price Update  --------------------------------------------------------------      

Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch      
Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DiscountTotal = @DiscountTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

If @Status = 6 AND @DestinventoryNo > 0-- AND (@destbranch= @intBranch )  -- Or dbo.AutoResid() = 1   
	Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch       
      IF @@ERROR <>0       

        GoTo EventHandler           
-------------------------------------Fill Detail Cash , ..........--------------------------------------------------      
DECLARE @Result INT 
IF (@Status =  1 OR @Status = 2 )      
	 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds ,  @intBranch  , @Remain ,1  , @Result OUT   

     IF @@ERROR <>0      
   GoTo EventHandler      
IF @Result = -1
   GoTo EventHandler      

-------------------------------------Monitoring---------------------------------------------------------------------      
--Declare  @Monitor1 int      
--Declare  @Monitor2 int       

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  @intBranch)      
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  @intBranch)      


--IF @Monitor1 > 0       
--   exec Notify_to_Clients      

--Else If @Monitor2 > 0       
--   exec Notify_to_Clients      

----------------------------History---------------------------      

Exec InsertHistory  @No1, @Status , @User , 1 , @AccountYear , @intBranch      
     IF @@ERROR <>0      
   GoTo EventHandler      
IF @STATUs = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 7 , @User , 1 , @AccountYear , @destbranch      
     IF @@ERROR <>0      
   GoTo EventHandler      


----------------------------Cash ---------------------------      

------------------------Mojodi Control Online--------------------------------------------      

Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @intBranch      
IF @@ERROR <>0      
 GoTo EventHandler      
IF @STATUS = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1      
 BEGIN      
 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch      
 IF @@ERROR <>0      
 GoTo EventHandler      
 Exec InsertMojodiCalculate  7,  @intserialNo2 , @AccountYear , @destbranch      
 IF @@ERROR <>0      
 GoTo EventHandler      
 END       

------------------------------------------Update Balance After Recived----------------------------
Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch


COMMIT TRAN

--DECLARE @TemporaryNo BIT 
--SELECT @TemporaryNo = TemporaryNo FROM dbo.tStations WHERE StationID = @StationId AND Branch = @intBranch
--IF @TemporaryNo = 0 set @lastFacMNo = @No1
--ELSE set @lastFacMNo = @TempNo

set @lastFacMNo = @intserialNo


---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @lastFacMNo , 1

--------------------------------------------------------------------------------------------------------------------------------------


Return @lastFacMNo      

EventHandler:      

    ROLLBACK TRAN      
    SET @LastFacMNo = -1      

    RETURN @lastFacMNo
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--تغییر تاریخ فاکتور خرید و عدم تغییر تاریخ فاکتور فروش
alter   PROCEDURE [dbo].[EditFactorMasterDetails]  (  


	@No       INT,  
	@Status  INT ,  
	@Owner  INT ,  
	@Customer  INT ,  
	@DiscountTotal Float ,  
	@CarryFeeTotal Float ,  
	@Recursive  INT ,  
	@InCharge  INT ,  
	@FacPayment  BIT ,  
	@OrderType  INT ,  
	@StationId  INT ,  
	@ServiceTotal  Float ,  
	@PackingTotal  Float ,  
	@TableNo  INT ,  
	@User INT ,  
	@Date   Nvarchar(50) =NULL,  
	@DetailsString  nText,  
	@ds nText = '',  
	@Balance Bit,  
	@AccountYear Smallint = Null ,  
	@NvcDescription Nvarchar(150) = Null ,  
	@TempAddress Nvarchar(255) = '', 
	@GuestNo INT,     
	@LastFacMNo  INT OUT  ,
	@Person INT = NULL 
  )  


AS  
DECLARE @SumPrice BIGINT  
DECLARE @SumPrice2 BIGINT  
DECLARE @intSerialNo BIGINT  
DECLARE @intSerialNo2 BIGINT  
--DECLARE @intSerialNo3 BIGINT  
DECLARE @OldRegDate Nvarchar(50)  
DECLARE  @FactorSerial BIGINT  

SET @Sumprice = 0  
SET @Sumprice2= 0  
SET @intSerialNo = 0  
SET @intSerialNo2 = 0  
--SET @intserialNo3 = 0  


 Declare @intBranch  int  
 Declare @ShiftNo int  

 Declare @DestBranch INT  
 SET @DestBranch = 0

IF @AccountYear IS  Null
    Set @AccountYear = dbo.get_AccountYear() 

 select @intBranch = branch from tInventory where inventoryNo=(SELECT TOP 1  IntInventoryNo FROM Split(@DetailsString))  
 SET @ShiftNo= dbo.Get_Shift(GETDATE())  

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  

--Control is difficult
--If No received then Bypass received
--DECLARE @DestinventoryNo INT 
--select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

IF @Status = 6  
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  


if @status=10   
set @OldRegDate = (SELECT tFacM.regdate FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  
else set @OldRegDate=dbo.Shamsi(GETDATE())  
-------------No Change StationId , If this Fich Is For Pocket Pc---------------------------------------  
DECLARE @OldStationId INT  
 SET @OldStationId = (Select StationId From tFacm Where intserialNo = @intSerialNo and Branch =  dbo.Get_Current_Branch())  

DECLARE @StationType INT  
 SET @StationType = (Select StationType From tStations Where StationId = @OldStationId and Branch =  dbo.Get_Current_Branch())  
If  @StationType = 8  
 SET @StationId = @OldStationId  
----------------------------------------------------------------------------------------------------------  
IF  @Owner = 0  
    SET @Owner = NULL  

IF  @TableNo < 1  
    SET @TableNo = NULL  

Declare @OldTableNo   int  

SET  @OldTableNo =  IsNull((SELECT tFacM.TableNo FROM tFacM WHERE intSerialNo = @intSerialNo and Branch = dbo.Get_Current_Branch()) , 0)  

IF  @Incharge < 1  
    SET @Incharge = NULL  

IF  @Customer=0  
    SET @Customer = NULL  
IF @Date IS NULL  
 SET @Date=Rtrim(LTRIM(dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())))  

BEGIN TRANSACTION  

If IsNull(@TableNo , 0) <> @OldTableNo  
BEGIN  
 IF @OldTableNo > 0   
	-- Add For Tablet & Ppc
	DECLARE @TableNotEmpty INT 

	SELECT @TableNotEmpty = COUNT(*) FROM dbo.tFacM WHERE Status = 2 AND [Date] = @Date 
	  --AND [Time] <= @NewTime AND [Time] >= CONVERT(VARCHAR(5),@d1,108) 
	  AND TableNo = @TableNo AND Recursive = 0 AND FacPayment = 0 --AND Balance = 0

		IF @TableNotEmpty > 0 
			GOTO EventHandler

	 Update ttable SET Empty = 1 where No = @OldTableNo  
END  

    DECLARE @MasterServePlace INT  

 SELECT @MasterServePlace = SUM(tmpTable.SServePlace)  
 FROM   
 (  SELECT DISTINCT ServePlace As SServePlace  FROM Split(@DetailsString)) tmpTable  


 if @Status = 2  
 begin  
       INSERT INTO tRepFacEditM (Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance , OrderType,  
                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate, AccountYear , TaxTotal , DutyTotal  )  
          SELECT Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance, OrderType,  
                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate , AccountYear , TaxTotal , DutyTotal    
   FROM tFacM WHERE tFacM.intSerialNo = @intSerialNo and Branch = @intBranch  

      IF @@ERROR <>0  
          GoTo EventHandler  

      INSERT INTO tFacD2(Code , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate], intInventoryNo )   
    SELECT @@identity , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate],intInventoryNo  
                 From tFacD  
                 WHERE intSerialNo = @intSerialNo  And Branch = @intBranch

      IF @@ERROR <>0  
          GoTo EventHandler  

 end  



If @status = 6 AND @intSerialNo2 > 0 --And (@destbranch = dbo.Get_Current_Branch()  )--or dbo.AutoResid() = 1   
 select @destbranch=branch from tInventory where inventoryNo=(SELECT TOP 1 DestInventoryNo FROM Split(@DetailsString))  

---------------------------------------Mojodi Control Online---------------------------------------------------------  
Exec DeleteMojodiCalculate @Status , @intserialNo  ,  1 , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = @intBranch )  --or dbo.AutoResid() = 1 
 Exec DeleteMojodiCalculate 7 , @intserialNo2  , 1 , @AccountYear , @DestBranch  
----------------------------------------Delete Old Details -----------------------------------------------------------  
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = @intBranch or dbo.AutoResid() = 1 )   
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo2 AND Branch =  @DestBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
------------------------------------------------------------    
  Exec DeleteFactorChildren @intSerialNo , @intBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
 If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
  Exec DeleteFactorChildren @intSerialNo2 , @DestBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
----------------------------------------Date From Server-----------------------------------------------------------------  
If @Status = 2 And dbo.Get_DateFromServer() = 1  
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())  
----------------------------------------Update Master-----------------------------------------------------------------  

    Update tFacM  
        SET Owner       = @Owner,  
        Customer        = @Customer,  
        DiscountTotal   = @DiscountTotal,  
        CarryFeeTotal   = @CarryFeeTotal,  
        SumPrice        = @SumPrice,  
        Recursive       = @Recursive,   
        InCharge        = @InCharge,  
        FacPayment      = @FacPayment,  
        Balance         = @Balance,  
        OrderType       = @OrderType,  
        ServePlace      = @MasterServePlace,  
        StationId       = @StationId,  
        ServiceTotal    = @ServiceTotal,  
        PackingTotal    = @PackingTotal,  
        ShiftNo         = @ShiftNo , --dbo.Get_Shift(GETDATE()),  
        TableNo         = @TableNo,  
        [Date]          =  CASE WHEN @Status = 2 THEN  [Date] WHEN @Status = 5 THEN [Date] ELSE @Date END,  
        [Time]          = dbo.SetTimeFormat(GETDATE()),  
        [User]          = @User,  
        RegDate 	= @OldRegDate,---dbo.Shamsi(GETDATE()),  
        NvcDescription  = @NvcDescription ,  
 		TempAddress     = @TempAddress,
		GuestNo		= @GuestNo ,
		TempNo = CASE WHEN @Status = 2 THEN  TempNo WHEN @Status = 5 THEN TempNo ELSE NULL END      
    WHERE tFacM.intSerialNo = @intSerialNo  AND Branch =  @intBranch  

    IF @@ERROR <>0  
        GoTo EventHandler  

DECLARE @No2 INT
SET @No2 = 0

If @Status = 6 AND @intSerialNo2 > 0 --And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
Begin  
    SET @NO2 = (SELECT [NO] FROM tFacM WHERE intserialNo = @intSerialNo2  And Branch =  @DestBranch )      

    Update tFacM  
        SET Owner       = @Owner,  
        Customer        = @Customer,  
        DiscountTotal   = @DiscountTotal,  
        CarryFeeTotal   = @CarryFeeTotal,  
        SumPrice        = @SumPrice,  
        Recursive       = @Recursive,   
        InCharge        = @InCharge,  
        FacPayment      = @FacPayment,  
        Balance         = @Balance,  
        OrderType       = @OrderType,  
        ServePlace      = @MasterServePlace,  
        StationId       = @StationId,  
        ServiceTotal    = @ServiceTotal,  
        PackingTotal    = @PackingTotal,  
        ShiftNo         = @ShiftNo ,--dbo.Get_Shift(GETDATE()),  
        TableNo         = @TableNo,  
        [Date]          = @Date,  
        [Time]          =dbo.SetTimeFormat(GETDATE()),  
        [User]          = @User,  
        RegDate 	= dbo.Shamsi(GETDATE()),  
        NvcDescription  = @NvcDescription,  
 		TempAddress     = @TempAddress ,
		GuestNo		= @GuestNo  
    WHERE tFacM.intSerialNo = @intSerialNo2  AND Branch =  @DestBranch  

END  

----------------------------------Fill Details Factor ----------------------------------------------------------------------  
If @Status = 6  AND @intSerialNo2 > 0--AND (@destbranch= @intBranch Or dbo.AutoResid() = 1 )    
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch  
Else         
 exec InsertFactorDetail @DetailsString , @intserialNo , 0 , @Customer , @intBranch        

     IF @@ERROR <>0  
        GoTo EventHandler  
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------  


----------------------------------Total SumPrice Calculate  --------------------------------------------------------------  
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(SUM(ROUND( (Amount * FeeUnit * discount/100 ) ,0) ) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(SUM( ROUND((Amount * FeeUnit) * (1 - discount/100)  ,0)) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

     IF @@ERROR <>0      
        GoTo EventHandler      
DECLARE @TaxTotal FLOAT  
SET @TaxTotal = 0
DECLARE @ValueGoodsTax FLOAT
SET @ValueGoodsTax = 0
IF @status = 2 OR @status = 5 OR @status = 10
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxBuy = 1
	            )  
SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
IF @@ERROR <> 0 
GOTO EventHandler

DECLARE @DutyTotal INT 
SET @DutyTotal = 0
DECLARE @ValueGoodsDuty FLOAT
SET @ValueGoodsDuty = 0
IF @status = 2 OR @status = 5 OR @status = 10
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutyBuy = 1
	            )  
SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
IF @@ERROR <> 0 
GOTO EventHandler

If @Status = 6 AND @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1 )  
 Set @SumPrice2 = (Select Cast (Sum(Amount * FeeUnit) as Bigint)   From tFacd Where intSerialNo = @intSerialNo2 And Branch = @DestBranch )    
IF @@ERROR <>0  
        GoTo EventHandler  
----------------------------------ServiceRate Calculate  --------------------------------------------------------------  
Declare @ReserveServiceRate Int  
Set @ReserveServiceRate = 0  
If  @TableNo >0  
Begin  
	Declare @Reserve Bit  
	Set @Reserve = (Select Reserve From tTable Where tTable.[No] = @TableNo)  
	If @Reserve = 1  
		Set @ReserveServiceRate = (Select t.ReserveServiceRate From(Select ReserveServiceRate From tPartitions  inner join tTable    
		On tPartitions.PartitionId = tTable.PartitionId and  tTable.[No] = @TableNo)t )  


	If   @Recursive = 0  
	 Update dbo.tTable  
	    Set   dbo.tTable.Empty  = 0  
	        Where dbo.tTable.[No] = @TableNo  AND @Balance = 0
	
	if  @Recursive = 1  
         Update dbo.tTable  
            Set   dbo.tTable.Empty  = 1  
                Where dbo.tTable.[No] = @TableNo  

	If dbo.Get_TableMonitoring() = 1 AND IsNull(@TableNo , 0) <> @OldTableNo   ---Table Monitoring      
	 	Begin      
		DECLARE @intTableUsedNo INT      
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@OldTableNo and vwSamar_TableUsage_BusyTable.intBranch=@intBranch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.intTableNo = @TableNo      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
		END        

End  

If @ReserveServiceRate > 0   
 Set @ServiceTotal = @ReserveServiceRate  


 If @ServiceTotal <> 0  
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)  

     IF @@ERROR <>0  
        GoTo EventHandler   
----------------------------------Round Sumprice  --------------------------------------------------------------  
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5 OR @status = 10
 BEGIN 
  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
 END 
    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal   

    Declare @Remain INT  
    SET @Remain = 0
    IF @Status = 2 OR @status = 10
    BEGIN
    Set @Remain = dbo.RoundSumPrice(@SumPrice )     
    Set @SumPrice = @SumPrice - @Remain  
    Set @DiscountTotal = @DiscountTotal + @Remain  
    END
----------------------------------Calculate Packing---------------------------------------------------------------  
IF dbo.Get_AutoPacking() = 1  
Begin  
    Declare @UserPacking INT  
    SET @UserPacking = ISNULL((Select ISNULL(Sum(Amount),0)  From tGood inner join tFacD On tFacD.GoodCode = tGood.Code   
 where intSerialNo = @intSerialNo and MainType =1 and Branch = @intBranch Group By MainType) ,0)  
    Set @UserPacking = @UserPacking * [dbo].[Get_User_Packing]()  
   Set @SumPrice = @SumPrice + @UserPacking  
   Update tFacm Set PackingTotal = @UserPacking  Where intSerialNo = @intserialNo And Branch = @intBranch   
End  
----------------------------------Net Price Update  --------------------------------------------------------------  

    Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch   
 IF @@ERROR <>0  
         GoTo EventHandler  
If @Status = 6 AND @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1)   

    Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch  
 IF @@ERROR <>0  
         GoTo EventHandler  

Update tFacm Set DiscountTotal = @DiscountTotal Where intSerialNo = @intserialNo  And Branch = @intBranch   
Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch   
Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch   
Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

-----------------------------------------Fill Detail Cash ,....---------------------------------------------------  
DECLARE @Result INT 
If (@Status = 2 OR @Status = 1)  
 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds  , @intBranch  , @Remain  , 2 , @Result OUT 
 IF @@ERROR <>0  
        GoTo EventHandler  
IF @Result = -1
   GoTo EventHandler      
-----------------------------------------Monitoring  --------------------------------------------------------------  

--Declare  @Monitor1 int  
--Declare  @Monitor2 int  

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  


--If @Monitor1 > 0   
--  exec Notify_to_Clients  
--Else If @Monitor2 > 0   
--  exec Notify_to_Clients  

-- IF @@ERROR <>0  
--        GoTo EventHandler  

-----------------------------------------History  --------------------------------------------------------------  

Exec InsertHistory  @No, @Status , @User , 2 ,@AccountYear  , @intBranch
 IF @@ERROR <>0  
        GoTo EventHandler  

-----------------------------------------Cash  --------------------------------------------------------------  

------------------------------------------Mojodi Control Online-----------------------------------------------------  

Exec InsertMojodiCalculate @Status , @intserialNo , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
IF @STATUs = 6 AND @intSerialNo2 > 0--AND (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
 BEGIN  
 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch  
 IF @@ERROR <>0  
 GoTo EventHandler  
 Exec InsertMojodiCalculate  7,  @intserialNo2 , @AccountYear , @destbranch  
 IF @@ERROR <>0  
 GoTo EventHandler  
 END   
------------------------------------------Update Balance After Recived----------------------------
Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch


COMMIT TRANSACTION  

---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @intSerialNo , 2

--------------------------------------------------------------------------------------------------------------------------------------
Set @LastFacMNo = @No  
Return @LastFacMNo  


EventHandler:  
    ROLLBACK TRAN  
    SET @LastFacMNo = -1   

    RETURN @LastFacMNo
GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


-------------------------------------*******************************************
--اصلاح نحوه جستجو در فرم جستجو کالا فاکتور خرید , فروش , کاردکس کالا و غیره********


ALTER   PROC [dbo].[Get_Good_Name]
    (
      @Name1 NVARCHAR(20),
      @NotSupportedGoodType INT
    )
AS 

Set @Name1 = Replace(  @Name1  , N'ك', N'ک'  ) 
Set @Name1 = Replace(  @Name1  , N'ي' , N'ی' )

    SELECT  *
    FROM    [dbo].[vw_Good]
    WHERE   CHARINDEX(@Name1, [Name]) > 0
            AND [dbo].[vw_Good].[GoodType] NOT IN ( @NotSupportedGoodType, 4 )
    ORDER BY [Name]

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   proc Update_Good_btnAscDefault 
AS

Update tGood SET btnAscDefault = ASCII(left(ltrim([Name] COLLATE Arabic_CI_AI),1))

Update tGood
Set [Name] = Replace(  [Name]   , N'ك', N'ک' ) 

Update tGood
Set [NamePrn] = Replace(  [NamePrn]  , N'ك', N'ک'  ) 

Update tPocketPC_Good
Set [NameDisplay] = Replace(  [NameDisplay]  , N'ك', N'ک'  ) 

Update tGood
Set [Name] = Replace(  [Name], N'ي' , N'ی'  ) 

Update tGood
Set [NamePrn] = Replace(  [NamePrn] , N'ي', N'ی'   ) 

Update tPocketPC_Good
Set [NameDisplay] = Replace(  [NameDisplay]  , N'ي', N'ی'  ) 

Update [tCust]
Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

	Update tper 
	Set nvcFirstName = Replace(  nvcFirstName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcFirstName = Replace(  nvcFirstName , N'ي' , N'ی' ) 

	Update tper 
	Set nvcSurName = Replace(  nvcSurName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcSurName = Replace(  nvcSurName , N'ي' , N'ی' ) 



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  PROCEDURE dbo.InserttGood
(
	@intLanguage	INT,
	@Code		INT,
	@GoodName	NVARCHAR(50),
	@GoodNamePrn	NVARCHAR(50),
	@SellPrice	FLOAT,
	@BuyPrice	FLOAT,
	@Barcode	NVARCHAR(50),
	@Level1		INT,
	@Level2		INT,
	@Model		INT,
	@Supplier	INT,
	@Unit		INT,
	@GoodType	INT,
	@Weight	Float,
	@NumberOfUnit 	INT,
	@SellPrice2 Float,
	@SellPrice3 Float ,
	@MainType Bit ,
	@SellPrice4 Float,
	@SellPrice5 Float,
	@SellPrice6 FLOAT ,
	@CategoryShow INT ,
	@PicturePath NVARCHAR(100) ,
	@nvcDescription NVARCHAR(100) ,
	@Picture IMAGE ,
	@GoodNamePrn2	NVARCHAR(100),
	@GoodNamePrn3	NVARCHAR(100),
	@Result INT OUT 
	


)

AS
Begin tran

	IF @intLanguage = 0 

		INSERT INTO dbo.tGood (Code,Name,NamePrn,SellPrice,BuyPrice,Barcode,Level1,Level2,Model,ProductCompany,Unit,GoodType,Weight,NumberOfUnit,SellPrice2 ,SellPrice3 , MainType , SellPrice4 ,SellPrice5 ,SellPrice6 , CategoryShow , PicturePath , nvcDescription , GoodNamePrn2 , GoodNamePrn3 )
		VALUES (@Code,dbo.Get_ArabicToFarsiString(@GoodName),dbo.Get_ArabicToFarsiString(@GoodNamePrn),@SellPrice,@BuyPrice,@Barcode,@Level1,@Level2,@Model,@Supplier,@Unit,@GoodType,@Weight,@NumberOfUnit, @SellPrice2 , @SellPrice3 , @MainType , @SellPrice4, @SellPrice5, @SellPrice6 , @CategoryShow , @PicturePath , @nvcDescription , dbo.Get_ArabicToFarsiString(@GoodNamePrn2) , dbo.Get_ArabicToFarsiString(@GoodNamePrn3)  )

	ELSE IF @intLanguage = 1 

		INSERT INTO dbo.tGood (Code,LatinName,LatinNamePrn,SellPrice,BuyPrice,Barcode,Level1,Level2,Model,ProductCompany,Unit,GoodType,Weight,NumberOfUnit, SellPrice2, SellPrice3 , MainType , SellPrice4 ,SellPrice5 ,SellPrice6 , CategoryShow ,PicturePath ,nvcDescription , GoodNamePrn2 , GoodNamePrn3)
		VALUES (@Code,@GoodName,@GoodNamePrn,@SellPrice,@BuyPrice,@Barcode,@Level1,@Level2,@Model,@Supplier,@Unit,@GoodType,@Weight,@NumberOfUnit , @SellPrice2 , @SellPrice3 , @MainType , @SellPrice4, @SellPrice5, @SellPrice6 ,@CategoryShow ,@PicturePath ,@nvcDescription , dbo.Get_ArabicToFarsiString(@GoodNamePrn2) , dbo.Get_ArabicToFarsiString(@GoodNamePrn3)  )

     IF @@ERROR <>0
        GoTo EventHandler

 --         if  @GoodType = 3 
   --          Begin
     --               INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
	--	VALUES (@Code,@Code,1,1)
               --     INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
		--VALUES (@Code,@Code,2,1)
                  --  INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
--	--	VALUES (@Code,@Code,4,1)
              --      INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
	--	VALUES (@Code,@Code,8,1)
             --      INSERT INTO dbo.tUsePercent (GoodCode,GoodFirstCode,intServePlace,fltUsedValue)
	--	VALUES (@Code,@Code,16,1)
         --   End

	update  [dbo].[tGood]  set [name] = latinname where ([Name] is null or [Name] = '' ) And Code = @Code

	update  [dbo].[tGood] set latinname = [name] where ([latinName] is null or latinname = '') And Code = @Code

	update  [dbo].[tGood] set [nameprn]=[latinnameprn] where ([Nameprn] is null or [Nameprn] = '' ) And Code = @Code

	update  [dbo].[tGood] set [GoodNamePrn2]=[nameprn] where ([GoodNamePrn2] is null or [GoodNamePrn2] = '' ) And Code = @Code

	update  [dbo].[tGood] set [GoodNamePrn3]=[nameprn] where ([GoodNamePrn3] is null or [GoodNamePrn3] = '' ) And Code = @Code

	update  [dbo].[tGood] set [latinnameprn] = [nameprn] where ([latinNameprn] is null or latinnameprn = '') And Code = @Code

	UPDATE dbo.tGood SET Picture = @Picture WHERE Code = @Code
	
insert into .dbo.[tInventory_Good](GoodCode,InventoryNo,Branch , AccountYear )
  select @Code,tInventory.InventoryNo,tInventory.Branch , dbo.Get_AccountYear()  from dbo.tInventory 
		inner join tInventory_Level1 On tInventory_Level1.Branch = tInventory.Branch  and tInventory_Level1.InventoryNo  = tInventory.InventoryNo 
	        Where tInventory_Level1.Level1 = @Level1 

     IF @@ERROR <>0
        GoTo EventHandler

Insert into dbo.tStation_Inventory_Good ( branch ,InventoryNo, StationID,  AccountYear , GoodCode , Active)
select tInventory.Branch , tInventory.InventoryNo,  t.stationid , dbo.Get_AccountYear() , @Code, 1 as active  from dbo.tInventory -- t.stationid
		inner join tInventory_Level1 On tInventory_Level1.Branch = tInventory.Branch  and tInventory_Level1.InventoryNo  = tInventory.InventoryNo 
         	Inner join (select  Distinct(StationId), InventoryNo , Branch , AccountYear from tStation_Inventory_Good )t On  t.InventoryNo = tInventory_Level1.InventoryNo AND t.Branch =  tInventory_Level1.Branch and t.AccountYear = dbo.Get_AccountYear()

	        Where tInventory_Level1.Level1 = @Level1 

     IF @@ERROR <>0
        GoTo EventHandler
	UPDATE dbo.tGood SET nvcDate = dbo.shamsi(GETDATE()) WHERE Code = @Code
	Update tGood SET btnAscDefault = ASCII(left(ltrim([Name] COLLATE Arabic_CI_AI),1)) where code =  @Code


	Update tGood
	Set [Name] = Replace(  [Name]  , N'ك', N'ک'  ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn]  , N'ك', N'ک'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'ك', N'ک'  ) 

	Update tGood
	Set [Name] = Replace(  [Name] , N'ي' , N'ی' ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn] , N'ي' , N'ی'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'ي', N'ی'  ) 


Commit Tran

SET @Result = 1 
RETURN @Result

EventHandler:
	RollBack Tran
SET @Result = -1 
RETURN @Result



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER     PROCEDURE dbo.UpdatetGood
(
	@intLanguage	INT,
	@Goodname	NVARCHAR(50),
	@GoodNamePrn	NVARCHAR(50),
	@SellPrice	FLOAT,
	@BuyPrice	FLOAT,
	@Unit		INT,
	@GoodType	INT,
	@Barcode	NVARCHAR(50),
	@Code		INT,
	@Weight	Float,
	@NumberOfUnit 	INT,
	@SellPrice2 Float,
	@SellPrice3 Float ,
	@MainType Bit ,
	@Supplier Int ,
	@Level1 Int ,
	@Level2 Int ,
	@SellPrice4 Float,
	@SellPrice5 Float,
	@SellPrice6 Float,
	@CategoryShow INT ,
	@PicturePath NVARCHAR(100) ,
	@nvcDescription NVARCHAR(100) ,
	@Picture IMAGE ,
	@GoodNamePrn2	NVARCHAR(100),
	@GoodNamePrn3	NVARCHAR(100),
	@RealNewCode INT ,
	@Result Int Out
)

AS

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tUsePercent_tGood1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
		ALTER TABLE [dbo].[tUsePercent] DROP CONSTRAINT [FK_tUsePercent_tGood1]



Declare  @NewCode INT
SET @NewCode = @RealNewCode
IF @RealNewCode = 0 
BEGIN 
	Set @NewCode = @Code
	Declare @Level2Code	INT
	Set @Level2Code = (Select Level2 From tGood Where Code = @Code)

	Begin Tran

	If @Level2 <>  @Level2Code
	Begin
	--	Set @NewCode =  (SELECT  ISNULL(MAX(RIGHT(RTRIM(LTRIM(STR(code))),LEN(RTRIM(LTRIM(STR(Code))))-4)),0) +1   
		Set @NewCode =  (SELECT  ISNULL(MAX(code),0) + 1   
		FROM dbo.tgood 
		WHERE Level2 = @Level2 )
		If Len(@NewCode) = 1 

		Set @NewCode = (@Level2 * 10000) + @NewCode 

	End

END 

IF @intLanguage = 0 
Begin		
		UPDATE dbo.tGood

		SET [Name]    = dbo.Get_ArabicToFarsiString(@GoodName) ,
		    NamePrn   = dbo.Get_ArabicToFarsiString(@GoodNamePrn) ,
		    SellPrice = @SellPrice ,
		    BuyPrice  = @BuyPrice ,
		    Unit      = @Unit ,
		    GoodType  = @GoodType ,
		    Barcode = @Barcode,
	                 Weight = @Weight,
		    NumberOfUnit=@NumberOfUnit,
		    SellPrice2 = @SellPrice2,
		    SellPrice3 = @SellPrice3 ,	    	
		    SellPrice4 = @SellPrice4 ,	    	
		    SellPrice5 = @SellPrice5 ,	    	
		    SellPrice6 = @SellPrice6 ,	    	
		    MainType = @MainType  ,
		    ProductCompany = @Supplier ,
		   Level1 = @Level1 ,
		   Level2 = @Level2 ,
		 Code = @NewCode ,
		 CategoryShow = @CategoryShow ,
		 PicturePath = @PicturePath ,
		 nvcDescription = @nvcDescription ,
	    GoodNamePrn2   = dbo.Get_ArabicToFarsiString(@GoodNamePrn2) ,
	    GoodNamePrn3   = dbo.Get_ArabicToFarsiString(@GoodNamePrn3) 
	WHERE Code = @Code		
        IF @@ERROR <>0
	        GoTo EventHandler

End
ELSE IF @intLanguage = 1 
Begin
		UPDATE dbo.tGood

		SET LatinName     = @GoodName ,
		    LatinNamePrn  = @GoodNamePrn ,
		    SellPrice     = @SellPrice ,
		    BuyPrice      = @BuyPrice ,
		    Unit          = @Unit ,
		    GoodType      = @GoodType,
		    Barcode = @Barcode,
		    Weight = @Weight,
		    NumberOfUnit=@NumberOfUnit,
		    SellPrice2 = @SellPrice2,
		    SellPrice3 = @SellPrice3 ,
		    SellPrice4 = @SellPrice4 ,	    	
		    SellPrice5 = @SellPrice5 ,	    	
		    SellPrice6 = @SellPrice6 ,	    	
		    MainType = @MainType ,
	 	    ProductCompany = @Supplier ,
		   Level1 = @Level1 ,
		   Level2 = @Level2 ,
		 Code = @NewCode ,
		 CategoryShow = @CategoryShow ,
		 PicturePath = @PicturePath ,
		 nvcDescription = @nvcDescription ,
	    GoodNamePrn2   = @GoodNamePrn2 ,
	    GoodNamePrn3   = @GoodNamePrn3
		WHERE Code = @Code

        IF @@ERROR <>0
	        GoTo EventHandler

End
Set @Result = 1
	update  [dbo].[tGood]  set [name] = latinname where ([Name] is null or [Name] = ''  ) And Code = @NewCode

	update  [dbo].[tGood] set latinname = [name] where ([latinName] is null or latinname = '') And Code = @NewCode

	update  [dbo].[tGood] set [nameprn]=[latinnameprn] where ([Nameprn] is null or [Nameprn] = '') And Code = @NewCode 

	update  [dbo].[tGood] set [GoodNamePrn2]=[nameprn] where ([GoodNamePrn2] is null or [GoodNamePrn2] = '' ) And Code = @NewCode

	update  [dbo].[tGood] set [GoodNamePrn3]=[nameprn] where ([GoodNamePrn3] is null or [GoodNamePrn3] = '' ) And Code = @NewCode

	update  [dbo].[tGood] set [latinnameprn] = [nameprn] where ([latinNameprn] is null or latinnameprn = '') And Code = @NewCode

	UPDATE dbo.tGood SET Picture = @Picture WHERE Code = @NewCode
	
	UPDATE [tUsePercent] SET GoodFirstCode = @NewCode WHERE GoodFirstCode = @Code
	
	ALTER TABLE [dbo].[tUsePercent]  WITH CHECK ADD  CONSTRAINT [FK_tUsePercent_tGood1] FOREIGN KEY([GoodFirstCode])
		REFERENCES [dbo].[tGood] ([Code])

	ALTER TABLE [dbo].[tUsePercent] CHECK CONSTRAINT [FK_tUsePercent_tGood1]

	update  [dbo].[tGood] SET nvcDate = dbo.shamsi(GETDATE()) WHERE Code = @Code

	Update tGood
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک'  ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn]  , N'ك', N'ک'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay]  , N'ك', N'ک'  ) 

	Update tGood
	Set [Name] = Replace(  [Name] , N'ي' , N'ی' ) 

	Update tGood
	Set [NamePrn] = Replace(  [NamePrn] , N'ي' , N'ی'  ) 

	Update tPocketPC_Good
	Set [NameDisplay] = Replace(  [NameDisplay] , N'ي', N'ی'   ) 

COMMIT TRANSACTION


Return @Result


EventHandler:
    ROLLBACK TRAN
    Set @Result = 0
    RETURN @Result



GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  PROCEDURE [dbo].[InsertPersonel]( 
	@PersonnelNumber nvarchar(50),
	@nvcFirstName nvarchar(50),
	@nvcSurName nvarchar(50),
	@Gender bit,
	@IdNumber nvarchar(50),
	@Job int,
	@InsuranceNo nvarchar(50) ,
	@Address nvarchar(300),
	@Tel nvarchar(30),
	@User int , 
	@UserName nvarchar(50) ,
	@Password nvarchar(50) ,
	@intAccessLevel int ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno int out

	)
 AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

set @Time = dbo.SetTimeFormat(getdate())

select @pPno = isnull(max(Ppno),0) + 1 from tper Where Branch = @Branch 
If @pPno < (@Branch * 1000 ) Set @pPno = (@Branch * 1000 )


begin Tran
insert into dbo.tper (
	pPno ,
	PersonnelNumber,
	nvcFirstName,
	nvcSurName,
	Gender ,
	IdNumber,
	Job ,
	InsuranceNo  ,
	Address ,
	Tel ,
	[Date] ,
	[Time] ,			
	[User] ,
	Branch,
	MaxCredit,
	ActDeAct

)
values(
	@pPno ,
	@PersonnelNumber,
	@nvcFirstName,
	@nvcSurName ,
	@Gender ,
	@IdNumber ,
	@Job ,
	@InsuranceNo ,
	@Address ,
	@Tel ,
	@Date,
	@Time ,
	@User ,
	@Branch,
	@MaxCredit,
	@ActDeAct
)
if @@Error <> 0 
		GOTO EventHandler	


--set @pPno=@@identity
DECLARE @UID INT
if @intAccessLevel<>0 and @UserName <> '' and @Password<>''

BEGIN

	select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch 
	If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )

	insert into dbo.tUser 
	(
		[Uid] ,
	 UserName ,
	 [Password] ,
	 intAccessLevel ,
	 pPno ,
	 addUser , 
	 Branch, 
	 CountRePrint, 
	 CountInvoicePrint,
	 CountInvoiceEditable,
	 CountInvoiceRefferable
	)
 values (
	@UID ,					
	@UserName  ,
	@Password  ,
	@intAccessLevel ,
	@pPno , 
	@User ,
	@Branch,
	@CountRePrint,
	@CountInvoicePrint,
	@CountInvoiceEditable,
	@CountInvoiceRefferable
	)

if @@Error <> 0 
		GOTO EventHandler	
--SET @UID = @@IDENTITY		
END	


	Update tper 
	Set nvcFirstName = Replace(  nvcFirstName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcFirstName = Replace(  nvcFirstName , N'ي' , N'ی' ) 

	Update tper 
	Set nvcSurName = Replace(  nvcSurName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcSurName = Replace(  nvcSurName , N'ي' , N'ی' ) 


commit Tran



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1




GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER     PROCEDURE [dbo].[UpdatePersonel]( 
	@CurrentPPNO 		INT,
	@PersonnelNumber 	NVARCHAR(50),
	@nvcFirstName 		NVARCHAR(50),
	@nvcSurName	 	NVARCHAR(50),
	@Gender 		BIT,
	@IdNumber 		NVARCHAR(50),
	@Job 			INT,
	@InsuranceNo 		NVARCHAR(50) ,
	@Address 		NVARCHAR(300),
	@Tel 			NVARCHAR(30),
	@User 			INT , 
	@UID 			INT ,
	@UserName 		NVARCHAR(50) ,
	@Password 		NVARCHAR(50) ,
	@intAccessLevel 	INT ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno 			INT OUT
	       )
AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

SET @Time= dbo.SetTimeFormat(getdate())

BEGIN TRANSACTION

	UPDATE tPer
		SET PersonnelNumber 	= @PersonnelNumber,
		    nvcFirstName    	= @nvcFirstName,
		    nvcSurName	    	= @nvcSurName,
		    Gender	    	= @Gender,
		    IdNumber       	= @IdNumber,
		    Job		    	= @Job,
		    InsuranceNo     	= @InsuranceNo,
		    Address	    	= @Address,
		    Tel   	    	= @Tel,
		    [Date]	    	= @Date,
		    [Time]	    	= @Time,
		    [User]	    	= @User,
		    MaxCredit		=@MaxCredit,
		    ActDeAct 		=@ActDeAct ,
		    Branch			= @Branch
	WHERE       pPNO = @CurrentPPNO  


	IF @@ERROR <> 0 
		GOTO EventHandler	

	set @pPno = @CurrentPPNO

	IF @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID<>0
		UPDATE tUser
			SET 		UserName       	= @UserName,
	        	   		 	[Password]     	= @Password,
			    		intAccessLevel 	= @intAccessLevel,
			    		pPno           	= @pPno,
			    		addUser        	= @User,
					 CountRePrint		=@CountRePrint,
		  			 CountInvoicePrint	=@CountInvoicePrint,
					 CountInvoiceEditable		=@CountInvoiceEditable,
		  			 CountInvoiceRefferable	=@CountInvoiceRefferable ,
		  			 Branch					= @Branch
			WHERE   UID = @UID    
	else 
		if @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID=0
		BEGIN 
			select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch   
			If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )  
			insert into dbo.tUser (
						UID ,
						UserName , 
						[Password] , 
						intAccessLevel , 
						pPno , 
						addUser , 
						Branch,
						CountRePrint,
						CountInvoicePrint,
						CountInvoiceEditable,
		  			 	CountInvoiceRefferable
			) values (	
						@UID ,				
						@UserName  ,
						@Password  ,
						@intAccessLevel ,
						@pPno , 
						@User , 
						@Branch,
						@CountRePrint,
						@CountInvoicePrint,
						@CountInvoiceEditable,
		  			 	@CountInvoiceRefferable
			)			
			END 
	IF @@ERROR <> 0 
		GOTO EventHandler	


	Update tper 
	Set nvcFirstName = Replace(  nvcFirstName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcFirstName = Replace(  nvcFirstName , N'ي' , N'ی' ) 

	Update tper 
	Set nvcSurName = Replace(  nvcSurName  , N'ك', N'ک'  ) 

	Update tper
	Set nvcSurName = Replace(  nvcSurName , N'ي' , N'ی' ) 


COMMIT TRANSACTION



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER    Procedure dbo.Insert_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,  
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,
	@TotalRemainingAmount INT ,   
	@Code Bigint out  
)  

as  

Begin Tran  

DECLARE @Branch INT 
 Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode ) --AND (Branch = @Branch ) )  
    end   
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId) <> 0 -- AND Branch = @Branch ) <> 0   
  Goto ErrHandler   

--Set @Code = (Select  isnull(Max(Code),0) + 1 from tSupplier where code > 0)  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

insert Into dbo.tSupplier  
(   
	--Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	State ,  
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	Discount,   
	[Description],   
	[Date],   
	[Time],   
	[User], 
	TotalRemainingAmount , 
	Branch  
)  
values  
(   
	--@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@State,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@Discount,   
	@Description,   
	@Date,   
	@Time,   
	@User , 
	@TotalRemainingAmount , 
	@Branch  
)  
if @@Error <> 0   
 goto ErrHandler  

Set @Code = @@Identity  
 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] )   

update tSupplier set address=dbo.addressedit(address) where code=@code  -- AND Branch = @Branch
   

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

Commit Tran
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code  
--Select @Code


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   Procedure dbo.Update_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint , 
	@TotalRemainingAmount  INT , 
	@Updated Bigint out  

)  

as  

Begin Tran  
--IF @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode )--  AND (Branch = @Branch ) )  
    end  
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code ) <> 0 --  AND (Branch = @Branch ) ) <> 0    
  Goto ErrHandler   
 else  

  Update dbo.tSupplier     
   Set MembershipId = @MembershipId   

  Where MasterCode = @Code   --AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tSupplier  

 Set MembershipId = @MembershipId ,  
 MasterCode  = @MasterCode ,    
 Owner = @Owner ,  
 Name = @Name ,  
 Family = @Family ,  
 Sex = @Sex ,  
 WorkName = @WorkName ,   
 InternalNo = @InternalNo ,  
 Unit = @Unit ,  
 State = @State ,  
 City = @City ,  
 ActKind = @ActKind ,  
 ActDeAct = @ActDeAct ,  
 Prefix = @Prefix ,  
 Address = @Address ,  
 PostalCode = @PostalCode ,  
 Tel1 = @Tel1 ,  
 Tel2 = @Tel2 ,  
 Tel3 = @Tel3 ,  
 Tel4 = @Tel4 ,  
 Mobile = @Mobile ,  
 Fax = @Fax ,  
 Email = @Email ,  
 Flour = @Flour ,  
 Discount = @Discount ,  
 [Description] = @Description ,  
 [Date] = @Date ,  
 [Time] = @Time ,  
 [User] = @User ,
 TotalRemainingAmount = @TotalRemainingAmount 
Where Code = @Code  -- AND (Branch = @Branch )  


if @@Error <> 0   
 goto ErrHandler  

 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] or dbo.tSupplier.[Branch] is Null)  

update tSupplier set address=dbo.addressedit(address) where code=@code  --AND (Branch = @Branch )

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 


Commit Tran   
Set @Updated = @Code  
return @Updated  

ErrHandler:  
RollBack Tran  
Set @Updated = 0  
return @Updated



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER     Procedure dbo.Insert_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int, 
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Code Bigint out 

)  

as  

Begin Tran  

--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  
if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  )  --AND (Branch = @Branch )
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode   )--AND (Branch = @Branch )
 end   
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId ) <> 0 --AND Branch = @Branch)   
  Goto ErrHandler   

Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  Branch = @Branch )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

if @nvcRFID = N''  
  SET @nvcRFID=N'-999'  

insert Into dbo.tCust  
(   
	Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Assansor,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	CarryFee,   
	PaykFee,   
	Distance,   
	Credit,   
	Discount,   
	BuyState,   
	[Description],   
	[Date],   
	[Time],   
	[User],  
	FamilyNo ,  
	Member ,  
	State ,  
	Central ,  
	Branch,  
	nvcRFID,  
	sellprice ,
	EconomicCode ,
	nvcBirthDate ,
	TotalRemainingAmount
	
)  
values  
(   
	@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Assansor,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@CarryFee,   
	@PaykFee,   
	@Distance,   
	@Credit,   
	@Discount,   
	@BuyState,   
	@Description,   
	@Date,   
	@Time,   
	@User ,  
	@FamilyNo ,  
	@Member ,  
	@State ,  
	@Central ,  
	@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate ,
	@TotalRemainingAmount
	
)  
if @@Error <> 0   
 goto ErrHandler  

--Set @Code = @@Identity  
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
  and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address)  
 , nvcRFID=CAST(Branch AS NVARCHAR(1))+CAST(Code AS NVARCHAR(8))  
  where code=@code  AND Branch = @Branch 


Update [tCust]
Set [Name] = Replace(  [Name] , N'ك' , N'ک'  ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي' , N'ی' ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ي' , N'ی' ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ي' , N'ی' ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ي' , N'ی' ) 


Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   Procedure dbo.Update_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint ,  
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Updated Bigint out  

)  

as  

Begin Tran  
--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  ) --AND (Branch = @Branch ) )  
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode )   --AND (Branch = @Branch ) )  
 end  
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code  ) <> 0  -- AND (Branch = @Branch )    
  Goto ErrHandler   
 else  

  Update dbo.tCust     
   Set MembershipId = @MembershipId   

   Where MasterCode = @Code   AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tCust  

 Set MembershipId = @MembershipId ,  
	MasterCode  = @MasterCode ,    
	Owner = @Owner ,  
	Name = @Name ,  
	Family = @Family ,  
	Sex = @Sex ,  
	WorkName = @WorkName ,   
	InternalNo = @InternalNo ,  
	Unit = @Unit ,  
	City = @City ,  
	ActKind = @ActKind ,  
	ActDeAct = @ActDeAct ,  
	Prefix = @Prefix ,  
	Assansor = @Assansor ,  
	Address = @Address ,  
	PostalCode = @PostalCode ,  
	Tel1 = @Tel1 ,  
	Tel2 = @Tel2 ,  
	Tel3 = @Tel3 ,  
	Tel4 = @Tel4 ,  
	Mobile = @Mobile ,  
	Fax = @Fax ,  
	Email = @Email ,  
	Flour = @Flour ,  
	CarryFee = @CarryFee ,  
	PaykFee = @PaykFee ,  
	Distance = @Distance ,  
	Credit = @Credit ,  
	Discount = @Discount ,  
	BuyState = @BuyState ,  
	[Description] = @Description ,  
	[Date] = @Date ,  
	[Time] = @Time ,  
	[User] = @User ,  
	FamilyNo = @FamilyNo ,  
	Member = @Member ,  
	State = @State ,  
	Central = @Central,  
	Sellprice=@Sellprice  ,
	EconomicCode = @EconomicCode ,
	nvcRFID = @nvcRFID ,
	nvcBirthDate = @nvcBirthDate ,
	TotalRemainingAmount = @TotalRemainingAmount
	
Where Code = @Code   AND (Branch = @Branch )   

if @@Error <> 0   
 goto ErrHandler  


Set @Updated = @Code   
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
 and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address) where code=@code  AND Branch = @Branch 
 
Update [tCust]
Set [Name] = Replace(  [Name] , N'ك' , N'ک'  ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي' , N'ی' ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ي' , N'ی' ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ي' , N'ی' ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ي' , N'ی' ) 


Commit Tran  
return @Updated  

ErrHandler:  
RollBack Tran  
return -1



GO
SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

--Script_V26_16_Fix10_9_FarsiCodePage
--اصلاح ی و ک فارسی در دیتا بیس برای کالاها و مشتریان و تامین کنندگان و پرسنل
--در برنامه فروش در قسمت جستجوی کالا اگر اصلاح ک و ی فارسی را بزنیم همه دیتاهای قبلی درست می شود
--93/10/29


ALTER   Proc Get_Customer_Name
@ActDeact int ,
@Name Nvarchar(50)     
as    

Set @Name = Replace(  @Name  , N'ك', N'ک'  ) 
Set @Name = Replace(  @Name  , N'ي' , N'ی' )

Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where  CHARINDEX ( @Name , [Name] ) > 0 and actdeact <> @ActDeact  -- AND Branch = @Branch
AND vw_Get_Cust.Code <> -1
Order By [Name]



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  Proc Get_Supplier_Name
@ActDeact int ,
@Name Nvarchar(50) 
as

Set @Name = Replace(  @Name  , N'ك', N'ک'  ) 
Set @Name = Replace(  @Name  , N'ي' , N'ی' )

Select * from dbo.vw_Get_Supplier where  CHARINDEX ( @Name , [Name] ) > 0 and actdeact <> @ActDeact  Order By [Name]



GO


--Script_V26_16_Fix10_10_ResidMovaghat
--چاپ رسید موقت
--نام ریپورت :  ResidMovaghat.rpt
--93/10/29


DECLARE @MaxPrintFormat INT 
SELECT @MaxPrintFormat = ISNULL(MAX(PrintFormat) ,0) + 1 FROM tPrintFormat

INSERT INTO dbo.tPrintFormat
        ( PrintFormat ,
          PrintFormatName ,
          RptFilePath ,
          NoticeNo ,
          LatinRptFilePath ,
          PrintFormatLatinName ,
          Active
        )
VALUES  ( @MaxPrintFormat , -- PrintFormat - int
          N'رسید موقت' , -- PrintFormatName - nvarchar(50)
          N'A4\ResidMOvaghat.rpt' , -- RptFilePath - nvarchar(50)
          NULL  , -- NoticeNo - int
          N'A4\ResidMOvaghat.rpt' , -- LatinRptFilePath - nvarchar(50)
          N'ResidMOvaghat.rpt' , -- PrintFormatLatinName - nvarchar(50)
          1  -- Active - bit
        )
        
GO



--Script_V26_16_TTMS
--اضافه شدن فرم محاسبه مالیات و عوارض بر ارزش افزوده
--گذاشتن دسترسی برای فرم ارزش افزوده
-- اضافه شدن کد اقتصادی و شناسه (کد) ملی به تامین کنندگان
--برای ارسال گزارش دقیق به دارایی
--94/02/05


ALTER TABLE dbo.tSupplier
ADD EconomicCode NVARCHAR(20) NULL ,
    NationalCode NVARCHAR(20) NULL 
GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 330 , -- intObjectCode - int
          N'frmTTMS' , -- ObjectId - nvarchar(50)
          N'گزارشات ارزش افزوده' , -- ObjectName - nvarchar(50)
          N'frmTTMS' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          102  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          330  -- intObjectCode - int
          )
          
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER     Procedure dbo.Insert_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,  
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,
	@TotalRemainingAmount INT ,   
	@EconomicCode nVarChar(20) = NULL ,
	@NationalCode nVarChar(20) = NULL ,
	@Code Bigint out  
	
)  

as  

Begin Tran  

DECLARE @Branch INT 
 Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode ) --AND (Branch = @Branch ) )  
    end   
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId) <> 0 -- AND Branch = @Branch ) <> 0   
  Goto ErrHandler   

--Set @Code = (Select  isnull(Max(Code),0) + 1 from tSupplier where code > 0)  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

insert Into dbo.tSupplier  
(   
	--Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	State ,  
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	Discount,   
	[Description],   
	[Date],   
	[Time],   
	[User], 
	TotalRemainingAmount , 
	Branch ,
	EconomicCode ,
	NationalCode
	 
)  
values  
(   
	--@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@State,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@Discount,   
	@Description,   
	@Date,   
	@Time,   
	@User , 
	@TotalRemainingAmount , 
	@Branch ,
	@EconomicCode ,
	@NationalCode 
)  
if @@Error <> 0   
 goto ErrHandler  

Set @Code = @@Identity  
 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] )   

update tSupplier set address=dbo.addressedit(address) where code=@code  -- AND Branch = @Branch
   

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

Commit Tran
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code  
--Select @Code


GO
SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  Procedure dbo.Update_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint , 
	@TotalRemainingAmount  INT , 
	@EconomicCode nVarChar(20) = NULL ,
	@NationalCode nVarChar(20) = NULL ,  
	@Updated Bigint OUT  

)  

as  

Begin Tran  
--IF @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode )--  AND (Branch = @Branch ) )  
    end  
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code ) <> 0 --  AND (Branch = @Branch ) ) <> 0    
  Goto ErrHandler   
 else  

  Update dbo.tSupplier     
   Set MembershipId = @MembershipId   

  Where MasterCode = @Code   --AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tSupplier  

 Set MembershipId = @MembershipId ,  
 MasterCode  = @MasterCode ,    
 Owner = @Owner ,  
 Name = @Name ,  
 Family = @Family ,  
 Sex = @Sex ,  
 WorkName = @WorkName ,   
 InternalNo = @InternalNo ,  
 Unit = @Unit ,  
 State = @State ,  
 City = @City ,  
 ActKind = @ActKind ,  
 ActDeAct = @ActDeAct ,  
 Prefix = @Prefix ,  
 Address = @Address ,  
 PostalCode = @PostalCode ,  
 Tel1 = @Tel1 ,  
 Tel2 = @Tel2 ,  
 Tel3 = @Tel3 ,  
 Tel4 = @Tel4 ,  
 Mobile = @Mobile ,  
 Fax = @Fax ,  
 Email = @Email ,  
 Flour = @Flour ,  
 Discount = @Discount ,  
 [Description] = @Description ,  
 [Date] = @Date ,  
 [Time] = @Time ,  
 [User] = @User ,
 TotalRemainingAmount = @TotalRemainingAmount ,
 EconomicCode = @EconomicCode ,
 NationalCode = @NationalCode
Where Code = @Code  -- AND (Branch = @Branch )  


if @@Error <> 0   
 goto ErrHandler  

 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] or dbo.tSupplier.[Branch] is Null)  

update tSupplier set address=dbo.addressedit(address) where code=@code  --AND (Branch = @Branch )

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 


Commit Tran   
Set @Updated = @Code  
return @Updated  

ErrHandler:  
RollBack Tran  
Set @Updated = 0  
return @Updated



GO



--اضافه شدن نام همه بانک ها به دیتابیس
--تغییر در پوز بانکی که بتواند همه بانک ها را پوشش دهد
--940116

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPub_Pos_tblPub_PosPort]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblPub_Pos] DROP CONSTRAINT FK_tblPub_Pos_tblPub_PosPort
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_PosPort]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_PosPort]
GO

CREATE TABLE [dbo].[tblPub_PosPort]
(
[PortId] [int] NOT NULL,
PortName NVARCHAR(50) NOT NULL 
) ON [PRIMARY]
GO


ALTER TABLE [dbo].[tblPub_PosPort] ADD CONSTRAINT [PK_tblPub_PosPort] PRIMARY KEY CLUSTERED  ([PortId] ) ON [PRIMARY]
GO


INSERT INTO [dbo].[tblPub_PosPort]
        ( [PortId], [PortName] )
VALUES  ( 1, -- PortId - int
          N'Usb'  -- PortName - nvarchar(50)
          )

GO

INSERT INTO [dbo].[tblPub_PosPort]
        ( [PortId], [PortName] )
VALUES  ( 2, -- PortId - int
          N'Serial'  -- PortName - nvarchar(50)
          )

GO

INSERT INTO [dbo].[tblPub_PosPort]
        ( [PortId], [PortName] )
VALUES  ( 3, -- PortId - int
          N'Lan'  -- PortName - nvarchar(50)
          )

GO

----###

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_All_PosPort') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_All_PosPort
GO

CREATE PROC Get_All_PosPort
AS 
select * from [dbo].[tblPub_PosPort] ORDER BY PortId

GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacCard_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT FK_tFacCard_tblPub_PosType
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPub_Pos_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblPub_Pos] DROP CONSTRAINT FK_tblPub_Pos_tblPub_PosType
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsPrimary') = 1)
ALTER TABLE [dbo].[tblPub_PosType] DROP CONSTRAINT [PK_tblPub_PosType]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_PosType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_PosType]
GO

CREATE TABLE [dbo].[tblPub_PosType]
(
[PosTypeId] [int] NOT NULL,
PosName NVARCHAR(50) NOT NULL 
) ON [PRIMARY]
GO


ALTER TABLE [dbo].[tblPub_PosType] ADD CONSTRAINT [PK_tblPub_PosType] PRIMARY KEY CLUSTERED  ([PosTypeId] ) ON [PRIMARY]
GO


INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 1, -- PosId - int
          N'پوز آسان پرداخت'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 2, -- PosId - int
          N'پوز بانکي پاسارگاد'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 3, -- PosId - int
          N'پوز بانکي ایران کیش'  -- PosName - nvarchar(50)
          )
GO

INSERT INTO dbo.tblPub_PosType
        ( PosTypeId, PosName )
VALUES  ( 4, -- PosId - int
          N'پوز بانکي ملت'  -- PosName - nvarchar(50)
          )
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblAcc_ReceivedSummary_tPos]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblAcc_ReceivedSummary] DROP CONSTRAINT FK_tblAcc_ReceivedSummary_tPos
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacCard_tblPub_PosType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT [FK_tFacCard_tblPub_PosType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacCard_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacCard] DROP CONSTRAINT FK_tFacCard_tblPub_Pos
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPub_Pos_tblAcc_Bank]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].tblPub_Pos DROP CONSTRAINT FK_tblPub_Pos_tblAcc_Bank
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_Pos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_Pos]
GO

CREATE TABLE [dbo].[tblPub_Pos] (
	[PosId] [int] IDENTITY (1, 1) NOT NULL ,
	[StationId] [int] NOT NULL ,
	[intBank] [int] NULL ,
	[nvcAccountNo] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AccountId] [int] NULL ,
	[PosType] [int] NULL ,
	[ComunicationType] [int] NULL ,
	[PosAddress] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PosPort] [int] NULL ,
	[nvcPosSerialNo] nvarchar(20)COLLATE SQL_Latin1_General_CP1_CI_AS NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblPub_Pos] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblPub_Pos] PRIMARY KEY  CLUSTERED 
	(
		[PosId]
	)  ON [PRIMARY] 
GO

--########

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblAcc_RecieveSanad_tblAcc_Bank]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].tblAcc_RecieveSanad DROP CONSTRAINT FK_tblAcc_RecieveSanad_tblAcc_Bank
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tblAcc_Bank]') and OBJECTPROPERTY(id, N'IsPrimaryKey') = 1)
ALTER TABLE [dbo].[tblAcc_Bank] DROP CONSTRAINT [PK_tblAcc_Bank]
GO

ALTER TABLE [dbo].[tblAcc_Bank]
ALTER COLUMN [tintBank] INT NOT NULL 
GO

ALTER TABLE [dbo].[tblAcc_Bank] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblAcc_Bank] PRIMARY KEY  CLUSTERED 
	(
		[tintBank]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPub_Pos] ADD 
	CONSTRAINT [FK_tblPub_Pos_tblAcc_Bank] FOREIGN KEY 
	(
		[intBank]
	) REFERENCES [dbo].[tblAcc_Bank] (
		[tintBank]
	) ON DELETE CASCADE  ON UPDATE CASCADE ,
	CONSTRAINT [FK_tblPub_Pos_tblPub_PosType] FOREIGN KEY 
	(
		[PosType]
	) REFERENCES [dbo].[tblPub_PosType] (
		[PosTypeId]
	)  
GO


ALTER TABLE [dbo].[tblPub_Pos] ADD 
	CONSTRAINT [FK_tblPub_Pos_tblPub_PosPort] FOREIGN KEY 
	(
		[ComunicationType]
	) REFERENCES [dbo].[tblPub_PosPort] (
		[PortId]
	) ON UPDATE CASCADE 
GO


ALTER TABLE dbo.tFacCard WITH NOCHECK ADD CONSTRAINT
	FK_tFacCard_tblPub_Pos FOREIGN KEY
	(
	PosId
	) REFERENCES dbo.tblPub_Pos
	(
	PosId
	) ON UPDATE CASCADE
	
GO



DELETE FROM tblAcc_Bank WHERE tintBank >= 13 
GO

INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 13, -- tintBank - int
          N'پست بانک'  -- nvcBankName - nvarchar(25)
          )
GO

INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 14, -- tintBank - int
          N'توسعه صادرات'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 15, -- tintBank - int
          N'صنعت و معدن'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 16, -- tintBank - int
          N'بانک مسکن'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 17, -- tintBank - int
          N'توسعه تعاون'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 18, -- tintBank - int
          N'کارآفرین'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 19, -- tintBank - int
          N'پاسارگاد'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 20, -- tintBank - int
          N'بانک سرمایه'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 21, -- tintBank - int
          N'بانک سینا'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 22, -- tintBank - int
          N'بانک شهر'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 23, -- tintBank - int
          N'بانک دی'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 24, -- tintBank - int
          N'بانک انصار'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 25, -- tintBank - int
          N'بانک حکمت ایرانیان'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 26, -- tintBank - int
          N'بانک ایران زمین'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 27, -- tintBank - int
          N'بانک قوامین'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 28, -- tintBank - int
          N'بانک خاورمیانه'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 29, -- tintBank - int
          N'بانک آینده'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 30, -- tintBank - int
          N'بانک مهر اقتصاد'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 31, -- tintBank - int
          N'بانک مهر ایران'  -- nvcBankName - nvarchar(25)
          )
GO
INSERT INTO dbo.tblAcc_Bank
        ( tintBank, nvcBankName )
VALUES  ( 32, -- tintBank - int
          N'بانک رسالت '  -- nvcBankName - nvarchar(25)
          )
GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tBanks]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tBanks]
GO

CREATE PROCEDURE [dbo].[Get_All_tBanks] AS
select * from dbo.tblAcc_Bank ORDER BY tintBank

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tRecvType_Acc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tRecvType_Acc]
GO


CREATE  PROCEDURE [dbo].[Get_All_tRecvType_Acc] AS
SELECT *
FROM tRecvType
WHERE (tintIsShow = 1)
AND tintType = 1 OR tintType = 5
ORDER BY tintType


GO

UPDATE tRecvType SET nvcDescription = N'کارت بانکی' WHERE tintType = 5
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insert_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Insert_tblPub_Pos]
GO


CREATE  PROCEDURE [dbo].[Insert_tblPub_Pos] 
(	@PosType INT ,
	@StationId INT  , 
	@BankNo INT , 
	@nvcAccountNo nvarchar(50) , 
	@AccountId INT ,
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@nvcPosSerialNo NVARCHAR(20) ,
	@intStatus int out)
AS


Begin Tran


Insert Into dbo.tblPub_Pos
        ( 
          StationId ,
          intBank ,
          nvcAccountNo ,
          AccountId ,
          PosType ,
          ComunicationType ,
          PosAddress ,
          nvcPosSerialNo
        )
VALUES  ( 
          @StationId , -- NvcPosNo - nvarchar(20)
          @BankNo , -- BankName - nvarchar(20)
          @nvcAccountNo , -- nvcAccountNo - nvarchar(20)
          @AccountId ,
          @PosType ,
          @CommunicationType ,
          @PosAddress ,
          @nvcPosSerialNo
        )

if @@Error <> 0 
	Goto ErrHandler

Commit Tran

SET @intStatus=@@IDENTITY
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return


GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_tblPub_Pos]
GO

CREATE  PROCEDURE [dbo].[Update_tblPub_Pos] (
	@PosId INT ,
	@StationId INT  , 
	@BankNo INT , 
	@nvcAccountNo nvarchar(50) ,
	@AccountId INT , 
	@PosType INT ,
	@CommunicationType INT ,
	@PosAddress NVARCHAR(20) ,
	@nvcPosSerialNo NVARCHAR(20) ,
	@intStatus int out)

AS

Begin Tran

UPDATE dbo.tblPub_Pos SET
	StationId = @StationId  , 
	intBank = @BankNo , 
	nvcAccountNo = @nvcAccountNo ,
	AccountId = @AccountId ,
	PosType = @PosType ,
	ComunicationType = @CommunicationType ,
	PosAddress = @PosAddress ,
	nvcPosSerialNo = @nvcPosSerialNo

   WHERE PosId = @PosId

if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus = 1
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tblPub_Pos_ById]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tblPub_Pos_ById]
GO

CREATE  PROCEDURE [dbo].[Get_tblPub_Pos_ById] 
@PosId INT 
AS
select * from [tblPub_Pos] WHERE  PosId = @PosId

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Delete_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Delete_tblPub_Pos]
GO

CREATE   PROCEDURE [dbo].[Delete_tblPub_Pos](
	@PosId INT )
AS
	DELETE FROM dbo.tblPub_Pos WHERE PosId = @PosId

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tblPub_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tblPub_Pos]
GO

CREATE  PROCEDURE [dbo].[Get_All_tblPub_Pos] AS
select * from [tblPub_Pos] ORDER BY PosId

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_PosType]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_All_PosType
GO

CREATE PROC Get_All_PosType
AS 
select * from [dbo].[tblPub_PosType] ORDER BY PosTypeId

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tblPub_Pos_ByStationId]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tblPub_Pos_ByStationId]
GO

CREATE   PROCEDURE [dbo].[Get_tblPub_Pos_ByStationId] 
@StationId INT 
AS
select * , CAST(PosType AS VARCHAR(2)) AS nvcPosType  from [tblPub_Pos] INNER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblPub_Pos.intBank
INNER JOIN dbo.tblPub_PosType ON dbo.tblPub_PosType.PosTypeId = dbo.tblPub_Pos.PosType
INNER JOIN dbo.tblPub_PosPort ON dbo.tblPub_PosPort.PortId = dbo.tblPub_Pos.ComunicationType
WHERE  StationId = @StationId 
ORDER BY PosId

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tfacm_Card_Detail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tfacm_Card_Detail]
GO

CREATE   PROCEDURE [dbo].[Get_tfacm_Card_Detail]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @User1 INT,
      @User2 INT,
      @Station1 INT,
      @Station2 INT,   
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50)
      
    )
AS 
    DECLARE @tmp1 INT  
    DECLARE @tmp2 NVARCHAR(50)  
    DECLARE @Time3 NVARCHAR(50)  
    DECLARE @Time4 NVARCHAR(50)  
    SET @Time3 = @Time1  
    SET @Time4 = @time2  
  
    IF @User2 < @User1 
        BEGIN   
            SET @tmp1 = @User2  
            SET @User2 = @User1  
            SET @User1 = @tmp1   
        END   
  
    IF @Time2 < @Time1 
        BEGIN  
  /*SET @tmp2 = @Time2  
  SET @Time2 = @Time1  
  SET @Time1 = @tmp2*/  
            SET @Time3 = '00:00'  
            SET @Time4 = '24:00'  
        END  
    DECLARE @TimeTitle NVARCHAR(10)  
    SET @TimeTitle = N' ساعت : '  

    SELECT  
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate,
			tfacm.* , dbo.tFacCard.* ,dbo.tblPub_Pos.* 
        FROM    tfacm  
			INNER JOIN dbo.tFacCard ON dbo.tFacM.Branch = dbo.tFacCard.Branch AND dbo.tFacM.intSerialNo = dbo.tFacCard.intSerialNo
			INNER JOIN dbo.tblPub_Pos ON dbo.tFacCard.PosId = dbo.tblPub_Pos.PosId AND dbo.tblPub_Pos.StationId = dbo.tFacM.StationID
    WHERE   [date ] >= @Date1
            AND [date] <= @Date2
            AND [User] >= @User1
            AND [User] < = @User2
            AND ( ( [Time] >= @Time1
                    AND [Time] <= @Time4
                  )
                  OR ( [Time] <= @Time2
                       AND [Time] >= @Time3
                     )
                )
            AND tfacM.StationID >= @Station1
            AND tfacM.StationID <= @Station2
    ORDER BY [Date]  
 
GO




--حذف شماره پوز بانکی از شرح سند
--اضافه کردن شماره حساب به شرح سند

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_SaleSummaryCustom]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_SaleSummaryCustom]
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE  PROCEDURE [dbo].[Get_SaleSummaryCustom]
(
@Branch INT ,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT = 0
)

 AS
BEGIN

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت صندوق' + '  ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date] AS [Name] ,  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
 INNER JOIN dbo.tUser TU ON TU.UID = TF.[User] AND TU.Branch = TF.Branch
 INNER JOIN dbo.tPer TP ON TP.pPno = TU.pPno AND TP.Branch = TU.Branch  
 INNER JOIN dbo.tFacCash TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 7
GROUP BY TF.[User] , TF.[Date]

UNION ALL

--SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت کارت' + N' بانک ' + MIN(TAB.nvcBankName) + N' شماره ' + MIN(TPP.nvcAccountNo) + N' درتاریخ ' +  TF.[Date]  AS [Name],  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TPP.AccountId) AS Tafsili FROM 
SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت کارت' + N' درتاریخ ' +  TF.[Date]  AS [Name],  SUM(TFC.intAmount) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein ,0 AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tFacCard TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
--INNER JOIN dbo.tblPub_Pos TPP on TPP.PosId = TFC.PosId AND Tf.StationID = TPP.StationId
--INNER JOIN dbo.tblAcc_Bank TAB ON TAB.tintBank = TPP.intBank
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 22
GROUP BY TFC.PosId , TF.[Date]


UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت پیک' + ' ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date]  AS [Name],  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)
     		        AND InCharge > 0 
     		        AND Balance = 0 and FacPayment = 0) TF
 INNER JOIN dbo.tPer TP ON TP.pPno = TF.InCharge AND TP.Branch = TF.Branch  and Job = 3
 INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 6
GROUP BY TP.pPno , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'دریافت گارسون' + ' ' + MIN(TP.nvcFirstName) + ' ' + MIN(TP.nvcSurName) + N' در تاریخ ' + TF.[Date]  AS [Name],  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)
     		        AND InCharge > 0 
     		        AND Balance = 0 and FacPayment = 0) TF
 INNER JOIN dbo.tPer TP ON TP.pPno = TF.InCharge AND TP.Branch = TF.Branch  and Job = 9
 INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 6
GROUP BY TP.pPno , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' بدهکاری مشتری' + ' ' + MIN(TC.Name + '' + TC.Family) + N' در تاریخ ' + TF.[Date]  AS [Name] ,  SUM(SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TC.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where  [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
					--AND (InCharge = NULL OR (InCharge > 0 AND FacPayment = 1)) 
     		        AND Balance = 0
     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer > 0 AND Credit > 0)
INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 7
GROUP BY TC.Code , TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' تخفیفات فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , SUM(TF.DiscountTotal) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  2
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'  فروش ' + ' ' + MIN(Ts.[Description]) + N' در تاریخ  ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(Tf.Amount * Tf.FeeUnit) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TS.Tafsili) AS Tafsili FROM 
(SELECT tFacM.* , Amount , FeeUnit , intInventoryNo FROM dbo.tFacM INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tInventory TS ON TS.InventoryNo = TF.intInventoryNo
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  1
GROUP BY TS.InventoryNo , TF.[Date]

--SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' فروش ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice + TF.DiscountTotal - TF.CarryFeeTotal - PackingTotal - ServiceTotal - TaxTotal - DutyTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
--(SELECT * FROM dbo.tFacM
--                    where [Date] >= @DateBefore
--                    AND [Date] <= @DateAfter
--                    AND Recursive = 0
--                    AND Status = 2
--                    AND transferAccounting = 0
--                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
--INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
--INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
--INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  1
--GROUP BY TP.PartitionID , TF.[Date]

--UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' بستانکاری مشتری' + ' ' + MIN(TC.Name + '' + TC.Family) + N' در تاریخ ' + TF.[Date]  AS [Name] ,  0 AS SumBedehKar , SUM(ISNULL(TFC.intAmount,0)+ISNULL(TFCA.intAmount,0)) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TC.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 2
----                    AND transferAccounting = 0
----     		        AND (InCharge = NULL OR (InCharge > 0 AND FacPayment = 1)) 
----     		        AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
----INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer > 0 AND Credit > 0)
----LEFT JOIN dbo.tFacCash TFC ON TFC.Branch = TF.Branch AND TFC.intSerialNo = TF.intSerialNo
----LEFT JOIN dbo.tFacCard TFCA ON TFCA.Branch = TF.Branch AND TFCA.intSerialNo = TF.intSerialNo
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TC.Code , TF.[Date]
----HAVING SUM(ISNULL(TFC.intAmount,0)+ISNULL(TFCA.intAmount,0)) > 0

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'موجودي مواد و کالا' AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM INNER JOIN 
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 16
----GROUP BY TF.[Date]


----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , MIN(TS.Name + ' ' + TS.Family + ' ' + TS.WorkName) AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , TS.Code AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TS.Code , TF.[Date]


----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'برگشت از خريد' AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 4
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tCust TC ON TF.Customer = TC.Code AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 17
----GROUP BY TF.[Date]

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , MIN(TS.Name + ' ' + TS.Family + ' ' + TS.WorkName) AS [Name] , SUM(TF.SumPrice) AS SumBedehKar , 0 AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , TS.Code AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where  [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 1
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 15
----GROUP BY TS.Code , TF.[Date]

----UNION ALL

----SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'برگشت از فروش' AS [Name] , 0 AS SumBedehKar , SUM(TF.SumPrice) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TAS.Tafsili) AS Tafsili FROM 
----(SELECT * FROM dbo.tFacM
----                    where [Date] >= @DateBefore
----                    AND [Date] <= @DateAfter
----                    AND Recursive = 0
----                    AND Status = 5
----                    AND transferAccounting = 0) TF
----INNER JOIN dbo.tSupplier TS ON TS.Code = TF.Owner 
----INNER JOIN dbo.TblAcc_Sale TAS ON TAS.Code = 18
----GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' عوارض فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.DutyTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  24
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' مالیات فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date]  AS [Name] , 0 AS SumBedehKar , SUM(TF.TaxTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  26
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' درآمد سرویس   ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.ServiceTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  38
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N'  درآمد بسته بندی   ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.PackingTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  3
GROUP BY TF.[Date]

UNION ALL

SELECT MIN(TAS.Code) AS [Type] , TF.[Date] , N' کرایه حمل فروش  ' + ' ' + MIN(TP.PartitionDescription) + N' در تاریخ ' + TF.[Date] AS [Name] , 0 AS SumBedehKar , SUM(TF.CarryFeeTotal) AS SumBestankar , MIN(TAS.Kol) AS Kol , MIN(TAS.Moein) AS Moein , MIN(TP.Tafsili) AS Tafsili FROM 
(SELECT * FROM dbo.tFacM
                    where [Date] >= @DateBefore
                    AND [Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0
                    AND (tfacm.[User] = @Uid OR @Uid = 0)) TF
INNER JOIN dbo.tStations TS ON TS.StationID = TF.StationID
INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
INNER JOIN  dbo.TblAcc_Sale TAS ON TAS.Code =  4
GROUP BY TF.[Date]


END

GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER     PROCEDURE [dbo].[Get_SaleSummary]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
SELECT SUM(SumPrice)AS SumPriceTotal ,
        tvw.[Date] ,
        Branch ,
        Tafsili ,
        InventoryName

FROM 
(
SELECT DISTINCT dbo.tFacM.Branch  ,--NO ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacD.intRow ,
                    tfacd.Amount ,
                    tfacd.Feeunit ,
                    ( tfacd.Amount * tfacd.Feeunit ) AS SumPrice ,
                    ISNULL(dbo.tInventory.Tafsili ,0) AS Tafsili ,
                    dbo.tInventory.Description AS InventoryName
          FROM      dbo.tFacM
                    INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                                        AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
					INNER JOIN dbo.tInventory ON dbo.tInventory.InventoryNo = dbo.tFacD.intInventoryNo
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND (dbo.tCust.Tafsili = 0 OR dbo.tCust.Tafsili IS NULL) ))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch , tvw.Tafsili , InventoryName
 ORDER BY tvw.[Date] 
 
 
END

GO


