


--exec Get_DailyGoodForHavale 1, 62, 1393, 2, N'93/01/01', N'93/08/08'
--GO


alter    PROCEDURE [dbo].[Insert_AutoHavale]
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
                                    AND tfacm.bitHavaleResid = 0
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
                                    AND tfacm.bitHavaleResid = 0
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





SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Get_DailyGoodForHavale]
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
                            AND tfacm.bitHavaleResid = 0
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
                            AND tfacm.bitHavaleResid = 0
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


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

-------------------------------------*******************************************
---------------------------------------------------اصلاح گزارش مواد مصرفی********


ALTER   PROC Rep_UseOfGood
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
                0 AS Discount , --T.Discount,
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
                @SystemTime AS SystemTime ,
                AvgBuyPrice
				FROM    
				( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
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
                            0 AS Discount , --dbo.tFacD.Discount,
                            ( SELECT    SellPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [SellPrice],
                            ( SELECT    BuyPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS [BuyPrice] ,
                            ( SELECT    AvgBuyPrice
                              FROM      dbo.tGood
                              WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                            ) AS AvgBuyPrice
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
                            0 AS Discount ,--dbo.tFacD.Discount,
                            dbo.tGood.SellPrice,
                            dbo.tGood.BuyPrice ,
                            dbo.tGood.AvgBuyPrice
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
                --T.Discount,
                T.NumberOfUnit,
                T.SellPrice,
                T.BuyPrice ,
                T.AvgBuyPrice
	
    END


GO


