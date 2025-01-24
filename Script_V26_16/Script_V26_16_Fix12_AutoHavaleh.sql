ALTER PROCEDURE [dbo].[Insert_AutoHavale]
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
        SELECT  @SumTotal = ISNULL(SUM(BuyPrice * Amount ) ,0)
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
