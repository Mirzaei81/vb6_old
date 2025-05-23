

--Script_V26_16_ExcelTransfer
--اضافه شدن فرم انتقال به فایل اکسل
--اضافه شدن فرم دریافت فروش از فایل اکسل
-- با امکان حذف داده های قبلی
--گذاشتن دسترسی برای فرم ایمپورت و اکسپورت
--94/10/07


--SELECT * FROM tObjects
--GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 312 , -- intObjectCode - int
          N'frmExcellTransfer' , -- ObjectId - nvarchar(50)
          N'انتقال با اکسل' , -- ObjectName - nvarchar(50)
          N'frmExcellTransfer' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          102  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          312  -- intObjectCode - int
          )
          
GO

-- ----------------------------
-- Procedure structure for sp_ET_GetFactorDetail
-- ----------------------------

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ET_GetFactorDetail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ET_GetFactorDetail]
GO

CREATE PROCEDURE [dbo].[sp_ET_GetFactorDetail]
    (
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8)
    )
AS
    BEGIN

        DECLARE @strTmp NVARCHAR(50)

        IF @DateAfter < @DateBefore
            BEGIN
                SET @strTmp = @DateAfter
                SET @DateAfter = @DateBefore
                SET @DateBefore = @strTmp
            END

        SELECT  TM.* ,
                TD.Amount ,
                TD.ChairName ,
                TD.DestInventoryNo ,
                TD.intInventoryNo ,
                TD.DifferencesCodes ,
                TD.DifferencesDescription ,
                TD.Discount ,
                TD.ExpireDate ,
                TD.FeeUnit ,
                TD.GoodCode ,
                TD.intRow ,
                TD.Rate ,
                TD.ServePlace AS ServerPlaceDetail ,
                TB.nvcBranchName ,
                TS.Description AS StationName ,
                TP.PartitionID ,
                TP.PartitionDescription ,
                TT.Name AS TableName ,
                ISNULL(Tc.WorkName , ISNULL(TC.NAME,'') + '' + ISNULL(TC.Family,'')) AS CustomerName ,
                ISNULL(TPer.nvcFirstName, '') + ' ' + ISNULL(TPer.nvcSurName,
                                                             '') AS UserName ,
                ISNULL(TPerCharge.nvcFirstName, '') + ' ' + ISNULL(TPerCharge.nvcSurName,
                                                             '') AS PaykName ,
                TG.Name AS GoodName ,
                TI.Description AS InventoryDescription ,
                TSP.Description AS ServerPlaceDescription
        FROM    ( SELECT    *
                  FROM      dbo.tFacM
                  WHERE     [Date] >= @DateBefore
                            AND [Date] <= @DateAfter
                            AND Status = 2
                ) TM
                INNER JOIN dbo.tFacD TD ON TM.Branch = TD.Branch
                                           AND TM.intSerialNo = TD.intSerialNo
                INNER JOIN dbo.tBranch TB ON TB.Branch = TM.Branch
                INNER JOIN dbo.tStations TS ON TS.StationID = TM.StationID
                INNER JOIN dbo.tPartitions TP ON TP.PartitionID = TS.PartitionId
                INNER JOIN dbo.tUser TU ON TU.UID = TM.[User]
                INNER JOIN dbo.tPer TPer ON TPer.pPno = TU.pPno
                INNER JOIN dbo.tGood TG ON TG.Code = TD.GoodCode
                INNER JOIN dbo.tServePlace TSP ON TSP.intServePlace = TM.ServePlace
                INNER JOIN dbo.tInventory TI ON TI.InventoryNo = TD.intInventoryNo
                INNER JOIN dbo.tcust TC ON TC.Code = TM.Customer AND TC.Branch = TM.Branch
                LEFT JOIN dbo.tTable TT ON TT.Branch = TM.Branch
                                           AND TT.No = TM.TableNo
                LEFT JOIN dbo.tPer TPerCharge ON TPerCharge.pPno = TM.InCharge
      ORDER by TM.intserialno   , TD.intRow      
    END

GO


-- ----------------------------
-- Procedure structure for sp_ET_InsertFactor
-- ----------------------------

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ET_InsertFactor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ET_InsertFactor]
GO

CREATE PROCEDURE [dbo].[sp_ET_InsertFactor]
    (
      @Owner INT = NULL ,
      @Customer INT = NULL ,
      @DiscountTotal FLOAT ,
      @CarryFeeTotal FLOAT ,
      @ServiceTotal FLOAT ,
      @PackingTotal FLOAT ,
      @TaxTotal FLOAT ,
      @DutyTotal FLOAT ,
      @SumPrice FLOAT ,
      @Recursive INT ,
      @InCharge INT = NULL ,
      @FacPayment BIT ,
      @OrderType INT ,
      @ServePlace INT ,
      @Branch INT ,
      @StationID INT ,
      @BascoleNo INT ,
      @ShiftNo INT ,
      @TableNo INT = NULL ,
      @intInventoryNo INT ,
      @DestInventoryNo INT ,
      @Date NVARCHAR(50) = NULL ,
      @Time NVARCHAR(50) = NULL ,
      @User INT = NULL ,
      @RegDate NVARCHAR(50) = NULL ,
      @Balance BIT ,
      @AccountYear SMALLINT ,
      @NvcDescription NVARCHAR(150) = NULL ,
      @RefFacM UNIQUEIDENTIFIER = NULL,
      @OrderRefrence BIGINT = NULL ,
      @TempAddress NVARCHAR(255) ,
      @RoundDiscount FLOAT ,
      @SaleSanad INT ,
      @transferAccounting BIT = NULL ,
      @BitHavaleResid BIT ,
      @CreditBalance BIT ,
      @BitLock BIT ,
      @GuestNo INT = NULL ,
      @BitSmsSend BIT ,
      @BitSmsTasvieh BIT ,
      @BitRecursiveSend BIT ,
      @Refrence_Acc INT = NULL,
      @RefrenceHavale INT = NULL ,
      @DestinationId INT = NULL ,
      @BitTempReceived BIT = NULL ,
      @DetailsString NTEXT
    )
AS
    BEGIN TRAN      


    DECLARE @MasterServePlace INT      
    SELECT  @MasterServePlace = SUM(tmpTable.SServePlace)
    FROM    ( SELECT DISTINCT
                        ServePlace AS SServePlace
              FROM      Split(@DetailsString)
            ) tmpTable      

	----------------------------------------Date From Server-----------------------------------------------------------------      

    DECLARE @intserialno BIGINT
    SELECT  @intserialno = ISNULL(MAX(intserialno), 0) + 1
    FROM    tFacm
    WHERE   Branch = @Branch 

    IF @intserialno < ( @Branch * 10000000 )
        SET @intserialno = ( @Branch * 10000000 )

    DECLARE @Status INT
    DECLARE @No BIGINT
    DECLARE @TempNo INT
    DECLARE @ServePlaceTempNo INT

    SET @Status = 2

    SET @NO = ( SELECT  ISNULL(MAX([NO]), 0) + 1
                FROM    tFacM
                WHERE   Status = @Status
                        AND Branch = @Branch
                        AND AccountYear = @AccountYear
              )      

    SET @TempNo = ( SELECT  ISNULL(MAX([TempNo]), 0) + 1
                    FROM    tFacM
                    WHERE   Status = @Status
                            AND Branch = @Branch
                            AND Date = @Date
                            AND ShiftNo = @ShiftNo
                  )      

    SET @ServePlaceTempNo = ( SELECT    ISNULL(MAX(ServePlaceTempNo), 0) + 1
                              FROM      tFacM
                              WHERE     Status = @Status
                                        AND Branch = @Branch
                                        AND Date = @Date
                                        AND ServePlace = @MasterServePlace
                            )      


    INSERT  INTO dbo.tFacM
            ( intSerialNo ,
              No ,
              Status ,
              Owner ,
              Customer ,
              DiscountTotal ,
              CarryFeeTotal ,
              SumPrice ,
              Recursive ,
              InCharge ,
              FacPayment ,
              OrderType ,
              ServePlace ,
              StationID ,
              ServiceTotal ,
              PackingTotal ,
              BascoleNo ,
              ShiftNo ,
              TableNo ,
              Date ,
              Time ,
              [User] ,
              RegDate ,
              Branch ,
              Balance ,
              AccountYear ,
              NvcDescription ,
              RefFacM ,
              OrderRefrence ,
              TempAddress ,
              RoundDiscount ,
              SaleSanad ,
              transferAccounting ,
              BitHavaleResid ,
              CreditBalance ,
              TaxTotal ,
              DutyTotal ,
              BitLock ,
              GuestNo ,
              TempNo ,
              BitSmsSend ,
              BitSmsTasvieh ,
              BitRecursiveSend ,
              Refrence_Acc ,
              RefrenceHavale ,
              DestinationId ,
              BitTempReceived ,
              ServePlaceTempNo
            )
    VALUES  ( @intSerialNo ,
              @No ,
              @Status ,
              @Owner ,
              @Customer ,
              @DiscountTotal ,
              @CarryFeeTotal ,
              @SumPrice ,
              @Recursive ,
              @InCharge ,
              @FacPayment ,
              @OrderType ,
              @ServePlace ,
              @StationID ,
              @ServiceTotal ,
              @PackingTotal ,
              @BascoleNo ,
              @ShiftNo ,
              @TableNo ,
              @Date ,
              @Time ,
              @User ,
              @RegDate ,
              @Branch ,
              @Balance ,
              @AccountYear ,
              @NvcDescription ,
              @RefFacM ,
              @OrderRefrence ,
              @TempAddress ,
              @RoundDiscount ,
              @SaleSanad ,
              @transferAccounting ,
              @BitHavaleResid ,
              @CreditBalance ,
              @TaxTotal ,
              @DutyTotal ,
              @BitLock ,
              @GuestNo ,
              @TempNo ,
              @BitSmsSend ,
              @BitSmsTasvieh ,
              @BitRecursiveSend ,
              @Refrence_Acc ,
              @RefrenceHavale ,
              @DestinationId ,
              @BitTempReceived ,
              @ServePlaceTempNo
            )      

    IF @@ERROR <> 0
        GOTO EventHandler       


----------------------------------Fill Details Factor  --------------------------------------------------------------      
    EXEC InsertFactorDetail @DetailsString, @intserialNo, 0, @Customer,
        @Branch      

    IF @@ERROR <> 0
        GOTO EventHandler      

    IF @Balance = 1
        BEGIN
            INSERT  INTO dbo.tFacCash
                    ( Branch, intSerialNo, intAmount )
            VALUES  ( @Branch, @intserialno, @SumPrice )
        END


    COMMIT TRAN

    RETURN @intserialno      

    EventHandler:      

    ROLLBACK TRAN      
    SET @intserialno = -1      

    RETURN @intserialno
GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ET_DeleteFactors]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ET_DeleteFactors]
GO

CREATE PROCEDURE [dbo].[sp_ET_DeleteFactors]
    (
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8)
    )
AS

	 DECLARE @strTmp NVARCHAR(50)

    IF @DateAfter < @DateBefore
        BEGIN
            SET @strTmp = @DateAfter
            SET @DateAfter = @DateBefore
            SET @DateBefore = @strTmp
        END
        
    DELETE FROM dbo.tFacM WHERE [Date] > = @DateBefore AND [Date] <= @DateAfter
    
    
GO