

--ScriptV26_16_Fix17_950624_FIFO.sql

--95/06/24


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_ISFIFO') )
drop Function [dbo].Get_ISFIFO
GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

CREATE FUNCTION dbo.Get_ISFIFO()
RETURNS int AS  
BEGIN 
return 1
END




GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO
ALTER  PROCEDURE [dbo].[Update_HavalehResid]
    (
	@InventoryNo INT ,
	@AccountYear SMALLINT ,
	@GoodCode INT ,
	@Flag INT ,
	@BeforeDate NVARCHAR(8),
	@AfterDate NVARCHAR(8),
	@NumberOfRecords INT OUT 
    )
AS 
	DECLARE @BuyPrice INT  
	DECLARE @GoodCode1 INT 
	SET @GoodCode1 = @GoodCode
    DECLARE  @intSerialNo INT
    DECLARE @Branch INT 
    DECLARE @fDate NVARCHAR(8)
	DECLARE @FirstPrice INT ;
    DECLARE @priceTamam INT ;
    DECLARE @Mablagh BIGINT ;
    DECLARE @DiscountD BIGINT ;
    DECLARE @Tedad INT ;
    DECLARE @Status1 INT 
    DECLARE @Status2 INT 
    DECLARE @HavaleNo INT
    DECLARE @GoodAmount FLOAT
-- 

IF @Flag = 0 
    SELECT @NumberOfRecords = ISNULL(COUNT(GoodCode), 0)  --, [TIME] 
    FROM [dbo].[tFacM]
    INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
    WHERE [Status] IN( 6,7) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
	AND (GoodCode = @GoodCode OR @GoodCode = 0) 
	AND tFacM.AccountYear = @AccountYear 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate

ELSE IF  dbo.Get_ISFIFO() = 0
BEGIN

    SET  @NumberOfRecords = 0			
    DECLARE  GoodsList CURSOR	 
    FOR 

 SELECT DISTINCT T2.GoodCode , dbo.tGood.BuyPrice FROM 
( SELECT DISTINCT T1.GoodCode  FROM 
(   SELECT  DISTINCT   ISNULL(tUsePercent.GoodFirstCode , tFacD.GoodCode ) AS GoodCode

    FROM    dbo.tFacM
    INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		AND [dbo].[tFacM].Branch = dbo.tFacD.Branch
		LEFT OUTER JOIN dbo.tUsePercent ON dbo.tUsePercent.GoodCode = dbo.tFacD.GoodCode
    WHERE   dbo.tFacM.Status IN ( 1, 2 ,3, 4,5 , 6, 7 )
    AND tFacM.AccountYear = @AccountYear
    AND tFacM.Recursive = 0
    AND [dbo].[tFacD].intInventoryNo = @InventoryNo
    AND (dbo.tFacD.GoodCode = @GoodCode OR @GoodCode = 0)  -- For One Good(FrmGoodTurnOver) or AllGood(FrmFinalPrice) 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate
UNION all
    SELECT  [GoodCode]
	FROM      tInventory_Good
	WHERE     (tInventory_Good.GoodCode = @GoodCode OR @GoodCode = 0)
	AND dbo.tInventory_Good.AccountYear = @AccountYear
	AND [InventoryNo] = @InventoryNo AND tInventory_Good.FirstMojodi <> 0
) T1
GROUP BY GoodCode 
)T2
INNER JOIN dbo.tGood ON T2.GoodCode = dbo.tGood.Code


	
    OPEN GoodsList
    FETCH FROM GoodsList INTO @GoodCode , @BuyPrice

    WHILE @@FETCH_STATUS = 0 
        BEGIN
            --DECLARE @fTime NVARCHAR(8)
            DECLARE Havale CURSOR 
            FOR 
            SELECT DISTINCT tFacM.Branch,tFacM.intSerialNo,[Date] --, GoodCode  
            FROM [dbo].[tFacM]
            INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
            WHERE [Status] IN( 6,7) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
            AND GoodCode = @GoodCode --AND (GoodCode = @GoodCode OR @GoodCode = 0) 
            AND tFacM.AccountYear = @AccountYear 
            AND dbo.tFacM.Date <=@AfterDate-- N'88/06/31'  --*****************
            AND dbo.tFacM.Date>=@BeforeDate
            ORDER BY [Date] ASC , [dbo].[tFacM].intSerialNo ASC  

            OPEN Havale
	
            FETCH  FROM Havale INTO @Branch ,@intSerialNo,@fDate --,@GoodCode 
	
            WHILE @@FETCH_STATUS = 0 
                BEGIN
                    --SELECT @Tedad = ISNULL(FirstMojodi , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    --SELECT @Mablagh = ISNULL(FirstPrice , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    SELECT  @Mablagh = SUM(T.FirstMojodi * T.FirstPrice) + SUM(T.Amount * T.Flag * T.FeeUnit) ,                                          
                            @Tedad = SUM(T.FirstMojodi) + Sum(T.Amount * T.Flag)
                    FROM    (
                              SELECT    dbo.tInventory_Good.FirstMojodi ,
                                        dbo.tInventory_Good.FirstPrice ,
                                        tInventory_Good.GoodCode ,
                                        0 AS Amount ,
                                        0 AS Flag ,
                                        0 AS FeeUnit 
                                        FROM      tInventory_Good
                              WHERE     tInventory_Good.GoodCode = @GoodCode
                                        AND dbo.tInventory_Good.AccountYear = @AccountYear
                                        AND [InventoryNo] = @InventoryNo
                              UNION ALL
                              SELECT    0 AS FirstMojodi ,
                              		0 AS FirstPrice ,
                                 		Goodcode ,
                                        Amount ,
                                        Flag ,
                                        FeeUnit 
                              FROM      dbo.[tFacM]
                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
                                        INNER JOIN dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo
                              WHERE     dbo.tFacM.[Date] <= @fDate 
                              		--( dbo.tFacM.[Date] + ' ' + dbo.tFacM.[Time] ) <= ( @fDate + ' ' + @fTime )
                                        AND tFacM.Status IN ( 1, 3, 4 , 6 , 7 ) -- , 6, 7  براي موجودي منفي
                                        AND (dbo.tFacM.intSerialNo < @intSerialNo OR Status = 1 OR Status = 4)
                                        AND dbo.tFacM.Branch = @Branch
                                        AND dbo.tFacM.AccountYear = @AccountYear
                                        AND tFacD.GoodCode = @GoodCode
                                        AND dbo.tFacM.Recursive = 0
                                        AND [intInventoryNo] = @InventoryNo
										AND dbo.tFacM.Date <=@AfterDate
										AND dbo.tFacM.Date>=@BeforeDate
                            ) T
                    GROUP BY GoodCode 
                    IF @Tedad <= 0 
                    	SET @priceTamam = @BuyPrice
             	    ELSE
             	        SET @priceTamam = CAST((@Mablagh/@Tedad) AS INT)
             	        
                    SET @Status1 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1)         
                    SET @Status2 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo)         
                    SET @HavaleNO = ISNULL((SELECT RefrenceHavale FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1 AND Status = 6) , 0)         
--                    PRINT @GoodCode
--                     PRINT @Mablagh
--                     PRINT @Tedad
--                     PRINT @priceTamam
--                    PRINT @intSerialNo
--		    PRINT @NumberOfRecords + 1
                    --PRINT @HavaleNO
                    IF @priceTamam >= 0 
                        UPDATE  dbo.tFacD
                        SET     FeeUnit = @priceTamam
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
                                AND intSerialNo <> @HavaleNo -- we don,t need update Resid  from Havale
                    IF @priceTamam < 0 			--Negative Price set to Zero
                        UPDATE  dbo.tFacD
                        SET     FeeUnit = 0
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
                                AND intSerialNo <> @HavaleNo -- we don,t need update Resid  from Havale

--Update Resid From Havale With Havale Fee
			IF @intSerialNo = @HavaleNo
				UPDATE dbo.tFacD
				SET FeeUnit = X.feeUnit		
	                        FROM (SELECT feeUnit FROM  dbo.[tFacM]
	                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
	                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
	                                        WHERE   tFacM.Status = 6
	                                        AND dbo.tFacM.intSerialNo = @intSerialNo -1
	                                        AND dbo.tFacM.Branch = @Branch
	                                        AND dbo.tFacM.AccountYear = @AccountYear
	                                        AND dbo.tFacM.Recursive = 0
	                                        AND dbo.tFacD.GoodCode = @GoodCode  
	                                       -- AND [intInventoryNo] = @InventoryNo  --No Inventory Needed because is resid from other inventory
	                                )X  
	                  	WHERE  dbo.tFacD.intSerialNo = @intSerialNo
	                                AND dbo.tFacD.Branch = @Branch
	                                AND dbo.tFacD.GoodCode = @GoodCode             		      
			
--Update Resid when Mojodi is zero or negative
				IF @Tedad <= 0 AND @Status1 = 5 AND @Status2 = 7 
					UPDATE dbo.tFacD
					SET FeeUnit = X.BuyPrice		
	                        FROM (SELECT ISNULL(BuyPrice ,0) AS BuyPrice FROM  dbo.[tGood]
	                                        WHERE dbo.tGood.Code = @GoodCode  
	                                )X  
	                  	WHERE  dbo.tFacD.intSerialNo = @intSerialNo
	                                AND dbo.tFacD.Branch = @Branch
	                                AND dbo.tFacD.GoodCode = @GoodCode             		      

                    SET @NumberOfRecords = @NumberOfRecords + 1
                    FETCH NEXT FROM Havale INTO @Branch ,@intSerialNo,@fDate --, @GoodCode 
	
                END
	
            CLOSE Havale
            DEALLOCATE Havale
           
	FETCH NEXT  FROM GoodsList INTO @GoodCode , @BuyPrice

        END
    CLOSE GoodsList
    DEALLOCATE GoodsList
END

ELSE IF  dbo.Get_ISFIFO() = 1
BEGIN

    SET  @NumberOfRecords = 0			
    DECLARE  GoodsList CURSOR	 
    FOR 

 SELECT DISTINCT T2.GoodCode , dbo.tGood.BuyPrice FROM 
( SELECT DISTINCT T1.GoodCode  FROM 
(   SELECT  DISTINCT   ISNULL(tUsePercent.GoodFirstCode , tFacD.GoodCode ) AS GoodCode

    FROM    dbo.tFacM
    INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		AND [dbo].[tFacM].Branch = dbo.tFacD.Branch
		LEFT OUTER  JOIN dbo.tUsePercent ON dbo.tUsePercent.GoodCode = dbo.tFacD.GoodCode
    WHERE   dbo.tFacM.Status IN ( 1, 2 ,3, 4,5 , 6, 7 )
    AND tFacM.AccountYear = @AccountYear
    AND tFacM.Recursive = 0
    AND [dbo].[tFacD].intInventoryNo = @InventoryNo
    AND (dbo.tFacD.GoodCode = @GoodCode OR @GoodCode = 0)  -- For One Good(FrmGoodTurnOver) or AllGood(FrmFinalPrice) 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate
UNION all
    SELECT  [GoodCode]
	FROM      tInventory_Good
	WHERE     (tInventory_Good.GoodCode = @GoodCode OR @GoodCode = 0)
	AND dbo.tInventory_Good.AccountYear = @AccountYear
	AND [InventoryNo] = @InventoryNo AND tInventory_Good.FirstMojodi <> 0
) T1
GROUP BY GoodCode 
)T2
INNER JOIN dbo.tGood ON T2.GoodCode = dbo.tGood.Code


	
    OPEN GoodsList
    FETCH FROM GoodsList INTO @GoodCode , @BuyPrice

    WHILE @@FETCH_STATUS = 0 
        BEGIN
            --DECLARE @fTime NVARCHAR(8)
            DECLARE Havale CURSOR 
            FOR 
            SELECT DISTINCT tFacM.Branch,tFacM.intSerialNo,[Date] , Amount --, GoodCode  
            FROM [dbo].[tFacM]
            INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
            WHERE [Status] IN( 6,7) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
            AND GoodCode = @GoodCode --AND (GoodCode = @GoodCode OR @GoodCode = 0) 
            AND tFacM.AccountYear = @AccountYear 
            AND dbo.tFacM.Date <=@AfterDate-- N'88/06/31'  --*****************
            AND dbo.tFacM.Date>=@BeforeDate
            ORDER BY [Date] ASC , [dbo].[tFacM].intSerialNo ASC  

            OPEN Havale
	
            FETCH  FROM Havale INTO @Branch ,@intSerialNo,@fDate , @GoodAmount --,@GoodCode 
	
            WHILE @@FETCH_STATUS = 0 
                BEGIN

					CREATE TABLE #TmpFactorList
						(
						  intSerialNo BIGINT,
						  Amount FLOAT ,
						  FeeUnit FLOAT
						)

					INSERT INTO #TmpFactorList
							( intSerialNo, Amount, FeeUnit )
					SELECT 0 , FirstMojodi , FirstPrice FROM dbo.tInventory_Good WHERE AccountYear = @AccountYear AND GoodCode = @GoodCode AND [InventoryNo] = @InventoryNo

					INSERT INTO #TmpFactorList
							( intSerialNo, Amount, FeeUnit )
					SELECT  TF.intSerialNo ,
							SUM(TD.Amount) AS Amount ,
							MIN(TD.FeeUnit) AS FeeUnit
					FROM    ( SELECT    *
							  FROM      tFacM
							  WHERE     intSerialNo < @IntSerialNo
										AND Date <= @FDate
										AND Date >= @BeforeDate
										AND Date <= @AfterDate
										AND Status IN ( 1, 7 )
										AND dbo.tFacM.Recursive = 0
                                        
										
							) TF
							INNER JOIN tFacD TD ON TD.Branch = TF.Branch
												   AND TD.intSerialNo = TF.intSerialNo
												   AND TD.GoodCode = @GoodCode
												   AND [intInventoryNo] = @InventoryNo
					GROUP BY TF.intSerialNo
         

					-- نمایش فاکتورهای
					--  رسید انتقالی به انبار و خرید
					--SELECT  * FROM #TmpFactorList

					DECLARE @SumOut FLOAT

					SELECT  @SumOut = SUM(TD.Amount)
					FROM    ( SELECT    *
							  FROM      tFacM
							  WHERE     intSerialNo < @IntSerialNo
									   AND Date <= @FDate
										AND Date >= @BeforeDate
										AND Date <= @AfterDate
										AND Status IN ( 4, 6 )
										AND Recursive = 0
										
							) TF
							INNER JOIN tFacD TD ON TD.Branch = TF.Branch
												   AND TD.intSerialNo = TF.intSerialNo
												   AND TD.GoodCode = @GoodCode
												   AND [intInventoryNo] = @InventoryNo


					SET @SumOut = ISNULL(@SumOut, 0)

					CREATE TABLE #TmpAvg
						(
						  Amount FLOAT ,
						  FeeUnit FLOAT
						)

					DECLARE @Amount FLOAT
					DECLARE @FeeUnit FLOAT

					DECLARE InvoiceSerials CURSOR
					FOR
						SELECT  Amount ,
								FeeUnit
						FROM    #TmpFactorList

					OPEN InvoiceSerials
					FETCH NEXT FROM InvoiceSerials INTO @Amount, @FeeUnit
					WHILE @@FETCH_STATUS = 0
						BEGIN
							
							DECLARE @AmountUsed FLOAT
							  
							IF @GoodAmount <= 0
								BREAK		  
							  
							ELSE
								IF @SumOut > 0
									BEGIN
							  
										IF @Amount > @SumOut
											SET @AmountUsed = @SumOut
										ELSE
											SET @AmountUsed = @Amount
								 
										SET @SumOut = @SumOut - @AmountUsed
							  
									SET @Amount = @Amount - @AmountUsed
							  
									END
							              
								IF @SumOut <= 0 AND @GoodAmount > 0
										BEGIN

											IF @Amount > @GoodAmount
												SET @AmountUsed = @GoodAmount
											ELSE
												SET @AmountUsed = @Amount
								 
											SET @GoodAmount = @GoodAmount - @AmountUsed

											INSERT  INTO #TmpAvg
													( Amount, FeeUnit )
											VALUES  ( @AmountUsed, -- Amount - float
													  @FeeUnit  -- FeeUnit - float
													  )
							  
							  
										END
							  
							FETCH NEXT FROM InvoiceSerials INTO @Amount, @FeeUnit
						END
					CLOSE InvoiceSerials
					DEALLOCATE InvoiceSerials

					DECLARE @SumAmount FLOAT
					DECLARE @SumFee FLOAT

					SET @SumAmount = 0
					SET @SumFee = 0

					--SELECT * FROM #TmpAvg

					SELECT  @SumFee = @SumFee + ( Amount * FeeUnit ) ,
							@SumAmount = @SumAmount + Amount
					FROM    #TmpAvg

					--IF @SumAmount <> 0 SELECT  @SumFee / @SumAmount
					PRINT @SumFee
					PRINT @SumAmount
					
					DROP TABLE #TmpFactorList
					DROP TABLE #TmpAvg

                    IF @SumAmount <= 0 
                    	SET @priceTamam = @BuyPrice
             	    ELSE
           				SET @priceTamam = CAST(((@SumFee)/@SumAmount) AS INT)
             	        
                    SET @Status1 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1)         
                    SET @Status2 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo)         
                    SET @HavaleNO = ISNULL((SELECT RefrenceHavale FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1 AND Status = 6) , 0)         
--                    PRINT @GoodCode
--                     PRINT @Mablagh
--                     PRINT @Tedad
--                     PRINT @priceTamam
--                    PRINT @intSerialNo
--		    PRINT @NumberOfRecords + 1
                    --PRINT @HavaleNO
                    IF @priceTamam >= 0 
                        UPDATE  dbo.tFacD
                        SET     FeeUnit = @priceTamam
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
                                AND intSerialNo <> @HavaleNo -- we don,t need update Resid  from Havale
                    IF @priceTamam < 0 			--Negative Price set to Zero
                        UPDATE  dbo.tFacD
                        SET     FeeUnit = 0
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
                                AND intSerialNo <> @HavaleNo -- we don,t need update Resid  from Havale

--Update Resid From Havale With Havale Fee
			IF @intSerialNo = @HavaleNo
				UPDATE dbo.tFacD
				SET FeeUnit = X.feeUnit		
	                        FROM (SELECT feeUnit FROM  dbo.[tFacM]
	                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
	                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
	                                        WHERE   tFacM.Status = 6
	                                        AND dbo.tFacM.intSerialNo = @intSerialNo -1
	                                        AND dbo.tFacM.Branch = @Branch
	                                        AND dbo.tFacM.AccountYear = @AccountYear
	                                        AND dbo.tFacM.Recursive = 0
	                                        AND dbo.tFacD.GoodCode = @GoodCode  
	                                       -- AND [intInventoryNo] = @InventoryNo  --No Inventory Needed because is resid from other inventory
	                                )X  
	                  	WHERE  dbo.tFacD.intSerialNo = @intSerialNo
	                                AND dbo.tFacD.Branch = @Branch
	                                AND dbo.tFacD.GoodCode = @GoodCode             		      
			
--Update Resid when Mojodi is zero or negative
				IF @Tedad <= 0 AND @Status1 = 5 AND @Status2 = 7 
					UPDATE dbo.tFacD
					SET FeeUnit = X.BuyPrice		
	                        FROM (SELECT ISNULL(BuyPrice ,0) AS BuyPrice FROM  dbo.[tGood]
	                                        WHERE dbo.tGood.Code = @GoodCode  
	                                )X  
	                  	WHERE  dbo.tFacD.intSerialNo = @intSerialNo
	                                AND dbo.tFacD.Branch = @Branch
	                                AND dbo.tFacD.GoodCode = @GoodCode             		      

                    SET @NumberOfRecords = @NumberOfRecords + 1
                    FETCH NEXT FROM Havale INTO @Branch ,@intSerialNo,@fDate , @GoodAmount --, @GoodCode 
	
                END
	
            CLOSE Havale
            DEALLOCATE Havale
           
	FETCH NEXT  FROM GoodsList INTO @GoodCode , @BuyPrice

        END
    CLOSE GoodsList
    DEALLOCATE GoodsList
END
--=====================================================
IF @Flag = 1
PRINT '***********'
	EXEC dbo.Update_FacDFinalPrice 
	    @InventoryNo ,
	    @AccountYear, -- smallint
	    @GoodCode1 , -- int
	    @Flag , -- int
	    @BeforeDate , -- nvarchar(8)
	    @AfterDate , -- nvarchar(8)
	    @NumberOfRecords  -- int
PRINT '***********'

	RETURN @NumberOfRecords


IF @@ERROR <> 0
    AND @@TRANCOUNT > 0 
    ROLLBACK TRANSACTION ;



GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER   PROCEDURE [dbo].[Update_FacDFinalPrice]
    (
	@InventoryNo INT ,
	@AccountYear SMALLINT ,
	@GoodCode INT ,
	@Flag INT ,
	@BeforeDate NVARCHAR(8),
	@AfterDate NVARCHAR(8),
	@NumberOfRecords INT OUT 
    )
AS 


--PRINT '###########*'

DECLARE @BuyPrice INT 
IF @Flag = 0 
    SELECT @NumberOfRecords = ISNULL(COUNT(GoodCode), 0)  --, [TIME] 
    FROM [dbo].[tFacM]
    INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
    WHERE [Status] IN( 2,5) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
	AND (GoodCode = @GoodCode OR @GoodCode = 0) 
	AND tFacM.AccountYear = @AccountYear 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate

ELSE 
BEGIN

    SET  @NumberOfRecords = 0			
    DECLARE  GoodsList CURSOR	 
    FOR 

 SELECT DISTINCT T2.GoodCode , dbo.tGood.BuyPrice FROM 
( SELECT DISTINCT  T1.GoodCode FROM 
(   SELECT  DISTINCT      [GoodCode]

    FROM    dbo.tFacM
    INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		AND [dbo].[tFacM].Branch = dbo.tFacD.Branch
    WHERE   dbo.tFacM.Status IN ( 1, 2 ,3, 4,5 , 6, 7 )
    AND tFacM.AccountYear = @AccountYear
    AND tFacM.Recursive = 0
    AND [dbo].[tFacD].intInventoryNo = @InventoryNo
    AND (GoodCode = @GoodCode OR @GoodCode = 0)  -- For One Good(FrmGoodTurnOver) or AllGood(FrmFinalPrice) 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate
UNION all
    SELECT  [GoodCode]
	FROM      tInventory_Good
	WHERE     (tInventory_Good.GoodCode = @GoodCode OR @GoodCode = 0)
	AND dbo.tInventory_Good.AccountYear = @AccountYear
	AND [InventoryNo] = @InventoryNo AND tInventory_Good.FirstMojodi <> 0
) T1
GROUP BY GoodCode 
)T2
INNER JOIN dbo.tGood ON T2.GoodCode = dbo.tGood.Code

	
    OPEN GoodsList
    FETCH FROM GoodsList INTO @GoodCode , @BuyPrice

    WHILE @@FETCH_STATUS = 0 
        BEGIN
            DECLARE @intSerialNo INT
	    DECLARE @Branch INT 
            DECLARE @fDate NVARCHAR(8)
            --DECLARE @fTime NVARCHAR(8)
            DECLARE Havale CURSOR 
            FOR 
            SELECT DISTINCT tFacM.Branch,tFacM.intSerialNo,[Date] --, GoodCode  
            FROM [dbo].[tFacM]
            INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
            WHERE [Status] IN( 2,5) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
            AND GoodCode = @GoodCode --AND (GoodCode = @GoodCode OR @GoodCode = 0) 
            AND tFacM.AccountYear = @AccountYear 
            AND dbo.tFacM.Date <=@AfterDate-- N'88/06/31'  --*****************
            AND dbo.tFacM.Date>=@BeforeDate
            ORDER BY [Date] ASC , [dbo].[tFacM].intSerialNo ASC  

            OPEN Havale
	
            FETCH  FROM Havale INTO @Branch ,@intSerialNo,@fDate --,@GoodCode 
	
            WHILE @@FETCH_STATUS = 0 
                BEGIN
                    DECLARE @priceTamam INT ;
                    DECLARE @Mablagh BIGINT ;
                    DECLARE @Tedad INT ;
                    --SELECT @Tedad = ISNULL(FirstMojodi , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    --SELECT @Mablagh = ISNULL(FirstPrice , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    SELECT  @Mablagh = SUM(T.FirstMojodi * T.FirstPrice) + SUM(T.Amount * T.Flag * T.FeeUnit) ,                                          
                            @Tedad = SUM(T.FirstMojodi) + Sum(T.Amount * T.Flag)
                    FROM    (
                              SELECT    dbo.tInventory_Good.FirstMojodi ,
                                        dbo.tInventory_Good.FirstPrice ,
                                        tInventory_Good.GoodCode ,
                                        0 AS Amount ,
                                        0 AS Flag ,
                                        0 AS FeeUnit 
                                        FROM      tInventory_Good
                              WHERE     tInventory_Good.GoodCode = @GoodCode
                                        AND dbo.tInventory_Good.AccountYear = @AccountYear
                                        AND [InventoryNo] = @InventoryNo
                              UNION ALL
                              SELECT    0 AS FirstMojodi ,
                              		0 AS FirstPrice ,
                              		Goodcode ,
                                        Amount ,
                                        Flag ,
                                        FeeUnit 
                              FROM      dbo.[tFacM]
                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
                                        INNER JOIN dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo
                              WHERE     dbo.tFacM.[Date] <= @fDate 
                              		--( dbo.tFacM.[Date] + ' ' + dbo.tFacM.[Time] ) <= ( @fDate + ' ' + @fTime )
                                        AND tFacM.Status IN (  6, 7) --, 1, 3, 4 , 6, 7
                                        AND (dbo.tFacM.intSerialNo < @intSerialNo OR status = 1 OR status = 4)
                                        AND dbo.tFacM.Branch = @Branch
                                        AND dbo.tFacM.AccountYear = @AccountYear
                                        AND tFacD.GoodCode = @GoodCode
                                        AND dbo.tFacM.Recursive = 0
                                        AND dbo.tfacD.[intInventoryNo] = @InventoryNo
					AND dbo.tFacM.Date <=@AfterDate
					AND dbo.tFacM.Date>=@BeforeDate
                            ) T
                    GROUP BY GoodCode 

	DECLARE @SerialHavale INT 
	SELECT @SerialHavale = RefrenceHavale FROM tfacM WHERE intSerialNo  = @intSerialNo 
	SET @SerialHavale = ISNULL(@SerialHavale , 0)   

	--If Good is Analytic 
	DECLARE @GoodAmount FLOAT 
	SELECT @GoodAmount = Amount FROM dbo.tFacD WHERE intSerialNo  = @intSerialNo 
	DECLARE @MablaghHavale BIGINT 
	SET @MablaghHavale = (SELECT SUM(Amount * FeeUnit) FROM tfacD
	WHERE intSerialNo = @SerialHavale AND GoodCode IN (SELECT GoodFirstCode FROM dbo.tUsePercent WHERE GoodCode = @GoodCode) )     	        

	PRINT @GoodCode 
	PRINT   @GoodAmount          	           	
	PRINT   @SerialHavale          	           	
	PRINT   @MablaghHavale          	           	

	IF  @MablaghHavale > 0 
	BEGIN 
	SET @Tedad = @Tedad + @GoodAmount
	SET @Mablagh = @Mablagh + @MablaghHavale
	END 

                    IF @Tedad <= 0 
                    	SET @priceTamam = @BuyPrice
             	    ELSE
             	        SET @priceTamam = CAST((@Mablagh/@Tedad) AS INT)
             	        
--PRINT @GoodCode 
--PRINT   @priceTamam          	           	
--PRINT @NumberOfRecords
 --                   IF @priceTamam >= 0 
                        UPDATE  dbo.tFacD
                        SET     FinalPrice = @priceTamam
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
--                     IF @priceTamam < 0 			--Negative Price set to Zero
--                         UPDATE  dbo.tFacD
--                         SET     FinalPrice = 0
--                         WHERE   dbo.tFacD.intSerialNo = @intSerialNo
--                                 AND dbo.tFacD.Branch = @Branch
--                                 AND dbo.tFacD.GoodCode = @GoodCode
			
                    SET @NumberOfRecords = @NumberOfRecords + 1
                    FETCH NEXT FROM Havale INTO @Branch ,@intSerialNo,@fDate --, @GoodCode 
	
                END
	
            CLOSE Havale
            DEALLOCATE Havale
           
	FETCH NEXT  FROM GoodsList INTO @GoodCode , @BuyPrice

        END
    CLOSE GoodsList
    DEALLOCATE GoodsList
--=====================================================

	END
	RETURN @NumberOfRecords


IF @@ERROR <> 0
    AND @@TRANCOUNT > 0 
    ROLLBACK TRANSACTION ;


GO
