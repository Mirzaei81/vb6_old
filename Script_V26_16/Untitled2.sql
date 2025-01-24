--ScriptV26_16_Fix_17_950615.sql

--For Sql 2008 

--«÷«›Â ‘œ‰ »—ê‘  «“ ›—Ê‘ »Â ”Ì” „
--ò‰ —· «‰»«— Ê ﬁÌ„   „«„ ‘œÂ
--”Êœ Ê “Ì«‰ »«“—ê«‰Ì »« «” ›«œÂ «“ ﬁÌ„   „«„ ‘œÂ
--”Êœ Ê “Ì«‰ ò·Ì »« «” ›«œÂ «“ ﬁÌ„   „«„ ‘œÂ
--ê—› ‰ „«·Ì«  œ” Ì —ÊÌ ›«ò Ê—
--—”„Ì ò—œ‰ ›«ò Ê— ›—Ê‘ Ê Œ—Ìœ

--95/06/15

IF NOT EXISTS(SELECT * FROM tblPub_Script2 WHERE [Version] = 26 AND Script = 16 AND FixNumber = 17 )

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
			  17
			)
GO

IF COL_LENGTH('tFacM','Rasmi') IS NULL
BEGIN

	ALTER TABLE dbo.tFacM
	ADD Rasmi [bit] NULL DEFAULT(0)
END

GO

IF NOT EXISTS(SELECT * FROM dbo.tObjects WHERE intObjectCode = 371 )

INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 371 , -- intObjectCode - int
          N'SaleReturn' , -- ObjectId - nvarchar(50)
          N'»—ê‘  «“ ›—Ê‘' , -- ObjectName - nvarchar(50)
          N'SaleReturn' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          126  -- ObjectParent - int
        )
        
GO
IF NOT EXISTS(SELECT * FROM dbo.tAccess_Object WHERE intAccessLevel = 1 AND intObjectCode = 371 )
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          371  -- intObjectCode - int
          )
GO


-- «÷«›Â ò—œ‰ ò«·«Â«Ì ›—ÊŒ ‰Ì »Â Å—Ê”ÌÃ— ”Êœ Ê “Ì«‰
--«÷«›Â ‘œ‰ ÷—Ì» „’—› »—«Ì „Õ«”»Â ÕÊ«·Â ò«·«
--«÷«›Â ‘œ‰ „Õ«”»Â ﬁÌ„   „«„ ‘œÂ ò«·«Ì ›—ÊŒ ‰Ì
--
--95/06/13

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Get_Benefit_Loss]
    (
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @AccountYear SMALLINT ,
      @InventoryNo INT ,
      @GoodLevel1 INT ,
      @SelectedLevelsString NVARCHAR(4000)
    )
AS 
     BEGIN
--SET NOCOUNT ON added to prevent extra result sets FROM interfering with SELECT statements.
SET NOCOUNT ON ;
      
-- DECLARE @DateBefore NVARCHAR(50) ;
-- DECLARE @DateAfter NVARCHAR(50) ;
-- DECLARE @AccountYear SMALLINT ;
-- DECLARE @Branch INT ;
-- DECLARE @InventoryNo INT ;
-- DECLARE      @GoodLevel1 INT 
-- DECLARE      @SelectedLevelsString NVARCHAR(4000)
-- 
-- SELECT  @DateBefore = N'88/01/01' ;
-- SELECT  @DateAfter = N'88/12/30' ;
-- SELECT  @AccountYear = 1388 ;
-- SELECT  @Branch = 1 ;
-- SELECT  @InventoryNo = 100 ;
-- SELECT @GoodLevel1 = -1
-- SELECT @SelectedLevelsString = N''
DECLARE @SaleDiscountTotal BIGINT
DECLARE @DiscountFacD BIGINT
SELECT @SaleDiscountTotal = SUM(ISNULL(DiscountTotal , 0))  
	                      FROM      [dbo].[tFacM] 
	                        WHERE     [tFacM].[Date] >= @DateBefore
                                AND [tFacM].[Date] <= @DateAfter
                                AND [tFacM].[AccountYear] = @AccountYear
                                AND [tFacM].[Recursive] = 0
				AND tFacM.Status = 2
				AND dbo.tFacM.intSerialNo IN 
				(
				SELECT intSerialNo FROM dbo.tFacD 
				INNER JOIN dbo.vw_Good ON dbo.tFacD.GoodCode = dbo.vw_Good.Code
                                WHERE  [tFacD].[intInventoryNo] = @InventoryNo
		                AND ( [vw_Good].[Level1] = @GoodLevel1
		                      OR @GoodLevel1 = -1
		                    )
		                AND ( [vw_Good].[Level2] IN (
		                      SELECT    CAST(Word AS INT)
		                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
		                                                              N',') )
		                      OR @SelectedLevelsString = N''
		                    )
				)

	SELECT 	@DiscountFacD = SUM(( [tFacD].[Amount] * [tFacD].[FeeUnit] ) * ( [tFacD].[Discount] / 100 )) 
                      FROM      [dbo].[tFacM] 
                                INNER JOIN [dbo].[tFacD]  ON [tFacM].[Branch] = [tFacD].[Branch]
                                                      AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
				INNER JOIN [vw_Good] ON vw_Good.Code = tFacD.GoodCode
	                        WHERE     [tFacM].[Date] >= @DateBefore
                                AND [tFacM].[Date] <= @DateAfter
                                AND [tFacM].[AccountYear] = @AccountYear
                                AND [tFacM].[Recursive] = 0
                                AND [tFacD].[intInventoryNo] = @InventoryNo
				AND tFacM.Status = 2
		                AND ( [vw_Good].[Level1] = @GoodLevel1
		                      OR @GoodLevel1 = -1
		                    )
		                AND ( [vw_Good].[Level2] IN (
		                      SELECT    CAST(Word AS INT)
		                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
		                                                              N',') )
		                      OR @SelectedLevelsString = N''
		                    )

	SET @SaleDiscountTotal = @SaleDiscountTotal - @DiscountFacD
DECLARE @BuyDiscountTotal BIGINT
DECLARE @BuyDiscountFacD BIGINT
SELECT @BuyDiscountTotal = SUM(ISNULL(DiscountTotal , 0))  
	                      FROM      [dbo].[tFacM] 
	                        WHERE     [tFacM].[Date] >= @DateBefore
                                AND [tFacM].[Date] <= @DateAfter
                                AND [tFacM].[AccountYear] = @AccountYear
                                AND [tFacM].[Recursive] = 0
				AND tFacM.Status = 1
				AND dbo.tFacM.intSerialNo IN 
				(
				SELECT intSerialNo FROM dbo.tFacD 
				INNER JOIN dbo.vw_Good ON dbo.tFacD.GoodCode = dbo.vw_Good.Code
                                WHERE  [tFacD].[intInventoryNo] = @InventoryNo
		                AND ( [vw_Good].[Level1] = @GoodLevel1
		                      OR @GoodLevel1 = -1
		                    )
		                AND ( [vw_Good].[Level2] IN (
		                      SELECT    CAST(Word AS INT)
		                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
		                                                              N',') )
		                      OR @SelectedLevelsString = N''
		                    )
				)

	SELECT 	@BuyDiscountFacD = SUM(( [tFacD].[Amount] * [tFacD].[FeeUnit] ) * ( [tFacD].[Discount] / 100 )) 
                      FROM      [dbo].[tFacM] 
                                INNER JOIN [dbo].[tFacD]  ON [tFacM].[Branch] = [tFacD].[Branch]
                                                      AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
				INNER JOIN [vw_Good] ON vw_Good.Code = tFacD.GoodCode
	                        WHERE     [tFacM].[Date] >= @DateBefore
                                AND [tFacM].[Date] <= @DateAfter
                                AND [tFacM].[AccountYear] = @AccountYear
                                AND [tFacM].[Recursive] = 0
                                AND [tFacD].[intInventoryNo] = @InventoryNo
						AND tFacM.Status = 1
		                AND ( [vw_Good].[Level1] = @GoodLevel1
		                      OR @GoodLevel1 = -1
		                    )
		                AND ( [vw_Good].[Level2] IN (
		                      SELECT    CAST(Word AS INT)
		                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
		                                                              N',') )
		                      OR @SelectedLevelsString = N''
		                    )

	SET @BuyDiscountTotal = @BuyDiscountTotal - @BuyDiscountFacD

        SELECT  [tInventory_Good].GoodCode  , 
                [dbo].[vw_Good].[Name] ,
                [dbo].[vw_Good].[BarCode] ,
				FirstMojodi , FirstPrice ,
				CAST([dbo].[tInventory_Good].[FirstMojodi]
                  * [dbo].[tInventory_Good].[FirstPrice] AS BIGINT) AS TotalFirstPrice ,
            ISNULL(CAST(T3.TotalBuyAmount AS BIGINT), 0) AS TotalBuyAmount ,
            ISNULL(CAST(T3.TotalBuyReturnAmount AS BIGINT), 0) AS TotalBuyReturnAmount ,
            ISNULL(CAST(T3.TotalLossAmount AS BIGINT), 0) AS TotalLossAmount ,
            ISNULL(CAST(T3.TotalHavalehAmount AS BIGINT), 0) AS TotalHavalehAmount ,
            ISNULL(CAST(T3.TotalResidAmount AS BIGINT), 0) AS TotalResidAmount ,
            Mojodi , [dbo].[tInventory_Good].[MojodiPrice] ,
            CAST([dbo].[tInventory_Good].[Mojodi]
                  * [dbo].[tInventory_Good].[MojodiPrice] AS BIGINT) AS TotalMojodiPrice ,
            ISNULL(CAST(T3.TotalSellAmount AS BIGINT), 0) AS TotalSellAmount ,
            ISNULL(CAST(T3.TotalFinalAmount AS BIGINT), 0) AS TotalFinalAmount ,
            ISNULL(CAST(T3.TotalSellReturnAmount AS BIGINT), 0) AS TotalSellReturnAmount ,
            ISNULL(CAST(T3.TotalFinalReturnAmount AS BIGINT), 0) AS TotalFinalReturnAmount ,
		(ISNULL(CAST(T3.TotalSellAmount AS BIGINT), 0) - ISNULL(CAST(T3.TotalFinalReturnAmount AS BIGINT), 0) ) -
		(ISNULL(CAST(T3.TotalFinalAmount AS BIGINT), 0) - ISNULL(CAST(T3.TotalFinalReturnAmount AS BIGINT), 0)) AS GoodBenefitLoss ,
		@SaleDiscountTotal AS SaleDiscountTotal ,@DiscountFacD AS DiscountFacD ,
		
                [dbo].[vw_Good].[TechnicalNo] ,
                [dbo].[vw_Good].[Unit] ,
                [dbo].[vw_Good].[UnitDescription] ,
                CAST([dbo].[tInventory_Good].[FirstMojodi]
                  * [dbo].[tInventory_Good].[FirstPrice] +  ISNULL(T3.TotalBuyAmount, 0) - ISNULL(T3.TotalLossAmount, 0)
				- ISNULL(T3.TotalBuyReturnAmount, 0) - ISNULL(T3.TotalHavalehAmount, 0) +ISNULL(T3.TotalResidAmount, 0)  AS BIGINT) AS TotalMojodiPrice2 
				, @BuyDiscountTotal AS BuyDiscountTotal ,@BuyDiscountFacD AS BuyDiscountFacD 
                FROM    [dbo].[vw_Good]
                INNER JOIN [dbo].[tInventory_Good] ON [dbo].[vw_Good].[Code] = [dbo].[tInventory_Good].[GoodCode] 
                AND [dbo].[tInventory_Good].[InventoryNo] = @InventoryNo
                AND [dbo].[tInventory_Good].[AccountYear] = @AccountYear

	FULL OUTER JOIN 
		(
                  SELECT    GoodCode ,
                            ISNULL(SUM(TotalBuyAmount), 0) AS TotalBuyAmount ,
                            ISNULL(SUM(TotalSellAmount), 0) AS TotalSellAmount ,
                            ISNULL(SUM(TotalLossAmount), 0) AS TotalLossAmount ,
                            ISNULL(SUM(TotalBuyReturnAmount), 0) AS TotalBuyReturnAmount ,
                            ISNULL(SUM(TotalSellReturnAmount), 0) AS TotalSellReturnAmount ,
                            ISNULL(SUM(TotalHavalehAmount), 0) AS TotalHavalehAmount ,
                            ISNULL(SUM(TotalResidAmount), 0) AS TotalResidAmount ,
                            ISNULL(SUM(TotalFinalAmount), 0) AS TotalFinalAmount ,
                            ISNULL(SUM(TotalFinalReturnAmount), 0) AS TotalFinalReturnAmount
                  FROM      (
                              SELECT    [D].[GoodCode] ,
                                        CASE WHEN [M].Status = 1
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalBuyAmount ,
                                        CASE WHEN [M].Status = 2
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalSellAmount ,
                                        CASE WHEN [M].Status = 2
                                             THEN SUM([D].[Amount]
                                                      * [D].[FinalPrice])
                                             ELSE 0
                                        END AS TotalFinalAmount ,
                                        CASE WHEN [M].Status = 3
                                             THEN SUM([D].[Amount]
                                                      * [D].[FeeUnit])
                                             ELSE 0
                                        END AS TotalLossAmount ,
                                        CASE WHEN [M].Status = 4
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalBuyReturnAmount ,
                                        CASE WHEN [M].Status = 5
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalSellReturnAmount ,
                                        CASE WHEN [M].Status = 5
                                             THEN SUM([D].[Amount]
                                                      * [D].[FinalPrice])
                                             ELSE 0
                                        END AS TotalFinalReturnAmount ,
                                        CASE WHEN [M].Status = 6
                                             THEN SUM([D].[Amount]
                                                      * [D].[FeeUnit])
                                             ELSE 0
                                        END AS TotalHavalehAmount ,
                                        CASE WHEN [M].Status = 7
                                             THEN SUM([D].[Amount]
                                                      * [D].[FeeUnit])
                                             ELSE 0
                                        END AS TotalResidAmount
                              FROM      [dbo].[tFacM] M
                                        INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                                              AND [M].[intSerialNo] = [D].[intSerialNo]
										INNER JOIN dbo.tGood ON [D].GoodCode = dbo.tGood.Code
                              WHERE     [M].[Date] >= @DateBefore
                                        AND [M].[Date] <= @DateAfter
                                        AND [M].[AccountYear] = @AccountYear
                                        AND [M].[Recursive] = 0
                                        AND [D].[intInventoryNo] = @InventoryNo
                                        AND (dbo.tGood.GoodType = 1 OR dbo.tGood.GoodType = 2 OR dbo.tGood.GoodType = 3)
                              GROUP BY  [D].[GoodCode] ,
                                        [M].[Status]
                            ) T1
                  GROUP BY  [T1].[GoodCode]
                ) T3
                ON [T3].[GoodCode] = [dbo].[tInventory_Good].[GoodCode]
        	WHERE   [dbo].[tInventory_Good].[InventoryNo] = @InventoryNo
                AND [dbo].[tInventory_Good].[AccountYear] = @AccountYear
                AND (vw_Good.GoodType = 1 OR vw_Good.GoodType = 2 OR vw_Good.GoodType = 3)
                AND ( [vw_Good].[Level1] = @GoodLevel1
                      OR @GoodLevel1 = -1
                    )
                AND ( [vw_Good].[Level2] IN (
                      SELECT    CAST(Word AS INT)
                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
                                                              N',') )
                      OR @SelectedLevelsString = N''
                    )
        ORDER BY [tInventory_Good].[GoodCode] ASC
    END
--===============================================




GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO
ALTER PROCEDURE [dbo].[Update_HavalehResid]
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
-- DECLARE @GoodCode INT
-- SELECT  @GoodCode = 4
-- 
-- DECLARE @InventoryNo INT 
-- DECLARE @Branch INT 
-- DECLARE @AccountYear SMALLINT
-- -- 
-- SELECT  @InventoryNo = 100
-- SELECT  @Branch = 1
-- SELECT  @AccountYear = 1389
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

ELSE 
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
            DECLARE  @intSerialNo INT
            DECLARE @Branch INT 
	        DECLARE @fDate NVARCHAR(8)
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
                                        AND tFacM.Status IN ( 1, 3, 4 , 6 , 7 ) -- , 6, 7  »—«Ì „ÊÃÊœÌ „‰›Ì
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
             	        
                    DECLARE @Status1 INT 
                    SET @Status1 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1)         
                    DECLARE @Status2 INT 
                    SET @Status2 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo)         
                    DECLARE @HavaleNo INT 
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
--=====================================================
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
	END
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

-- DECLARE @GoodCode INT
-- SELECT  @GoodCode = 4
-- 
-- DECLARE @InventoryNo INT 
-- DECLARE @Branch INT 
-- DECLARE @AccountYear SMALLINT
-- -- 
-- SELECT  @InventoryNo = 100
-- SELECT  @Branch = 1
-- SELECT  @AccountYear = 1389
-- 

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
                                        AND tFacM.Status IN ( 1, 3, 4 , 6, 7) --, 6, 7
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

	--If Good is Analytic 
	DECLARE @GoodAmount FLOAT 
	SELECT @GoodAmount = Amount FROM dbo.tFacD WHERE intSerialNo  = @intSerialNo 
	DECLARE @SerialHavale INT 
	SELECT @SerialHavale = RefrenceHavale FROM tfacM WHERE intSerialNo  = @intSerialNo 
	SET @SerialHavale = ISNULL(@SerialHavale , 0)   
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

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO



ALTER   PROCEDURE dbo.Update_tblTotal_tInventory_tGood_For_FinalPrice
(  
 @SystemDate   NVARCHAR(50),  
 @SystemDay    NVARCHAR(50),  
 @SystemTime   NVARCHAR(50),   
 @DateBefore   NVARCHAR(50),  
 @DateAfter    NVARCHAR(50),  
 @Type  int  ,  
 @InventoryNo Int ,  
 @AccountYear Smallint ,
 @ZeroNegative BIT 
)   
  
AS  
BEGIN TRAN  

	UPDATE dbo.tInventory_Good
	SET BuyPriceAverage = FirstPrice , MojodiPrice = FirstPrice
	WHERE InventoryNo = @InventoryNo AND AccountYear = @AccountYear
	
	UPDATE dbo.tInventory_Good
	SET BuyPriceAverage = T.AverageBuyFee , MojodiPrice = T.AverageBuyFee  FROM (
	Select (IsNull(Sum(FeeUnit * Amount) ,0) + ISNULL(FirstPrice * FirstMojodi , 0)) /(ISNULL(Sum(Amount),1) + ISNULL(FirstMojodi ,1)) AS AverageBuyFee , dbo.tInventory_Good.GoodCode , tInventory_Good.AccountYear , tInventory_Good.InventoryNo
--	Select IsNull(Sum(FeeUnit * Amount) ,0) /(ISNULL(Sum(Amount),1) ) AS AverageBuyFee , dbo.tInventory_Good.GoodCode , tInventory_Good.AccountYear , tInventory_Good.InventoryNo
	From tFacM inner join tfacd On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch AND dbo.tFacD.intInventoryNo = @InventoryNo
	INNER JOIN dbo.tInventory_Good ON tfacd.GoodCode = dbo.tInventory_Good.GoodCode 
		AND dbo.tInventory_Good.AccountYear = @AccountYear AND dbo.tInventory_Good.Branch = dbo.tFacD.Branch 
		AND dbo.tInventory_Good.InventoryNo = @InventoryNo
	Where tfacm.Status = 1 and Recursive = 0 And tfacm.AccountYear = @AccountYear AND tfacD.intInventoryNo = @InventoryNo 
	GROUP BY tfacd.GoodCode ,tInventory_Good.GoodCode, tInventory_Good.AccountYear , tInventory_Good.InventoryNo ,  tInventory_Good.FirstMojodi  ,  tInventory_Good.FirstPrice)T
	WHERE tInventory_Good.AccountYear = t.AccountYear  AND dbo.tInventory_Good.InventoryNo = t.InventoryNo AND tInventory_Good.GoodCode = t.GoodCode

UPDATE  tInventory_Good  
    
 Set    BuyAmount = T2.BuyAmount,  
		SaleAmount = T2.SaleAmount ,  
		LossAmount = T2.LossAmount ,
		BuyReturnAmount = T2.BuyReturnAmount ,  
		SaleReturnAmount = T2.SaleReturnAmount ,  
		FromStoreAmount = T2.FromStoreAmount ,  
		toStoreAmount = T2.toStoreAmount ,  
		Mojodi = T2.Mojodi , 
 	    MojodiPrice = CASE tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount  - tInventory_Good.FromStoreAmount + tInventory_Good.toStoreAmount  WHEN 0 THEN 0 ELSE ( firstMojodiRial + BuyRial - FromStoreRial + toStoreRial ) / (tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount  - tInventory_Good.FromStoreAmount + tInventory_Good.toStoreAmount) END 
   
 FROM dbo.tblTotal_tInventory_tGood_For_FinalPrice  
  (  
   @DateBefore   ,  
   @DateAfter    ,  
   @Type    ,  
   @InventoryNo  ,  
   @AccountYear  
  )  
   AS T2    
     Where tInventory_Good.GoodCode = T2.GoodCode And tInventory_Good.InventoryNo = T2.InventoryNo and tInventory_Good.AccountYear = @AccountYear  
	if @@Error <> 0   
	 goto ErrHandler  
  

 	  
UPDATE  tInventory_Good  
	SET MojodiPrice = 0 WHERE Mojodi = 0

IF @ZeroNegative = 1
UPDATE  tInventory_Good  
	SET MojodiPrice = 0 WHERE MojodiPrice < 0

Commit Tran   
  
return  
  
ErrHandler:  
RollBack Tran  
return 
  




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
 TotalFinalSellAmount BIGINT ,
 TotalFinalSellReturnAmount BIGINT ,
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

DECLARE  @nvcDate1 NVARCHAR(8) 
DECLARE  @nvcDate2 NVARCHAR(8) 
SET @nvcDate1 = SUBSTRING(CAST(@DateBefore AS NVARCHAR(10)) ,3 ,2) + '/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(10)) ,5 ,2) + '/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(10)) ,7 ,2)
SET @nvcDate2 = SUBSTRING(CAST(@DateAfter AS NVARCHAR(10)) ,3 ,2) + '/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(10)) ,5 ,2) + '/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(10)) ,7 ,2)
	DECLARE @TotalSellAmount BIGINT
	DECLARE @TotalSellReturnAmount BIGINT
	DECLARE @TotalFinalSellAmount BIGINT
	DECLARE @TotalFinalSellReturnAmount BIGINT

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
	

		--Select @TotalSellAmount = Sum(-Bedehkar + Bestankar) 
		--From tblAcc_DocumentDetail
		--INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		--Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		--AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		--AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		--Select @TotalSellReturnAmount = Sum(Bedehkar - Bestankar) 
		--From tblAcc_DocumentDetail
		--INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		--Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		--AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17)
		--AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)

		Select @TotalSellAmount = SUM( [D].[Amount] * [D].[FeeUnit] )
                                          
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
		
		Select @TotalSellReturnAmount = SUM( [D].[Amount] * [D].[FeeUnit] )
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 5
		
		Select @TotalFinalSellAmount = SUM( [D].[Amount] * [D].[FinalPrice] )
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
		
		Select @TotalFinalSellReturnAmount = SUM( [D].[Amount] * [D].[FinalPrice] )
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 5

		Select @TotalSaleDiscount =  SUM( [M].[DiscountTotal] )
        FROM      [dbo].[tFacM] M
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
			
		Select @TotalFirstPrice = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35)
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalBuyAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16)
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalBuyReturnAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18)
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalBuyDiscount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29)
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)

		Select @TotalLosses = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32)
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalHazine = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalHazineMali = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 36  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalHazineTozie = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 37  )
		AND MoeinId <> (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32) --Losses  moein code calculated in totallosses
		AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalCareeFee =  SUM( [M].[CarryFeeTotal] )
        FROM      [dbo].[tFacM] M
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
			
		--Select @TotalCareeFee = Sum(-Bedehkar + Bestankar) 
		--From tblAcc_DocumentDetail
		--INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		--Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		--AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4)
		--AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalPacking =  SUM( [M].[PackingTotal] )
        FROM      [dbo].[tFacM] M
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
			

		--Select @TotalPacking = Sum(-Bedehkar + Bestankar) 
		--From tblAcc_DocumentDetail
		--INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		--Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		--AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3)
		--AND (TafsiliId = @MarkazHazineh OR @MarkazHazineh = 0)
		
		Select @TotalPacking =  SUM( [M].[ServiceTotal] )
        FROM      [dbo].[tFacM] M
          WHERE     [M].[Date] >= @nvcDate1
                    AND [M].[Date] <= @nvcDate2
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
			
		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount  , TotalFinalSellAmount  , TotalFinalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,  Totalservice ,
			 TotalLosses  , TotalHoghough  , TotalHazine , TotalHazineMali , TotalHazineTozie )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFinalSellAmount  , @TotalFinalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  , @Totalservice ,
			 @TotalLosses  , @TotalHoghough  , @TotalHazine , @TotalHazineMali , @TotalHazineTozie)
		            

RETURN 


End

GO


--exec Get_TarazSoodZian 13950101,13950531,1395,1,0
--GO


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
	   ISNULL(TotalFinalSellAmount , 0) AS TotalFinalSellAmount ,
       ISNULL(TotalFinalSellReturnAmount , 0) AS TotalFinalSellReturnAmount ,
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


ALTER  PROC Update_FirstPriceByBuyPrice
	(
	@AccountYear INT,
	@Flag BIT,
	@InventoryNO INT
	)
AS
--delete inventoryno from calculate
--All Inventories affected by procedure


IF @Flag = 0
BEGIN 
		UPDATE tInventory_Good
		SET FirstPrice = ISNULL(T2.FeeUnit , 0)

		from (
		SELECT T.intserialNo , T.GoodCode , FeeUnit
		FROM tfacd INNER JOIN
		(
		SELECT MAX(dbo.tFacD.intSerialNo) AS intserialNo  , GoodCode  FROM tfacm
			INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
			WHERE Status = 1 AND AccountYear < dbo.Get_AccountYear()  AND tfacd.Branch = dbo.Get_Current_Branch()
			GROUP BY GoodCode
			)T
		ON T.GoodCode = dbo.tFacD.GoodCode AND T.intserialNo = dbo.tFacD.intSerialNo 
		--ORDER BY T.GoodCode
		) T2

		WHERE dbo.tInventory_Good.AccountYear = @AccountYear AND tInventory_Good.GoodCode = T2.GoodCode 
		AND T2.FeeUnit > 0

		-- if feeunit = 0 then replace firstprice with BuyPrice
		UPDATE dbo.tInventory_Good
		SET dbo.tInventory_Good.FirstPrice=dbo.tGood.BuyPrice
		FROM dbo.tGood 
			JOIN dbo.tInventory_Good ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
		WHERE dbo.tInventory_Good.FirstPrice = 0 AND 
			 dbo.tInventory_Good.AccountYear = @AccountYear
			--AND dbo.tInventory_Good.InventoryNo = @InventoryNO--ISNULL(@InventoryNO,dbo.tInventory_Good.InventoryNo)
			AND dbo.tGood.BuyPrice <> 0
END 
IF @Flag = 1
BEGIN 
		UPDATE tInventory_Good
		SET FirstPrice = ISNULL(T2.FeeUnit , 0)

		from (
		SELECT T.intserialNo , T.GoodCode , FeeUnit
		FROM tfacd INNER JOIN
		(
		SELECT MAX(dbo.tFacD.intSerialNo) AS intserialNo  , GoodCode  FROM tfacm
			INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
			WHERE Status = 1 AND AccountYear < dbo.Get_AccountYear()  AND tfacd.Branch = dbo.Get_Current_Branch()
			GROUP BY GoodCode
			)T
		ON T.GoodCode = dbo.tFacD.GoodCode AND T.intserialNo = dbo.tFacD.intSerialNo 
		--ORDER BY T.GoodCode
		) T2

		WHERE dbo.tInventory_Good.AccountYear = @AccountYear AND tInventory_Good.GoodCode = T2.GoodCode 
		AND tInventory_Good.FirstPrice = 0 AND T2.FeeUnit > 0

		-- if feeunit = 0 then replace firstprice with BuyPrice
		UPDATE dbo.tInventory_Good
		SET dbo.tInventory_Good.FirstPrice=dbo.tGood.BuyPrice
		FROM dbo.tGood 
			JOIN dbo.tInventory_Good ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
		WHERE dbo.tInventory_Good.FirstPrice = 0 AND 
			 dbo.tInventory_Good.AccountYear = @AccountYear
			--AND dbo.tInventory_Good.InventoryNo = @InventoryNO--ISNULL(@InventoryNO,dbo.tInventory_Good.InventoryNo)
			AND dbo.tGood.BuyPrice <> 0
				
END


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE dbo.Update_BuyPrice_by_LastPrice
AS 

UPDATE tgood
SET BuyPrice = T2.FeeUnit

from (
SELECT T.intserialNo , T.GoodCode , FeeUnit
FROM tfacd INNER JOIN
(
SELECT MAX(dbo.tFacD.intSerialNo) AS intserialNo  , GoodCode  FROM tfacm
    INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
    WHERE Status = 1 AND AccountYear = dbo.Get_AccountYear()  AND tfacd.Branch = dbo.Get_Current_Branch()
    GROUP BY GoodCode
    )T
ON T.GoodCode = dbo.tFacD.GoodCode AND T.intserialNo = dbo.tFacD.intSerialNo 
--ORDER BY T.GoodCode
) T2

WHERE tGood.code = T2.GoodCode AND T2.FeeUnit > 0

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS OFF
GO

ALTER Function [dbo].Fn_SoodZian_Sale

(
  @DateBefore NVARCHAR(8)  ,
  @DateAfter NVARCHAR(8)  ,
  @AccountYear SMALLINT ,
  @Branch INT 
)

RETURNS  @ReturnTable TABLE(
 TotalSellAmount BIGINT ,
 TotalSellReturnAmount BIGINT ,
 TotalFinalSellAmount BIGINT ,
 TotalFinalSellReturnAmount BIGINT ,
 TotalFirstPrice BIGINT ,
 TotalBuyAmount BIGINT ,
 TotalBuyReturnAmount BIGINT ,
 TotalSaleDiscount BIGINT ,
 TotalBuyDiscount BIGINT ,

 TotalCareeFee BIGINT ,
 TotalPacking BIGINT ,
 TotalService BIGINT ,
 TotalTax BIGINT ,

 TotalLosses BIGINT ,
 TotalHoghough BIGINT ,
 TotalHazineTolid BIGINT , 
 TotalHazineTax BIGINT 
)	
As

BEGIN


	DECLARE @TotalSellAmount BIGINT
	DECLARE @TotalSellReturnAmount BIGINT
	DECLARE @TotalFinalSellAmount BIGINT
	DECLARE @TotalFinalSellReturnAmount BIGINT
	DECLARE @TotalFirstPrice BIGINT
	DECLARE @TotalBuyAmount BIGINT
	DECLARE @TotalBuyReturnAmount BIGINT
	DECLARE @TotalSaleDiscount BIGINT
	DECLARE @TotalBuyDiscount BIGINT

	DECLARE @TotalCareeFee BIGINT
	DECLARE @TotalPacking BIGINT
	DECLARE @TotalService BIGINT
	DECLARE @TotalTax BIGINT

	DECLARE @TotalLosses BIGINT
	DECLARE @TotalHoghough BIGINT
	DECLARE @TotalHazineTolid BIGINT
	DECLARE @TotalHazineTax BIGINT
	

		Select @TotalSellAmount = SUM( [D].[Amount] * [D].[FeeUnit] ) FROM dbo.tFacM M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
		WHERE [M].Recursive = 0 AND [M].Status = 2 AND  [M].AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND [M].Branch = @Branch
		
		
		Select @TotalSellReturnAmount = SUM( [D].[Amount] * [D].[FeeUnit] )
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @DateBefore
                    AND [M].[Date] <= @DateAfter
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 5
		
		Select @TotalFinalSellAmount = SUM( [D].[Amount] * [D].[FinalPrice] )
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @DateBefore
                    AND [M].[Date] <= @DateAfter
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 2
		
		Select @TotalFinalSellReturnAmount = SUM( [D].[Amount] * [D].[FinalPrice] )
        FROM      [dbo].[tFacM] M
                    INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                          AND [M].[intSerialNo] = [D].[intSerialNo]
          WHERE     [M].[Date] >= @DateBefore
                    AND [M].[Date] <= @DateAfter
                    AND [M].[AccountYear] = @AccountYear
                    AND [M].[Branch] = @Branch
                    AND [M].[Recursive] = 0
                    AND [M].[Status] = 5

			
		Select @TotalFirstPrice =  SUM(FirstMojodi * FirstPrice) FROM dbo.tInventory_Good
		WHERE AccountYear = @AccountYear AND Branch = @Branch 

		Select @TotalBuyAmount = SUM(Sumprice) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 1 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalBuyReturnAmount = SUM(Sumprice) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 4 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalSaleDiscount = SUM(DiscountTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalBuyDiscount = SUM(DiscountTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 1 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalCareeFee = SUM(CarryFeeTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @TotalPacking = SUM(PackingTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		Select @Totalservice = SUM(ServiceTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
		Select @TotalTax = SUM(TaxTotal) + SUM(DutyTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			

		Select @TotalLosses = SUM(dbo.tgood.FinalPrice * Amount) FROM dbo.tFacM 
		INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
		WHERE Recursive = 0 AND Status = 3 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND dbo.tFacM.Branch = @Branch
			
			
		SET @TotalHoghough = 0
		--Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		--From tblAcc_DocumentDetail
		--INNER JOIN tblAcc_DocumentHeader ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
		--Where tblAcc_DocumentHeader.AccountYear = @AccountYear And tblAcc_DocumentHeader.Branch = @Branch AND tblAcc_DocumentHeader.DocumentDate >= @DateBefore AND tblAcc_DocumentHeader.DocumentDate <= @DateAfter
		--AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		----AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate


		Select @TotalHazineTolid = SUM(dbo.tgood.FinalPrice * Amount) FROM dbo.tFacM 
		INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
		WHERE Recursive = 0 AND Status = 2 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND dbo.tFacM.Branch = @Branch
			
			
		Select @TotalHazineTax = SUM(TaxTotal) + SUM(DutyTotal) FROM dbo.tFacM
		WHERE Recursive = 0 AND Status = 1 AND  AccountYear = @AccountYear AND Date >= @DateBefore AND Date <= @DateAfter
			AND Branch = @Branch
			
			
		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount , TotalFinalSellAmount  , TotalFinalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,  Totalservice ,
			 TotalTax ,TotalLosses  , TotalHoghough  , TotalHazineTolid , TotalHazineTax )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFinalSellAmount  , @TotalFinalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  , @Totalservice ,
			 @TotalTax ,@TotalLosses  , @TotalHoghough  , @TotalHazineTolid ,@TotalHazineTax)

		            
RETURN 


End




GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_TarazSoodZian_Sale]
    (
      @DateBefore NVARCHAR(8)  ,
      @DateAfter NVARCHAR(8)  ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS 


SELECT ISNULL(TotalSellAmount , 0)  AS TotalSellAmount ,
       ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
	   ISNULL(TotalFinalSellAmount , 0) AS TotalFinalSellAmount ,
       ISNULL(TotalFinalSellReturnAmount , 0) AS TotalFinalSellReturnAmount ,
       ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
       ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
       ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
       ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
       ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
       ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
       ISNULL(TotalPacking , 0) AS TotalPacking ,
       ISNULL(TotalService , 0) AS TotalService ,
       ISNULL(TotalTax , 0) AS TotalTax ,
       ISNULL(TotalLosses , 0) AS TotalLosses ,
       ISNULL(TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(TotalHazineTolid , 0) AS TotalHazineTolid ,
       ISNULL(TotalHazineTax , 0) AS TotalHazineTax 
       
	FROM dbo.Fn_SoodZian_Sale(@DateBefore ,@DateAfter ,@AccountYear ,@Branch )
--===============================================


GO


--exec Update_tblTotal_tInventory_tGood_For_Mojodi 0,N'95/06/15',N'œÊ ‘‰»Â',N'07:17',N'95/01/01',N'95/06/15',3,1,1,1,0,1395
--GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

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
						ELSE T1.Mojodi  END  AS Mojodi,
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
            Mojodi = 	 CASE WHEN @Type = 1
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


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
-- »—ÿ—› ò—œ‰ „‘ò· “Ì—
-- œ— ’Ê— Ì òÂ ”ÿ—Ì œ— ÃœÊ· „ÊÃÊœÌ «‰»«— «“ ò«·« ‰»«‘œ ‰„Ì  Ê«‰”  »— «”«” ›«ò Ê— Â« œ«œÂ „ÊÃÊœÌ —« »Ì«Ê—œ



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
            SET @TimeTitle = N' ”«⁄  : '
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
                                                --- ISNULL(SUM(W.SaleAmount), 0)
                                                - ISNULL(MAX(W.LossAmount), 0)
                                                - ISNULL(MAX(W.BuyReturnAmount),
                                                         0)
                                               -- + ISNULL(MAX(W.SaleReturnAmount), 0)
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
                                                              --AND tGood.GoodType = @Type
                                                INNER JOIN tUnitGood ON tGood.Unit = tUnitGood.Code
                                      --WHERE     GoodType = @Type
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

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[InsertFactorDetail]  (
	 @DetailsString NVARCHAR(4000) ,
	 @intSerialNo bigint ,
	 @intserialNo2 bigint ,
	 @Customer Bigint ,
	 @Branch int = Null
	
) 
As


if @Branch is null
    select @Branch = branch from tInventory where inventoryNo=(SELECT Top 1  intInventoryNo FROM Split(@DetailsString))

Declare @Status Int 

Set @Status = (Select Status from tfacm Where intserialno = @intSerialNo and Branch = @Branch)


     INSERT INTO tFacD
	(
	    
		intRow,
		Amount ,
		GoodCode  ,
		FeeUnit ,
		Discount ,
		Rate ,
		ChairName ,
		[ExpireDate] ,
		intInventoryNo ,
		DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
		ServePlace ,
		DifferencesCodes , 
		DifferencesDescription ,
		intSerialNo , 
		Branch 
	)
	     SELECT
		
		tmpTable.Row ,
		tmpTable.Amount ,
		tmpTable.GoodCode ,
		tmpTable.FeeUnit ,
		tmpTable.Discount ,
		tmpTable.Rate ,
		tmpTable.ChairName ,
		tmpTable.[ExpireDate],
		tmpTable.intInventoryNo ,
		tmpTable.DestInventoryNo ,--Because Has a Relation and Can not insert for  Another Branch
		tmpTable.ServePlace ,
		tmpTable.DifferencesCode ,
		tmpTable.DifferencesDescription ,
		@intSerialNo , 
		@Branch 	
	
	FROM (SELECT * FROM Split(@DetailsString )) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode

	DECLARE @InventoryNo INT 
	select @InventoryNo=  (SELECT TOP 1  intInventoryNo FROM Split(@DetailsString))      
	DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

	If @Status = 6 AND @intSerialNo2 > 0 AND @DestinventoryNo > 0
	Begin
	
	declare @destbranch INT
	select @destbranch=@Branch --branch from tInventory where inventoryNo=(SELECT Top 1  DestInventoryNo FROM Split(@DetailsString))
	  	   begin
			 INSERT INTO tFacD
			(
			    
				intRow,
				Amount ,
				GoodCode  ,
				FeeUnit ,
				Discount ,
				Rate ,
				ChairName ,
				[ExpireDate] ,
				intInventoryNo ,
				DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
				ServePlace ,
				DifferencesCodes , 
				DifferencesDescription ,
				intSerialNo , 
				Branch
			)
				 SELECT
				
				tmpTable.Row ,
				tmpTable.Amount ,
				tmpTable.GoodCode ,
				tmpTable.FeeUnit ,
				tmpTable.Discount ,
				tmpTable.Rate ,
				tmpTable.ChairName ,
				tmpTable.[ExpireDate],
				tmpTable.DestInventoryNo ,
				tmpTable.intInventoryNo ,--Because Has a Relation and Can not insert for  Another Branch
				tmpTable.ServePlace ,
				tmpTable.DifferencesCode ,
				tmpTable.DifferencesDescription ,
				@intSerialNo2 , 
				@DestBranch --dbo.Get_Current_Branch()
		
		
			FROM (SELECT * FROM Split(@DetailsString )) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode
	
		   end
	end
	

Update tFacD Set Amount = 1 where amount = 0 and intserialno = @intSerialNo and Branch = @Branch
--Update tFacD Set DestInventoryNo = Null Where intserialno = @intSerialNo and Branch = dbo.Get_Current_Branch()
	If (@Status = 2 OR @Status = 5) AND @intSerialNo2 > 0 
	Begin

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
                          WHERE     tfacm.Branch = @Branch AND tfacM.intSerialNo = @intSerialNo
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
                          WHERE     dbo.tFacM.Branch = @Branch AND tfacM.intSerialNo = @intSerialNo
                                    AND dbo.tFacD.GoodCode NOT IN (
                                    SELECT  dbo.tUsePercent.GoodCode
                                    FROM    dbo.tUsePercent
                                    WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                                    AND dbo.tGood.GoodType = 3
                        ) T
                GROUP BY T.GoodCode,
                        T.BuyPrice

  	   INSERT INTO tFacD
			(
			    
				intRow,
				Amount ,
				GoodCode  ,
				FeeUnit ,
				Discount ,
				Rate ,
				ChairName ,
				[ExpireDate] ,
				intInventoryNo ,
				DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
				ServePlace ,
				DifferencesCodes , 
				DifferencesDescription ,
				intSerialNo , 
				Branch
			)
	 SELECT
				
				tmpTable.Row ,
				tmpTable.Amount ,
				tmpTable.GoodCode ,
				tmpTable.BuyPrice ,
				0 , --tmpTable.Discount ,
				1 , --tmpTable.Rate ,
				NULL , --tmpTable.ChairName ,
				'' , --tmpTable.[ExpireDate],
				@InventoryNo   ,
				NULL , --tmpTable.DestInventoryNo ,
				1 , --tmpTable.ServePlace ,
				'', --tmpTable.DifferencesCode ,
				'', --tmpTable.DifferencesDescription ,
				@intSerialNo2 , 
				@Branch --dbo.Get_Current_Branch()
		
		
			FROM @ReturnTable tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode
	
	end

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

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
            @DetailsString NVARCHAR(4000),      
            @ds nText = '',      
            @Balance BIT ,      
            @AccountYear smallint = null  ,       
            @NvcDescription Nvarchar(150) = Null ,      
            @HavaleNo int = Null  ,      
            @TempAddress Nvarchar(255) = '',  
			@GuestNo INT,    
			@DetailsString2 NVARCHAR(4000) = NULL ,
			@DetailsString3 NVARCHAR(4000) = NULL ,
			@DetailsString4 NVARCHAR(4000) = NULL ,
			@AddedTotal FLOAT = NULL  ,
			@Rasmi BIT = NULL  ,
            @lastFacMNo INT OUT  ,
		    @Person INT = NULL     
             )      

AS      
IF @AddedTotal IS NULL SET @AddedTotal = 0
IF @Rasmi IS NULL SET @Rasmi = 0

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''

DECLARE @D1 NVARCHAR(4000) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(4000))  +  CAST(@DetailsString2 AS NVARCHAR(4000))  + CAST(@DetailsString3 AS NVARCHAR(4000))  + CAST(@DetailsString4 AS NVARCHAR(4000)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  


Declare @intserialNo int      
Declare @intserialNo2 int      
--Declare @intserialNo3 Bigint    

SET @intserialNo = 0        
SET @intserialNo2   = 0      
--SET @intserialNo3   = 0      

DECLARE @No1  INT     
DECLARE @No2  INT     
--DECLARE @No3  INT     

DECLARE @SumPrice  FLOAT       
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
         FROM Split(@D1)      
           ) tmpTable      

----------------------------------------Date From Server-----------------------------------------------------------------      
If @Status = 2 And dbo.Get_DateFromServer() = 1      
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      
ELSE
	IF LEN(@Date) < 8
		SET @Date = dbo.fnFixDateString(@Date) ------For Check Date String in Valid Format YY/MM/DD-----


------Start New Line For Avoid Repeat in tFacm------
DECLARE @RepeatNo INT

SELECT @RepeatNo = COUNT(*) FROM dbo.tFacM WHERE Status = 2 AND [Date] = @Date 
    AND TableNo = @TableNo AND Recursive = 0 AND FacPayment = 0 --AND Balance = 0

IF @RepeatNo > 0 
    GOTO EventHandler

----End New Line -----------------------------------------------------------------------------------------------      

 Declare @intBranch  int      
 Declare @ShiftNo int      
 DECLARE @TempNo INT 

 SELECT @intBranch = dbo.Get_Current_Branch()
 
 --select @intBranch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@DetailsString))      
 --IF @intBranch = 0 OR @intBranch IS NULL     SET @intBranch = dbo.Get_Current_Branch()

    DECLARE @IdentityNo INT
    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
    FROM    tFacm
    WHERE   Branch = @intBranch 

    IF @IdentityNo < ( @intBranch * 10000000 ) 
        SET @IdentityNo = ( @intBranch * 10000000 )

 SET @NO1 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND AccountYear = @AccountYear)      

 SET @ShiftNo= dbo.Get_Shift(GETDATE())      
 SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      

IF COL_LENGTH('[tFacM]','ServePlaceTempNo') IS NULL
	ALTER TABLE dbo.tFacM  ADD ServePlaceTempNo INT NULL 

DECLARE @ServePlaceTempNo INT 
 SET @ServePlaceTempNo = (SELECT ISNULL(MAX(ServePlaceTempNo),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ServePlace = @MasterServePlace)      


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
		TempNo ,
		ServePlaceTempNo  ,
		Rasmi  
		
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
		@TempNo ,
		@ServePlaceTempNo  ,
		@Rasmi
 )      
     IF @@ERROR <>0      
        GoTo EventHandler       

    SET @intserialNo = @IdentityNo

declare @destbranch  INT 
SET @destbranch = 0
DECLARE @TempNo2 INT 
DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@D1))      

	select @destbranch=  @intBranch --   branch from tInventory where inventoryNo=(SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

    DECLARE @DestStatus INT 	
    IF @Status = 2
        AND dbo.AutoHavale() = 1 
 			SELECT  @DestStatus = 6 ,
                 @NO2 =   ISNULL(MAX([NO]), 0) + 1
                    FROM    tFacM
                    WHERE   Status = 6
                            AND Branch = @intBranch
                            AND AccountYear = @AccountYear
                        
    IF @Status = 5
        AND dbo.AutoHavale() = 1 
 			SELECT  @DestStatus = 7 ,
                 @NO2 =   ISNULL(MAX([NO]), 0) + 1
                    FROM    tFacM
                    WHERE   Status = 7
                            AND Branch = @intBranch
                            AND AccountYear = @AccountYear
                        
   IF ( @Status = 6
             AND [dbo].[AutoResid]() = 1 AND @DestinventoryNo > 0
           ) 
	        SELECT  --@Customer = NULL , 
	        	@DestStatus = 7 ,
	                @NO2 = ISNULL(MAX([NO]), 0) + 1
	                    FROM    tFacM
	                    WHERE   Status = 7
	                            AND Branch = @destbranch
	                            AND AccountYear = @AccountYear
	
    IF ( @Status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1
             AND @DestinventoryNo > 0
           ) 
           OR 
           (@Status = 5
			AND dbo.AutoHavale() = 1 )

  BEGIN
 
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
                @DestStatus,      
                @Customer ,      
                @SumPrice ,      
                @OrderType ,      
                1 , --@MasterServePlace ,      
                @StationId ,      
                0 , --@ServiceTotal ,      
                @Recursive ,      
                0 , --@CarryFeeTotal ,      
                0 , --@PackingTotal ,      
                0 , --@DiscountTotal ,      
                0 ,
                0 ,      
                @newtime,      
                @User ,      
                @TableNo,      
                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
                NULL , --@Incharge ,      
                @owner ,      
                @FacPayment ,      
                @Balance ,      
				@DestBranch ,     
				@AccountYear ,      
				@NvcDescription,      
				@TempAddress,
				0 , --@GuestNo ,
				NULL --@TempNo2    
		
 )      
		 IF @@ERROR <>0      
			GoTo EventHandler      
		SET @intserialNo2 = @IdentityNo + 1      

             IF @status = 2
                BEGIN 
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' ÕÊ«·Â - '
                        + CAST(@No2 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo AND Branch = @intBranch
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' ›«ò Ê— ›—Ê‘  - '
                        + CAST(@No1 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo2 AND Branch = @intBranch
                UPDATE  tfacm
                SET     RefrenceHavale = @intserialNo2
                WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch
				
				END 
             IF @status = 5
                BEGIN 
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' —”Ìœ - '
                        + CAST(@No2 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo AND Branch = @intBranch
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' »—ê‘  «— ›—Ê‘  - '
                        + CAST(@No1 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo2 AND Branch = @intBranch
                UPDATE  tfacm
                SET     RefrenceHavale = @intserialNo2
                WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch
				
				END 
            IF @status = 6
             BEGIN 
               UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' —”Ìœ - '
                        + CAST(@No2 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' ÕÊ«·Â  - '
                        + CAST(@No1 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo2 AND Branch = @intBranch
                UPDATE  tfacm
                SET     RefrenceHavale = @intserialNo2
                WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch
			END 

end      


----------------------------------Fill Details Factor  --------------------------------------------------------------      
 exec InsertFactorDetail @D1 , @intserialNo , @intserialNo2, @Customer , @intBranch      

     IF @@ERROR <>0      
        GoTo EventHandler      
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------      

----------------------------------Total SumPrice Calculate  --------------------------------------------------------------      
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(SUM(ROUND( (Amount * FeeUnit * discount/100 ) ,0) ) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(SUM( ROUND((Amount * FeeUnit) * (1 - discount/100),0) )  AS BIGINT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

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

Declare @SumPrice2 FLOAT       
Set @SumPrice2 = (Select Cast(Sum(Amount * FeeUnit) as FLOAT ) From tFacd Where intSerialNo = @intserialNo2 And Branch = @DestBranch )        
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

-- ===================For Calculate Service In Delivery Or Out ==================
	IF @MasterServePlace = 2 OR @MasterServePlace = 4 
		SET @ServiceTotal = 0

-------------------------------------

 If @ServiceTotal <> 0      
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)      
     IF @@ERROR <>0      
        GoTo EventHandler       
----------------------------------Round Sumprice  --------------------------------------------------------------      
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5  OR @status = 10
 BEGIN 
	IF @Rasmi = 1
	  begin
	  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
	  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
	  END 
	ELSE
	    BEGIN 
		SET @DutyTotal = 0
		SET @TaxTotal = @AddedTotal
		END 
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
IF @Status = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 7 , @User , 1 , @AccountYear , @destbranch      
     IF @@ERROR <>0      
   GoTo EventHandler      
IF @Status = 2 AND dbo.AutoHavale() = 1 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 6 , @User , 1 , @AccountYear , @destbranch      
     IF @@ERROR <>0      
   GoTo EventHandler      


----------------------------Cash ---------------------------      

------------------------Mojodi Control Online--------------------------------------------      

Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @intBranch      
IF @@ERROR <>0      
 GoTo EventHandler      
 IF ( @Status = 2
     AND dbo.AutoHavale() = 1
   )
    OR ( @Status = 6
         AND [dbo].[AutoResid]() = 1  AND @DestinventoryNo > 0
       ) 
 BEGIN      
	 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch      
	 IF @@ERROR <>0      
	 GoTo EventHandler      

	    EXEC InsertMojodiCalculate @DestStatus, @intserialNo2, @AccountYear, @intBranch
	 IF @@ERROR <>0      
	 GoTo EventHandler      
 END       
IF dbo.AutoHavale() = 1
        UPDATE  tfacm
        SET     [BitHavaleResid] = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch

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


-- €ÌÌ—  «—ÌŒ ›«ò Ê— Œ—Ìœ Ê ⁄œ„  €ÌÌ—  «—ÌŒ ›«ò Ê— ›—Ê‘
ALTER  PROCEDURE [dbo].[EditFactorMasterDetails]  (  


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
	@DetailsString  NVARCHAR(4000),  
	@ds nText = '',  
	@Balance Bit,  
	@AccountYear Smallint = Null ,  
	@NvcDescription Nvarchar(150) = Null ,  
	@TempAddress Nvarchar(255) = '', 
	@GuestNo INT,     
	@DetailsString2 NVARCHAR(4000) = NULL ,
	@DetailsString3 NVARCHAR(4000) = NULL ,
	@DetailsString4 NVARCHAR(4000) = NULL ,
	@AddedTotal FLOAT = NULL ,
	@Rasmi BIT = NULL ,
	@LastFacMNo  INT OUT  ,
	@Person INT = NULL 
  )  


AS 
IF @AddedTotal IS NULL SET @AddedTotal = 0
IF @Rasmi IS NULL SET @Rasmi = 0

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''

DECLARE @D1 NVARCHAR(4000) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(4000))  +  CAST(@DetailsString2 AS NVARCHAR(4000))  + CAST(@DetailsString3 AS NVARCHAR(4000))  + CAST(@DetailsString4 AS NVARCHAR(4000)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  
 
DECLARE @SumPrice FLOAT  
DECLARE @SumPrice2 FLOAT  
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

 SELECT @intBranch = dbo.Get_Current_Branch()

-- select @intBranch = branch from tInventory where inventoryNo=(SELECT TOP 1  IntInventoryNo FROM Split(@DetailsString))  
 SET @ShiftNo= dbo.Get_Shift(GETDATE())  

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  

--Control is difficult
--If No received then Bypass received
--DECLARE @DestinventoryNo INT 
--select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      


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
 (  SELECT DISTINCT ServePlace As SServePlace  FROM Split(@D1)) tmpTable  


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

DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@D1))      

    DECLARE @DestStatus INT 
    DECLARE @BitHavaleResid INT 
    IF ( @status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1  --AND  @DestinventoryNo > 0 
           ) 
         OR (@status = 5
         AND dbo.AutoHavale() = 1 )
        SELECT  @intSerialNo2 = ISNULL(RefrenceHavale , 0) ,
		        @BitHavaleResid = ISNULL(BitHavaleResid , 0) 
                               FROM      dbo.tFacM
                              WHERE     intSerialNo = @intSerialNo
                                        AND Branch = @intBranch

        SELECT  @DestStatus = Status 
                               FROM      dbo.tFacM
                              WHERE     intSerialNo = @intSerialNo2
                                        AND Branch = @intBranch


 select @destbranch= @intBranch -- branch from tInventory where inventoryNo=(SELECT TOP 1 DestInventoryNo FROM Split(@DetailsString))  

---------------------------------------Mojodi Control Online---------------------------------------------------------  
Exec DeleteMojodiCalculate @Status , @intserialNo  ,  1 , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
    IF ( @status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1 --AND @DestinventoryNo > 0  Because AutoHavale is without destination
           ) 
           OR (@status = 5
         AND dbo.AutoHavale() = 1 )
        EXEC DeleteMojodiCalculate @DestStatus, @intserialNo2, 1, @AccountYear, @intBranch

    IF @@ERROR <> 0 
        GOTO EventHandler
 ----------------------------------------Delete Old Details -----------------------------------------------------------  
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
If  @intSerialNo2 > 0--And (@destbranch = @intBranch or dbo.AutoResid() = 1 )   
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo2 AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
------------------------------------------------------------    
  Exec DeleteFactorChildren @intSerialNo , @intBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
 If @intSerialNo2 > 0--And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
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
		TempNo = CASE WHEN @Status = 2 THEN  TempNo WHEN @Status = 5 THEN TempNo ELSE NULL END ,
		Rasmi = @Rasmi     
    WHERE tFacM.intSerialNo = @intSerialNo  AND Branch =  @intBranch  

    IF @@ERROR <>0  
        GoTo EventHandler  

    IF ( @status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1 AND  @DestinventoryNo > 0 
           ) 
           OR ( @status = 5
         AND dbo.AutoHavale() = 1
       )

    BEGIN

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
    WHERE tFacM.intSerialNo = @intSerialNo2  AND Branch =  @intBranch  

END  

----------------------------------Fill Details Factor ----------------------------------------------------------------------  
 exec InsertFactorDetail @D1 , @intserialNo , @intserialNo2, @Customer , @intBranch  

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

If @intSerialNo2 > 0  --And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1 )  
   Set @SumPrice2 = (Select Cast (Sum(Amount * FeeUnit) as FLOAT )   From tFacd Where intSerialNo = @intSerialNo2 And Branch = @intBranch )    
   IF @@ERROR <>0  
        GoTo EventHandler  
PRINT @SumPrice2
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

-- ===================For Calculate Service In Delivery Or Out ==================
	IF @MasterServePlace = 2 OR @MasterServePlace = 4 
		SET @ServiceTotal = 0

-------------------------------------

 If @ServiceTotal <> 0  
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)  

     IF @@ERROR <>0  
        GoTo EventHandler   
----------------------------------Round Sumprice  --------------------------------------------------------------  
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5 OR @status = 10
 BEGIN 
  IF @Rasmi  = 1
  BEGIN 
	  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
	  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  END
  ELSE
  BEGIN
	SET @DutyTotal = 0
	SET @TaxTotal = @AddedTotal
  END 	
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
If @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1)   

    Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @intBranch  
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
	IF ( @status = 2  AND dbo.AutoHavale() = 1)
		OR ( @Status = 6   AND [dbo].[AutoResid]() = 1 AND  @DestinventoryNo > 0    ) 
		OR ( @status = 5    AND dbo.AutoHavale() = 1 )

	 BEGIN  
	 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch  
	 IF @@ERROR <>0  
	 GoTo EventHandler  

	EXEC InsertMojodiCalculate @DestStatus, @intserialNo2, @AccountYear, @intBranch
    IF @@ERROR <> 0 
        GOTO EventHandler
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
        BEGIN
	--IF dbo.AutoHavale() = 0
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
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
            SET     Mojodi = Mojodi - X.Amount ,
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

            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + X.Amount ,
                    SaleReturnAmount = SaleReturnAmount + X.Amount
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


ALTER PROCEDURE [dbo].[DeleteMojodiCalculate]
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
                BEGIN
		--IF dbo.AutoHavale() = 0
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
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
                    SET     Mojodi = Mojodi + Amount ,
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

                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - Amount ,
                            SaleReturnAmount = SaleReturnAmount - Amount
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
        END

    ELSE 
        IF @Recursive = 0 
            BEGIN
                IF @Status = 2 
                    BEGIN
	   	    --IF dbo.AutoHavale() = 0
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - t.Amount ,
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
                        SET     Mojodi = Mojodi - Amount ,
                                SaleAmount = SaleAmount + Amount
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

						UPDATE  tInventory_Good
						SET     Mojodi = Mojodi + Amount ,
								SaleReturnAmount = SaleReturnAmount + Amount
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
            END
--===============================================

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  VIEW dbo.vw_FacM_Per
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
		, ISNULL(tfacM.Rasmi , 0) AS Rasmi
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




