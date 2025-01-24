
--ScriptV26_16_Fix17_UpdatefacD_FinalPrice_950805.sql
--95/08/05

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
	SELECT @GoodAmount = Amount FROM dbo.tFacD WHERE intSerialNo  = @intSerialNo AND GoodCode = @GoodCode
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
