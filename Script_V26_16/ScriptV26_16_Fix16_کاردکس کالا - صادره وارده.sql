

--ScriptV26_16_Fix16_˜ÇÑÏ˜Ó ˜ÇáÇ - ÕÇÏÑå æÇÑÏå.sql
--95/05/16


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   FUNCTION [dbo].[FnGetSellBuyKindInfo]
    (
      @DateBefore NVARCHAR(50),
      @DateAfter NVARCHAR(50),
      @InventoryNo INT,
      @Branch INT,
      @AccountYear SMALLINT   
    )
RETURNS @ReturnTable TABLE
    (
      intInventoryNo INT,
      --Fromstore NVARCHAR(50),
      --DestDescription NVARCHAR(50),
      GoodCode INT,
      --Status INT,
      --NvcDescription NVARCHAR(50),
      Branch INT,
      [Name] NVARCHAR(50),
      Amountv FLOAT ,
      Amounts FLOAT ,
      feev BIGINT,
      fees BIGINT,
      Mojodi FLOAT,
      FirstMojodi FLOAT,
      FirstPrice INT,
      FirstMojodiPrice BIGINT,
      [Date] NVARCHAR(50),
      [No] INT,
      --NamePrn NVARCHAR(50),
      BarCode NVARCHAR(50),
      BuyPrice INT
    )
AS BEGIN   

    INSERT  INTO @ReturnTable
            (
              intInventoryNo,
              --Fromstore,
              --DestDescription,
              GoodCode,
              --Status,
              --NvcDescription,
              Branch,
              Name,
              Amountv,
              Amounts,
              feev,
              fees,
              Mojodi,
              FirstMojodi,
              FirstPrice,
              FirstMojodiPrice   
	--,[Date] ,[No] 
              ,
              --NamePrn,
              BarCode,
              BuyPrice                
                
            )
            SELECT  intInventoryNo,
                    --Fromstore,
                    --DestDescription,
                    GoodCode,
                    --Status,
                    --NvcDescription,
                    Branch,
                    [Name],
                    CAST(SUM(Amountv) AS FLOAT),
                    CAST(SUM(Amounts) AS FLOAT),
                    SUM(feev),
                    SUM(fees),
                    Mojodi,
                    FirstMojodi,
                    FirstPrice,
                    FirstMojodiPrice,
                    --NamePrn,
                    BarCode,
                    BuyPrice
            FROM    ( SELECT    dbo.tFacD.intInventoryNo,
--                                 ( SELECT    Description
--                                   FROM      tInventory
--                                   WHERE     inVentoryNo = tfacD.intInventoryNo
--                                 ) AS Fromstore,
--                                 ISNULL(dbo.tInventory.[Description], '') AS DestDescription,
                                dbo.tFacD.GoodCode,
                                --Status,
                                --tStatusType.NvcDescription,
                                tFacD.Branch,
                                tGood.Name,
                                CASE WHEN tfacm.Status IN ( 1, 7 )
                                     THEN CAST(SUM(tFacd.Amount * tStatusType.Flag) AS FLOAT)
                                     ELSE 0
                                END AS Amountv,
                                CASE WHEN tfacm.Status IN ( 4, 6, 3 )
                                     THEN CAST(SUM(tFacd.Amount * tStatusType.Flag) AS FLOAT)
                                     ELSE 0
                                END AS Amounts,
                                CASE WHEN tfacm.Status IN ( 1, 7 )
                                     THEN SUM(tFacd.Amount * tStatusType.Flag)
                                          * tfacd.FeeUnit
                                     ELSE 0
                                END AS feev,
                                CASE WHEN tfacm.Status IN ( 4, 6, 3 )
                                     THEN SUM(tFacd.Amount * tStatusType.Flag)
                                          * tfacd.FeeUnit
                                     ELSE 0
                                END AS fees,
                                tInventory_Good.Mojodi,
                                tInventory_Good.FirstMojodi,
                                tInventory_Good.FirstPrice,
                                tInventory_Good.FirstMojodi
                                * tInventory_Good.FirstPrice AS FirstMojodiPrice ,  
				--, tFacm.Date ,Min(tFacm.No) as No ,  
                                --tgood.NamePrn,
                                tgood.BarCode,
                                tgood.buyPrice
                      FROM    dbo.tInventory_Good  
                                INNER JOIN tInventory ON tInventory.InventoryNo = tInventory_Good.InventoryNo 
							AND tInventory_Good.InventoryNo = @InventoryNo
                                INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
                                INNER JOIN dbo.tFacD ON tInventory_Good.GoodCode = tFacd.GoodCode
                                                   AND tInventory_Good.InventoryNo = tFacD.intInventoryNo
                                                   AND tInventory_Good.InventoryNo = @InventoryNo
						   AND tInventory_Good.AccountYear = @AccountYear
                                                   AND tInventory_Good.Branch = @Branch
				INNER JOIN dbo.tFacM ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                                        AND dbo.tFacM.Branch = dbo.tFacD.Branch
							AND dbo.tFacM.Recursive = 0
			                                AND dbo.tFacM.AccountYear = @AccountYear
			                                AND tFacM.[Date] >= @DateBefore
			                                AND tFacM.[Date] <= @DateAfter
			                                AND tFacm.Branch = @Branch
			                                AND tfacm.Status IN ( 1, 3, 4, 6, 7 )
                                INNER JOIN dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo

                      GROUP BY  Mojodi,
                                FirstMojodi,
                                FirstPrice ,--tfacm.[Date],tfacm.TIME,  
                                tGood.Name,
                                tfacD.Branch,
                                tfacD.intInventoryNo,
                                --tgood.NamePrn,
                                tFacd.Goodcode,
                                tfacM.status,
                                --tStatusType.NvcDescription,
                                dbo.tInventory.[Description],
                                dbo.tFacD.FeeUnit,
                                tgood.BarCode,
                                tgood.BuyPrice 
	   
	--Order by tfacm.Date,tfacm.Time  
                      
                    ) T
            GROUP BY intInventoryNo,
                    --Fromstore,
                    --DestDescription,
                    GoodCode,
                    --Status,
                    --NvcDescription,
                    Branch,
                    [Name],
                    Mojodi,
                    FirstMojodi,
                    FirstPrice,
                    FirstMojodiPrice,
                    --NamePrn,
                    BarCode,
                    BuyPrice
    RETURN   
   END

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER    PROCEDURE [dbo].[GetInventoryGood_AllMojodi_Report]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @AccountYear1 SMALLINT,
      @Inventory1 INT,
      @Branch1 INT
    )
AS 
    BEGIN
        --DECLARE @Fromstore NVARCHAR(50)  
        --DECLARE @DestDescription NVARCHAR(50)  
        DECLARE @GoodCode INT  
        --DECLARE @Status INT  
        DECLARE @NvcDescription NVARCHAR(50)  
        DECLARE @Name NVARCHAR(50)  
        DECLARE @Amountv FLOAT    
        DECLARE @Amounts FLOAT   
        DECLARE @feev BIGINT  
        DECLARE @fees BIGINT   
        DECLARE @Mojodi FLOAT  
        DECLARE @FirstMojodi FLOAT  
        DECLARE @FirstPrice INT  
        DECLARE @FirstMojodiPrice INT  
        --DECLARE @NamePrn NVARCHAR(50)  
        DECLARE @BarCode NVARCHAR(50) 
        DECLARE @BuyPrice INT  
  
        DECLARE @tblFirstDateMojodi TABLE
            (
              GoodCode INT,
              FirstDateMojodi FLOAT 
            )  
  
        DECLARE @tblReturnDateMojodi TABLE
            (
              intInventoryNo INT,
              --Fromstore NVARCHAR(50),
              --DestDescription NVARCHAR(50),
              GoodCode INT,
              --Status INT,
              NvcDescription NVARCHAR(50),
              Branch INT,
              [Name] NVARCHAR(50),
              Amountv FLOAT ,
              Amounts FLOAT ,
              feev BIGINT,
              fees BIGINT,
              Mojodi FLOAT,
              FirstMojodi FLOAT,
              FirstPrice INT,
              FirstMojodiPrice INT,
              --NamePrn NVARCHAR(50),
              BarCode NVARCHAR(50),
              BuyPrice INT
            )  
  
        DECLARE @CurrentMojodi INT  
  
        INSERT  INTO @tblFirstDateMojodi
                (
                  GoodCode,
                  FirstDateMojodi
                )
                SELECT  GoodCode,
                        FirstDateMojodi
                FROM    dbo.FnFirstDateMojodi(@Date1, @Inventory1, @Branch1,
                                              @AccountYear1 , 0 , 0 , 0 , 0)  
  
        DECLARE total_cursor CURSOR
            FOR SELECT  intInventoryNo,
--                         Fromstore,
--                         DestDescription,
                        GoodCode,
--                         Status,
                        N'' AS  NvcDescription,
                        Branch,
                        Name,
                        CASE WHEN Amountv >= 0 THEN Amountv ELSE -1*Amountv END AS Amountv,
                        CASE WHEN Amounts >= 0 THEN Amounts ELSE -1*Amounts END AS Amounts,
                        feev,
                        fees,
                        Mojodi,
                        FirstMojodi,
                        FirstPrice,
                        FirstMojodiPrice,
--                         NamePrn,
                        BarCode,
                        BuyPrice
                FROM    dbo.FnGetSellBuyKindInfo(@Date1, @Date2, @Inventory1,
                                                 @Branch1, @AccountYear1)  
  
 
          OPEN total_cursor  
  
        FETCH NEXT FROM total_cursor INTO @Inventory1, --@Fromstore,@DestDescription,
            @GoodCode, @NvcDescription, @Branch1,
            @Name, @Amountv, @Amounts, @feev, @fees, @Mojodi, @FirstMojodi,
            @FirstPrice, @FirstMojodiPrice, @BarCode, @BuyPrice    
  
        WHILE @@FETCH_STATUS = 0 
            BEGIN  
                SELECT  @CurrentMojodi = FirstDateMojodi
                FROM    @tblFirstDateMojodi
                WHERE   GoodCode = @GoodCode  
                SET @CurrentMojodi = ISNULL(@CurrentMojodi, 0)  
 -------  
                IF @Amountv > 0 
                    SET @CurrentMojodi = @CurrentMojodi + @Amountv  

                IF @Amounts > 0 
                    SET @CurrentMojodi = @CurrentMojodi - @Amounts  

 -------  
                INSERT  INTO @tblReturnDateMojodi
                        (
                          intInventoryNo,
                          --Fromstore,
                          --DestDescription,
                          GoodCode,
                          --Status,
                          NvcDescription,
                          Branch,
                          Name,
                          Amountv,
                          Amounts,
                          feev,
                          fees,
                          Mojodi,
                          FirstMojodi,
                          FirstPrice,
                          FirstMojodiPrice,
                          --NamePrn,
                          BarCode,
                          BuyPrice 
                        )
                VALUES  (
                          @Inventory1,
                          --@Fromstore,
                          --@DestDescription,
                          @GoodCode,
                          --@Status,
                          @NvcDescription,
                          @Branch1,
                          @Name,
                          @Amountv,
                          @Amounts,
                          @feev,
                          @fees,
                          @CurrentMojodi,
                          @FirstMojodi,
                          @FirstPrice,
                          @FirstMojodiPrice,
                          --@NamePrn,
                          @BarCode,
                          @BuyPrice 
                        )  
   
                UPDATE  @tblFirstDateMojodi
                SET     FirstDateMojodi = @CurrentMojodi
                WHERE   GoodCode = @GoodCode  
   
  
                FETCH NEXT FROM total_cursor INTO @Inventory1, --@Fromstore,@DestDescription,
                    @GoodCode, @NvcDescription,
                    @Branch1, @Name, @Amountv, @Amounts, @feev, @fees, @Mojodi,
                    @FirstMojodi, @FirstPrice, @FirstMojodiPrice, 
                    @BarCode, @BuyPrice    
            END  
        CLOSE total_cursor  
        DEALLOCATE total_cursor  
  
        SELECT  *,
                @SystemDay + ' ' + @SystemDate + ' ' + N' ÓÇÚÊ : '
                + @SystemTime AS Sysdate
        FROM    @tblReturnDateMojodi
    END

GO
