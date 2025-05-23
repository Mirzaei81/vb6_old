

--Script_V26_16_Fix7_Inventory
--فقط در ورژن الماس
-- گزارشات مصرف در یک بازه زمانی 
--گزارش از ورود و خروج کالاها با مانده اولیه
--گزارش موجودی نعدادی انبار در بازه زمانی با مانده اولیه
--گزارش موجودی ریالی انبار در بازه زمانی با مانده اولیه
--با مانده اولیه
-- 93/07/23



INSERT INTO dbo.tbltotal_ItemReports
        ( intReportId ,
          intGroupReportId ,
          ReportName ,
          LatinReportName ,
          Refrence_Sp
        )
VALUES  ( 94 , -- intReportId - int
          2 , -- intGroupReportId - int
          N'گزارش از ورود و خروج کالاها' , -- ReportName - nvarchar(100)
          'RepMojodiControlGoods' , -- LatinReportName - varchar(50)
          'Get_InventoryMojodiControl'  -- Refrence_Sp - varchar(50)
        )
GO

UPDATE tbltotal_ItemReports_Details SET Quantity = 1 WHERE intReportId = 8 AND Row = 4
GO


--گزارش موجودی ریالی انبار در بازه زمانی

INSERT INTO dbo.tbltotal_ItemReports
        ( intReportId ,
          intGroupReportId ,
          ReportName ,
          LatinReportName ,
          Refrence_Sp
        )
VALUES  ( 95 , -- intReportId - int
          2 , -- intGroupReportId - int
          N'گزارش موجودی ریالی انبار در بازه زمانی' , -- ReportName - nvarchar(100)
          'RepInventoryAtomicRials_Report_InDate' , -- LatinReportName - varchar(50)
          'Get_InventoryAtomicRials_Report_InDate'  -- Refrence_Sp - varchar(50)
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
VALUES  ( 448 , -- intObjectCode - int
          N'RepMojodiControlGoods' , -- ObjectId - nvarchar(50)
          N'گزارش از ورود و خروج کالاها' , -- ObjectName - nvarchar(50)
          N'RepMojodiControlGoods' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          108  -- ObjectParent - int
        )
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          448  -- intObjectCode - int
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
VALUES  ( 449 , -- intObjectCode - int
          N'RepInventoryAtomicRials_Report_InDate' , -- ObjectId - nvarchar(50)
          N'گزارش موجودی ریالی انبار در بازه زمانی' , -- ObjectName - nvarchar(50)
          N'RepInventoryAtomicRials_Report_InDate' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          108  -- ObjectParent - int
        )
GO


INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          449  -- intObjectCode - int
          )
GO


INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  94 ,
          1 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 81 AND Row = 1 

GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  94 ,
          2 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          1 ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 81 AND Row = 2 

GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  94 ,
          3 ,
          N'انبارفروش' ,
          ' ' ,
          'intInventory' ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          1 ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 81 AND Row = 2 

GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  94 ,
          4 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 84 AND Row = 5 

GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  94 ,
          5 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 85 AND Row = 2 

GO

--DELETE FROM tblTotal_ItemReports_Details WHERE intReportId = 95


INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
		  95 ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 8

GO


--SELECT * FROM dbo.tblTotal_ItemReports_Details WHERE intReportId = 94
--GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FnUsedGoods_FirstDate]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FnUsedGoods_FirstDate]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE  FUNCTION [dbo].[FnUsedGoods_FirstDate]
    (
      @DateAfter VARCHAR(10),
      @UsedInventory INT,
      @SaleInventory INT,
      --@Branch INT,
      @AccountYear SMALLINT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT ,
      @GoodCode INT   
    )
RETURNS @ReturnTable TABLE
    (
      FirstGoodCode INT,
      FirstDateMojodi FLOAT ,
      FirstDateUsed FLOAT
    )
AS 
BEGIN 

DECLARE @UsedBranch INT 
SET @UsedBranch = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @UsedInventory)

DECLARE @SaleBranch INT 
SET @SaleBranch = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @SaleInventory)

IF @GoodLevel11 = 0
 SELECT @GoodLevel11 = MIN(Code) FROM dbo.tGoodLevel1
IF @GoodLevel12 = 0
 SELECT @GoodLevel12 = Max(Code) FROM dbo.tGoodLevel1

    INSERT  INTO @ReturnTable
            (
              FirstGoodCode,
              FirstDateMojodi ,
              FirstDateUsed                
            )


SELECT X1.FirstGoodCode AS FirstGoodCode ,
			CAST(X2.FirstDateMojodi AS DECIMAL(20,3)) AS FirstDateMojodi  ,
			X1.FirstDateUsed AS FirstDateUsed 
	 FROM 
(
        SELECT    
			ISNULL(k.GoodCode,dbo.tInventory_Good.GoodCode) AS FirstGoodCode ,
			CAST(ISNULL(k.Amount, 0) AS DECIMAL(20,3)) AS FirstDateUsed 
			
			FROM      (

					SELECT  SUM(T5.Amount) AS Amount,
							T5.Branch,T5.AccountYear ,
							T5.GoodCode
							FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
										( dbo.tFacD.Amount * dbo.tUsePercent.fltUsedValue ) AS Amount,
										 dbo.tFacM.Branch , dbo.tFacM.AccountYear
										FROM      dbo.tFacM
										JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
														  AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
										JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
																AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
										JOIN dbo.tGood ON dbo.tUsePercent.GoodCode = dbo.tGood.Code
							  WHERE     tfacm.Branch = @SaleBranch
										AND dbo.tFacM.AccountYear = @AccountYear
										AND tfacm.Recursive = 0
										AND tfacm.status = 2
										AND tfacm.Date < @DateAfter
										AND tfacd.intInventoryNo = @SaleInventory
										AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
									AND (tFacd.GoodCode = @GoodCode OR @GoodCode = 0)
								)T5
						GROUP BY T5.GoodCode ,T5.Branch,T5.AccountYear
			) k
			
			INNER JOIN tinventory_good ON k.goodcode = tinventory_good.goodcode
			--AND (tinventory_good.Mojodi <> 0 OR @ZeroMojodi = 0)
			AND tinventory_good.inventoryno = @UsedInventory
			AND tinventory_good.AccountYear = k.AccountYear
			--AND tinventory_good.Branch = k.Branch
        

)X1

INNER JOIN 

(
         SELECT 
			ISNULL(dbo.tInventory_Good.FirstMojodi ,0)
			+ ISNULL(SUM(k.TotalBuy), 0)
			- ISNULL(SUM(k.TotalHavaleh), 0)
			+ ISNULL(SUM(k.TotalResid), 0)
			- ISNULL(SUM(k.TotalLoss), 0)
			- ISNULL(SUM(k.TotalBuyReturn), 0) AS FirstDateMojodi,
			k.GoodCode AS FirstGoodCode 

			FROM   
			   (
				SELECT   
				DISTINCT [D].[GoodCode] ,D.intinventoryNo,
				M.AccountYear,M.Branch ,
				--------------------------------------------------------------------------------------------مقدار موجودي در  حالتهاي مختلف ----------------------------------------
				CASE WHEN [M].Status = 1 THEN SUM( [D].[Amount])ELSE 0 END AS TotalBuy ,
				CASE WHEN [M].Status = 3 THEN SUM([D].[Amount]) ELSE 0 END AS TotalLoss ,
				CASE WHEN [M].Status = 4 THEN SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturn ,
				CASE WHEN [M].Status = 6 THEN SUM([D].[Amount]) ELSE 0 END AS TotalHavaleh ,
				CASE WHEN [M].Status = 7 THEN SUM([D].[Amount]) ELSE 0 END AS TotalResid 

				-----------------------------------------------------------------------------------------------------------------------------------------------------------------
				FROM      [dbo].[tFacM] M
						INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
								AND [M].[intSerialNo] = [D].[intSerialNo]
						INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode 
				  WHERE     M.Branch = @UsedBranch
							AND M.AccountYear = @AccountYear
							AND M.Recursive = 0
							AND M.Date < @DateAfter
							AND D.intInventoryNo = @SaleInventory -- @UsedInventory
				GROUP BY  [D].[GoodCode] ,
						[M].[Status] ,
						D.intInventoryNo ,
						M.AccountYear ,
						M.Branch
			) K
			INNER JOIN dbo.tInventory_Good 
				 ON k.goodcode = tInventory_Good.goodcode
				 AND K.AccountYear = tInventory_Good.AccountYear
				 --AND K.Branch = tInventory_Good.Branch
				 AND K.intInventoryNo = tInventory_Good.inventoryno
			GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    tInventory_Good.firstMojodi ,
                    K.GoodCode ,
                    K.Branch
)X2
ON X1.FirstGoodCode = X2.FirstGoodCode
	ORDER BY   X1.FirstGoodCode

    RETURN
   END
--==========================================
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_InventoryMojodiControl]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_InventoryMojodiControl]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

------------------------------------------------------------گزارش ورود و خروج کالا---------------------------------
CREATE     PROCEDURE [dbo].[Get_InventoryMojodiControl]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(10),
      @Date2 NVARCHAR(10),
      @Inventory1 INT ,
      @intInventory1 INT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT,
      @AccountYear1 SMALLINT   
     
    )
AS 

DECLARE @Branch1 INT 
SET @Branch1 = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @Inventory1)

DECLARE @Branch2 INT 
SET @Branch2 = (SELECT TOP 1 Branch FROM dbo.tInventory where InventoryNo = @intInventory1)

SELECT DISTINCT 
ISNULL(T1.GoodCode , T2.[GoodCode]) AS [GoodCode] ,
ISNULL(T1.FirstDateUsed ,0) AS FirstDateUsed ,
ISNULL(FirstMojodi ,0) AS FirstMojodi ,
ISNULL(FromStoreAmount ,0) AS FromStoreAmount,
ISNULL(toStoreAmount ,0) AS toStoreAmount ,
ISNULL(BuyAmount ,0) AS BuyAmount ,
ISNULL(T2.Amount ,0) AS SaleAmount2,
ISNULL(Mojodi ,0) AS Mojodi ,
ISNULL(T1.NAME , T2.NAME) AS Name
,(Select  Description FROM dbo.tInventory WHERE InventoryNo = @Inventory1) AS DESCRIPTION 
,(Select  Description FROM dbo.tInventory WHERE InventoryNo = @intInventory1) AS Destination 
,@Inventory1 AS InventoryNo,@AccountYear1 AS  AccountYear
,N'امروز'+' '+@SystemDay+'  '+N'مورخ '+@SystemDate+'  '+N'ساعت '+@SystemTime AS TotalDate
FROM 
(		

         SELECT 
            T4.Name ,
			ISNULL(T4.FirstDateUsed ,0) AS FirstDateUsed ,
			ISNULL(T4.FirstDateMojodi ,0) AS FirstMojodi ,
			ISNULL(T4.FirstDateMojodi ,0)
			+ ISNULL(SUM(k.TotalBuy), 0)
			- ISNULL(SUM(k.TotalHavaleh), 0)
			+ ISNULL(SUM(k.TotalResid), 0)
			- ISNULL(SUM(k.TotalLoss), 0)
			- ISNULL(SUM(k.TotalBuyReturn), 0) AS Mojodi,
			ISNULL(k.GoodCode,T4.GoodCode) AS GoodCode ,
			
			ISNULL(SUM(k.TotalBuy), 0) AS BuyAmount ,
			ISNULL(SUM(k.TotalLoss), 0) AS LossAmount ,
			ISNULL(SUM(k.TotalBuyReturn), 0) AS BuyReturnAmount,
			ISNULL(SUM(k.TotalHavaleh), 0) AS FromStoreAmount,
			ISNULL(SUM(k.TotalResid), 0) AS toStoreAmount


			FROM   
			   (
				SELECT   
				DISTINCT [D].[GoodCode] ,D.intinventoryNo,
				M.AccountYear,
				--------------------------------------------------------------------------------------------مقدار موجودي در  حالتهاي مختلف ----------------------------------------
				CASE WHEN [M].Status = 1 THEN SUM( [D].[Amount])ELSE 0 END AS TotalBuy ,
				CASE WHEN [M].Status = 3 THEN SUM([D].[Amount]) ELSE 0 END AS TotalLoss ,
				CASE WHEN [M].Status = 4 THEN SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturn ,
				CASE WHEN [M].Status = 6 THEN SUM([D].[Amount]) ELSE 0 END AS TotalHavaleh ,
				CASE WHEN [M].Status = 7 THEN SUM([D].[Amount]) ELSE 0 END AS TotalResid 

				-----------------------------------------------------------------------------------------------------------------------------------------------------------------
				FROM      [dbo].[tFacM] M
						INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
								AND [M].[intSerialNo] = [D].[intSerialNo]
						INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode 
								AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
				WHERE     [M].[Date] >= @Date1
						AND [M].[Date] <= @Date2
						AND [M].[AccountYear] = @AccountYear1
						AND [M].[Recursive] = 0
						AND [D].[intInventoryNo] = @intInventory1 -- @Inventory1
						AND [M].Branch = @Branch1
				GROUP BY  [D].[GoodCode] ,
						[M].[Status] ,
						D.intInventoryNo ,
						M.AccountYear
			) K

			RIGHT OUTER  JOIN 
			(SELECT  GoodCode, FirstDateMojodi , Name  , FirstDateUsed
			  FROM 
				(SELECT  FirstGoodCode AS GoodCode, FirstDateMojodi , FirstDateUsed 
                       FROM    dbo.FnUsedGoods_FirstDate(@Date1, @Inventory1, @intInventory1, --@Branch1
                                              @AccountYear1 , @GoodLevel11 , @GoodLevel12 , 0 )
                                              --@AccountYear1 , @GoodLevel11 , @GoodLevel12 , 0 )
                                              --WHERE  FirstDateMojodi <> 0
				)T 
				INNER JOIN tGood ON T.GoodCode = dbo.tGood.Code AND (tgood.GoodType = 1 OR tgood.GoodType = 3)
			 )T4
			 ON k.goodcode = T4.goodcode
			 
          GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    T4.FirstDateMojodi ,
                    T4.GoodCode ,
                    T4.Name ,
                    T4.FirstDateUsed
)T1

FULL OUTER JOIN  
(

        SELECT  CAST(SUM(T5.Amount) AS DECIMAL(20,3)) AS Amount,
                T5.[Name],
                T5.GoodCode
                FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                            ( dbo.tFacD.Amount * dbo.tUsePercent.fltUsedValue ) AS Amount,
                            dbo.tGood.Name
							FROM      dbo.tFacM
                            JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                              AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                            JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                    AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                            JOIN dbo.tGood ON dbo.tUsePercent.GoodFirstCode = dbo.tGood.Code
                  WHERE     tfacm.Branch = @Branch2
                            AND tfacm.Recursive = 0
                            AND tfacm.status = 2
                            AND tfacm.Date >= @Date1
                            AND tfacm.Date <= @Date2
                            AND tfacd.intInventoryNo = @intInventory1
							AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
					)T5
			GROUP BY T5.GoodCode,
					T5.Name


)T2

	ON T1.[GoodCode] = T2.[GoodCode]
WHERE 	
FirstMojodi <> 0 OR FromStoreAmount <> 0 OR toStoreAmount <> 0 OR BuyAmount <> 0 OR T2.Amount <> 0

ORDER BY Isnull(T1.[GoodCode] , T2.[GoodCode]) -- ISNULL(T1.Name , T2.Name)

GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FnFirstDateMojodi]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FnFirstDateMojodi]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE  FUNCTION [dbo].[FnFirstDateMojodi]
    (
      @DateAfter VARCHAR(10),
      @inventory INT,
      @Branch INT,
      @AccountYear SMALLINT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT ,
      @GoodCode INT  ,
      @ZeroMojodi INT  
    )
RETURNS @ReturnTable TABLE
    (
      GoodCode INT,
      FirstDateMojodi FLOAT ,
      FirstDateMojodiPrice BIGINT  
    )
AS 
BEGIN 

--Maybe @GoodCode = 0
--Maybe @InventoryNo = 0

IF @GoodLevel11 = 0
 SELECT @GoodLevel11 = MIN(Code) FROM dbo.tGoodLevel1
IF @GoodLevel12 = 0
 SELECT @GoodLevel12 = Max(Code) FROM dbo.tGoodLevel1

    INSERT  INTO @ReturnTable
            (
              GoodCode,
              FirstDateMojodi ,
              FirstDateMojodiPrice                
            )

SELECT M2.GoodCode , M2.FirstDateMojodi ,				 
		CASE WHEN (ISNULL(M2.FirstMojodiTotal,0) + ISNULL(M2.BuyTotal,0) - ISNULL(M2.BuyReturnTotal,0)) = 0 THEN 0
			WHEN  M2.FirstDateMojodi = 0 THEN 0
		ELSE 
		CAST(
		(ISNULL(M2.FirstMojodiPriceTotal ,0) + ISNULL(M2.BuyPriceTotal ,0) - ISNULL(M2.BuyReturnPriceTotal ,0)) 
		/ (ISNULL(M2.FirstMojodiTotal,0) + ISNULL(M2.BuyTotal,0) - ISNULL(M2.BuyReturnTotal,0) )
		AS BIGINT )END AS FirstDateMojodiPrice 

 from 
(
	SELECT M.GoodCode , M.FirstDateMojodi
	 ,
				(SELECT ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) 
							 FROM [dbo].[tFacM] M1
								INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
										AND [M1].[intSerialNo] = [D1].[intSerialNo]
								WHERE  
									M1.[date] < @DateAfter
									AND M1.Status = 1
									AND D1.GoodCode = M.GoodCode
									AND M1.AccountYear = @AccountYear
									--AND M1.Branch = @Branch
									AND Recursive = 0
				  	) AS BuyPriceTotal,
				(SELECT ISNULL(SUM(D2.FeeUnit * D2.Amount ) ,0)
							 FROM [dbo].[tFacM] M2
								INNER JOIN [dbo].[tFacD] D2 ON [M2].[Branch] = [D2].[Branch]
										AND [M2].[intSerialNo] = [D2].[intSerialNo]
								WHERE  
									M2.[date] < @DateAfter
									AND M2.Status = 4
									AND (D2.GoodCode = M.GoodCode)
									AND M2.AccountYear = @AccountYear
									--AND M2.Branch = @Branch
									AND Recursive = 0
				  	) AS BuyReturnPriceTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									[tFacM].[date] < @DateAfter
									AND [tFacM].Status = 1
									AND [tFacD].GoodCode = M.GoodCode
									AND [tFacM].AccountYear = @AccountYear
									--AND [tFacM].Branch = @Branch
									AND [tFacM].Recursive = 0
				  	) AS BuyTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									[tFacM].[date] < @DateAfter
									AND [tFacM].Status = 4
									AND [tFacD].GoodCode = M.GoodCode
									AND [tFacM].AccountYear = @AccountYear
									--AND [tFacM].Branch = @Branch
									AND [tFacM].Recursive = 0
				  	) AS BuyReturnTotal ,
				(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi) ,0)
							 FROM dbo.tInventory_Good
								WHERE  
									 tInventory_Good.GoodCode = M.GoodCode
									AND tInventory_Good.AccountYear = @AccountYear
									AND tInventory_Good.Branch = @Branch
				  	) AS FirstMojodiTotal ,
				(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi * tGood.BuyPrice) ,0)
							 FROM dbo.tInventory_Good
								INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
								WHERE  
									 tInventory_Good.GoodCode = M.GoodCode
									AND tInventory_Good.AccountYear = @AccountYear
									AND tInventory_Good.Branch = @Branch
				  	) AS FirstMojodiPriceTotal 
FROM 
(
        SELECT    
			ISNULL(k.GoodCode,dbo.tInventory_Good.GoodCode) AS GoodCode ,
			CAST(tinventory_good.FirstMojodi
			+ ISNULL(SUM(k.SumBuyAmount), 0)
			- ISNULL(SUM(k.SumFromStoreAmount), 0)
			+ ISNULL(SUM(k.SumtoStoreAmount), 0)
			- ISNULL(SUM(k.SumLossesAmount), 0)
			- ISNULL(SUM(k.SumBuyReturnAmount), 0) AS DECIMAL(20,3)) AS FirstDateMojodi
			
			FROM      (

				SELECT 
				CASE M.Status
				  WHEN 1 THEN SUM(D.Amount)  ELSE 0 END AS SumBuyAmount,
				CASE M.Status
				  WHEN 4 THEN SUM(D.Amount)  ELSE 0 END AS SumBuyReturnAmount,
				CASE M.Status
				  WHEN 6 THEN SUM(D.Amount)  ELSE 0 END AS SumFromStoreAmount,
				CASE M.Status
				  WHEN 7 THEN SUM(D.Amount)  ELSE 0  END AS SumtoStoreAmount,
				CASE M.Status
				  WHEN 3 THEN SUM(D.Amount)  ELSE 0  END AS SumLossesAmount,

				--CASE M.Status 
				--  WHEN 1 THEN D.FeeUnit * SUM(D.Amount)  ELSE 0  END AS BuyPriceTotal,
				--CASE M.Status
				--  WHEN 4 THEN D.FeeUnit * SUM(D.Amount) ELSE 0  END AS BuyReturnPriceTotal,
				--CASE M.Status
				--  WHEN 6 THEN D.FeeUnit * SUM(D.Amount) ELSE 0 END AS FromStorePriceTotal,
				--CASE M.Status
				--  WHEN 7 THEN D.FeeUnit * SUM(D.Amount) ELSE 0 END AS ToStorePriceTotal,
				--CASE M.Status
				--  WHEN 3 THEN D.FeeUnit * SUM(D.Amount) ELSE 0 END AS LossePriceTotal,
				D.GoodCode,
				D.intInventoryNo,
				M.AccountYear,
				M.Branch
			 FROM [dbo].[tFacM] M
				INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
						AND [M].[intSerialNo] = [D].[intSerialNo]
				INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode
						AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
						AND (tGood.Code = @GoodCode OR @GoodCode = 0)
				WHERE  
					M.[date] < @DateAfter
					AND M.Status IN ( 1, 3, 4, 6, 7 )
					AND (D.intInventoryNo = @inventory OR @inventory = 0)
					AND M.AccountYear = @AccountYear
					--AND M.Branch = @Branch
					AND Recursive = 0
				GROUP BY  D.GoodCode,
					D.FeeUnit,
					M.Status,
					D.intInventoryNo,
					M.AccountYear,
					M.Branch

			) k
			
			RIGHT OUTER  JOIN tinventory_good ON k.goodcode = tinventory_good.goodcode
			AND (tinventory_good.Mojodi <> 0 OR @ZeroMojodi = 0)
			AND tinventory_good.inventoryno = k.intInventoryNo 
			AND tinventory_good.AccountYear = k.AccountYear
			--AND tinventory_good.Branch = k.Branch
			INNER JOIN dbo.tGood ON dbo.tGood.Code = tinventory_good.GoodCode 
				AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12 
				AND (tGood.Code = @GoodCode OR @GoodCode = 0)
		  WHERE     tinventory_good.AccountYear = @AccountYear
					AND tinventory_good.Branch = @Branch
					AND (tinventory_good.InventoryNo = @inventory OR @inventory = 0) 
          GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    k.Branch,
                    tinventory_good.FirstMojodi,
                    tinventory_good.firstPrice,
                    tinventory_good.GoodCode ,
                    tGood.BuyPrice ,
                    tinventory_good.AccountYear ,
                    tinventory_good.Branch
        
)M
)M2
	ORDER BY goodcode


    RETURN
   END
--==========================================

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_InventoryAtomicRials_Report_InDate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_InventoryAtomicRials_Report_InDate]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


CREATE  PROCEDURE [dbo].[Get_InventoryAtomicRials_Report_InDate]
    (
      @SystemDate NVARCHAR(10) ,
      @SystemDay NVARCHAR(10) ,
      @SystemTime NVARCHAR(10) ,
      @Date1 NVARCHAR(50) ,
      @Date2 NVARCHAR(50) ,
      @AccountYear1 SMALLINT ,
      @Branch1 INT ,
      @Inventory1 INT ,
      @GoodLevel11 INT ,
      @GoodLevel12 INT 
    )
AS 

 BEGIN 

	SELECT DISTINCT X1.* ,
			CASE WHEN (X1.FirstMojodiTotal + X1.BuyTotal - X1.BuyReturnTotal) = 0 THEN X1.BuyPrice  
				WHEN  X1.Mojodi = 0 THEN 0
			ELSE 
			CAST(
			((X1.FirstMojodiTotal * X1.BuyPrice) + X1.BuyPriceTotal - X1.BuyReturnPriceTotal)  
			/ (X1.FirstMojodiTotal + X1.BuyTotal - X1.BuyReturnTotal) 
			AS BIGINT )END AS AvrageBuyPrice ,
		  @SystemDay + N' ' + @SystemDate + N' در ساعت' + @SystemTime AS SysDate ,
		(select [Description] FROM dbo.tInventory WHERE InventoryNo = @Inventory1) AS Inventory_Description ,
 		@Date1 AS DateBefore ,
		@Date2 AS DateAfter 
From (
  SELECT
		 X.* , tGood.Name , tGood.Code , tGood.BuyPrice ,
				(SELECT ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) 
							 FROM [dbo].[tFacM] M1
								INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
										AND [M1].[intSerialNo] = [D1].[intSerialNo]
								WHERE  
									    M1.[Date] <= @Date2
									AND M1.Status = 1
									AND D1.GoodCode = X.GoodCode
									AND M1.AccountYear = @AccountYear1
									--AND M1.Branch = @Branch1
									AND Recursive = 0
				  	) AS BuyPriceTotal,
				(SELECT ISNULL(SUM(D2.FeeUnit * D2.Amount ) ,0)
							 FROM [dbo].[tFacM] M2
								INNER JOIN [dbo].[tFacD] D2 ON [M2].[Branch] = [D2].[Branch]
										AND [M2].[intSerialNo] = [D2].[intSerialNo]
								WHERE  
									    M2.[Date] <= @Date2
									AND M2.Status = 4
									AND (D2.GoodCode = X.GoodCode)
									AND M2.AccountYear = @AccountYear1
									--AND M2.Branch = @Branch1
									AND Recursive = 0
				  	) AS BuyReturnPriceTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									    [tFacM].[date] <= @Date2
									AND [tFacM].Status = 1
									AND [tFacD].GoodCode = X.GoodCode
									AND [tFacM].AccountYear = @AccountYear1
									--AND [tFacM].Branch = @Branch1
									AND [tFacM].Recursive = 0
				  	) AS BuyTotal ,
				(SELECT ISNULL(SUM([tFacD].Amount) ,0)
							 FROM [dbo].[tFacM] 
								INNER JOIN [dbo].[tFacD] ON [tFacM].[Branch] = [tFacD].[Branch]
										AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
								WHERE  
									    [tFacM].[date] <= @Date2
									AND [tFacM].Status = 4
									AND [tFacD].GoodCode = X.GoodCode
									AND [tFacM].AccountYear = @AccountYear1
									--AND [tFacM].Branch = @Branch1
									AND [tFacM].Recursive = 0
				  	) AS BuyReturnTotal ,
				(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi) ,0)
							 FROM dbo.tInventory_Good
								WHERE  
									 tInventory_Good.GoodCode = X.GoodCode
									AND tInventory_Good.AccountYear = @AccountYear1
									AND tInventory_Good.Branch = @Branch1
				  	) AS FirstMojodiTotal --,
				--(SELECT ISNULL(SUM(tInventory_Good.FirstMojodi * tGood.BuyPrice) ,0)
				--			 FROM dbo.tInventory_Good
				--				INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
				--  					AND tInventory_Good.GoodCode = X.GoodCode
				--					AND tInventory_Good.AccountYear = @AccountYear1
				--					AND tInventory_Good.Branch = @Branch1
				--  	) AS FirstMojodiPriceTotal 
		
 FROM (
         SELECT    
			T.FirstDateMojodi AS FirstMojodi ,
			T.FirstDateMojodiPrice AS FirstDateMojodiPrice ,
			CASE WHEN T.FirstDateMojodi = 0 THEN 0 
			ELSE CAST(T.FirstDateMojodiPrice * T.FirstDateMojodi  AS BIGINT )END  AS FirstDatePrice ,
			T.FirstDateMojodi
			+ ISNULL(SUM(k.TotalBuy), 0)
			- ISNULL(SUM(k.TotalHavaleh), 0)
			+ ISNULL(SUM(k.TotalResid), 0)
			- ISNULL(SUM(k.TotalLoss), 0)
			- ISNULL(SUM(k.TotalBuyReturn), 0) AS Mojodi,
			--CASE WHEN (T.FirstDateMojodi + ISNULL(SUM(k.TotalBuy), 0) - ISNULL(SUM(k.TotalBuyReturn), 0)) = 0 THEN K.BuyPrice  
			--ELSE 
			--CAST((
			--(T.FirstDateMojodiPrice * T.FirstDateMojodi)
			--+ ISNULL(SUM(k.TotalBuyPrice), 0) - ISNULL(SUM(k.TotalBuyReturnPrice), 0))  
			--/ (T.FirstDateMojodi + ISNULL(SUM(k.TotalBuy), 0)  - ISNULL(SUM(k.TotalBuyReturn), 0) ) 
			--AS BIGINT )END AS AvrageBuyPrice ,
			ISNULL(k.GoodCode,T.GoodCode) AS GoodCode ,
			
			ISNULL(SUM(k.TotalBuy), 0) AS TotalBuy ,
			ISNULL(SUM(k.TotalSell), 0) AS TotalSell,
			ISNULL(SUM(k.TotalLoss), 0) AS TotalLoss ,
			ISNULL(SUM(k.TotalBuyReturn), 0) AS TotalBuyReturn,
			ISNULL(SUM(k.TotalSellReturn), 0) AS TotalSellReturn,
			ISNULL(SUM(k.TotalHavaleh), 0) AS TotalHavaleh,
			ISNULL(SUM(k.TotalResid), 0) AS TotalResid,

			ISNULL(SUM(k.TotalBuyPrice), 0) AS TotalBuyPrice,
			ISNULL(SUM(k.TotalSellPrice), 0)AS TotalSellPrice ,
			ISNULL(SUM(k.TotalLossPrice), 0) AS TotalLossPrice,
			ISNULL(SUM(k.TotalBuyReturnPrice), 0) AS TotalBuyReturnPrice,
			ISNULL(SUM(k.TotalSellReturnPrice), 0) AS TotalSellReturnPrice,
			ISNULL(SUM(k.TotalHavalehPrice), 0) AS TotalHavalehPrice ,
			ISNULL(SUM(k.TotalResidPrice), 0) AS TotalResidPrice 

			FROM   
			   (
                 
				SELECT   
				DISTINCT [D].[GoodCode] ,D.intinventoryNo,
				M.AccountYear,
				M.Branch ,
				--------------------------------------------------------------------------------------------مقدار موجودي در  حالتهاي مختلف ----------------------------------------
				CASE WHEN [M].Status = 1 THEN SUM( [D].[Amount])ELSE 0 END AS TotalBuy ,
				CASE WHEN [M].Status = 2 THEN SUM( [D].[Amount])ELSE 0 END AS TotalSell ,
				CASE WHEN [M].Status = 3 THEN SUM([D].[Amount]) ELSE 0 END AS TotalLoss ,
				CASE WHEN [M].Status = 4 THEN SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturn ,
				CASE WHEN [M].Status = 5 THEN SUM([D].[Amount]) ELSE 0 END AS TotalSellReturn ,
				CASE WHEN [M].Status = 6 THEN SUM([D].[Amount]) ELSE 0 END AS TotalHavaleh ,
				CASE WHEN [M].Status = 7 THEN SUM([D].[Amount]) ELSE 0 END AS TotalResid ,

				CASE WHEN [M].Status = 1 THEN D.FeeUnit * SUM( [D].[Amount])ELSE 0 END AS TotalBuyPrice ,
				CASE WHEN [M].Status = 2 THEN D.FeeUnit * SUM( [D].[Amount])ELSE 0 END AS TotalSellPrice ,
				CASE WHEN [M].Status = 3 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalLossPrice ,
				CASE WHEN [M].Status = 4 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalBuyReturnPrice ,
				CASE WHEN [M].Status = 5 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalSellReturnPrice ,
				CASE WHEN [M].Status = 6 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalHavalehPrice ,
				CASE WHEN [M].Status = 7 THEN D.FeeUnit * SUM([D].[Amount]) ELSE 0 END AS TotalResidPrice ,

				tGood.BuyPrice
				-----------------------------------------------------------------------------------------------------------------------------------------------------------------
				FROM      [dbo].[tFacM] M
						INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
								AND [M].[intSerialNo] = [D].[intSerialNo]
						INNER JOIN dbo.tGood ON dbo.tGood.Code = D.GoodCode 
								AND tGood.Level1 >= @GoodLevel11 AND tGood.Level1 <= @GoodLevel12
				WHERE     [M].[Date] >= @Date1
						AND [M].[Date] <= @Date2
						AND [M].[AccountYear] = @AccountYear1
						--AND [M].[Branch] = @Branch1
						AND [M].[Recursive] = 0
						AND [D].[intInventoryNo] = @Inventory1  
				GROUP BY  [D].[GoodCode] ,
						[M].[Status] ,
						D.intInventoryNo ,
						D.FeeUnit ,
						M.AccountYear,
						M.Branch ,
						tGood.BuyPrice
			) K

			RIGHT OUTER  JOIN 
			(SELECT  GoodCode, FirstDateMojodi ,FirstDateMojodiPrice
                FROM    dbo.FnFirstDateMojodi(@Date1, @Inventory1, @Branch1,
                                              @AccountYear1 , @GoodLevel11 , @GoodLevel12 , 0 , 0) 
			)T 
			 ON k.goodcode = T.goodcode
          GROUP BY 
                    k.GoodCode,
                    k.intInventoryNo,
                    k.AccountYear,
                    k.Branch ,
                    T.FirstDateMojodi ,
                    T.FirstDateMojodiPrice ,
                    T.GoodCode ,
                    K.BuyPrice
    )X
    INNER JOIN tgood ON X.GoodCode = tGood.Code  AND   X.mojodi <> 0

)X1

	ORDER BY   goodcode --Name

    END


GO
--------------------------------گزارش مانده انبار--- Step 9


ALTER   PROC [dbo].[Get_tblTotal_tGood_By_Prams_Report]
    (
      @SystemDate NVARCHAR(10),
      @SystemDay NVARCHAR(10),
      @SystemTime NVARCHAR(10),
      @AccountYear1 SMALLINT,
      @Level11 INT,
      @Level12 INT,
      @Level21 INT,
      @Level22 INT,
      @Inventory1 INT,
      @Inventory2 INT,
      @SortOrder1 INT,
      @Branch1 INT
    )
AS 
    SELECT  @SystemDay + N' ' + @SystemDate + N' در ساعت' + @SystemTime AS SysDate,
            vw_Good.Code,
            vw_Good.SellPrice,
            [dbo].[vw_Good].[SellPrice] AS LastSellPrice,
            vw_Good.SellPrice2,
            vw_Good.SellPrice3,
            vw_Good.SellPrice4,
            vw_Good.SellPrice5,
            vw_Good.SellPrice6,
            vw_Good.BuyPrice,
            [dbo].[vw_Good].[BuyPrice] AS LastBuyPrice,
            vw_Good.BarCode,
            vw_Good.[Name],
            vw_Good.Unit,
            vw_Good.UnitDescription,
            vw_Good.TypeDescription,
            dbo.vw_Good.TechnicalNo,
            tInventory_Good.GoodCode,
            tInventory_Good.InventoryNo,
            tInventory_Good.Branch,
            tInventory_Good.Mojodi,
            tInventory_Good.MojodiControl,
            tInventory_Good.OrderPoint,
            tInventory_Good.MinValue,
            tInventory_Good.MaxValue,
            tInventory_Good.BuyPriceAverage,
            tInventory_Good.SalePriceAverage,
            tInventory_Good.BuyAmount,
            tInventory_Good.SaleAmount,
            tInventory_Good.Counting1,
            tInventory_Good.Counting2,
            tInventory_Good.Counting3,
            tInventory_Good.AccountYear,
            tInventory_Good.CountDifference,
            tInventory_Good.LossAmount,
            tInventory_Good.BuyReturnAmount,
            tInventory_Good.SaleReturnAmount,
            tInventory_Good.FromStoreAmount,
            tInventory_Good.ToStoreAmount,
            tInventory_Good.FirstPrice,
            tInventory_Good.MojodiPrice,
            tInventory_Good.FirstMojodi,
            [dbo].[tInventory].[Description] AS InventoryDescription
    FROM    [dbo].[vw_Good]
            INNER JOIN [dbo].[tInventory_Good] ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode
            INNER JOIN [dbo].[tInventory] ON [dbo].[tInventory_Good].[Branch] = [dbo].[tInventory].[Branch]
                                             AND [dbo].[tInventory_Good].[InventoryNo] = [dbo].[tInventory].[InventoryNo]
    WHERE   ( [dbo].[vw_Good].[Level1] BETWEEN @Level11
                                       AND     @Level12 )
            AND ( [dbo].[tInventory_Good].[InventoryNo] BETWEEN @Inventory1
                                                        AND     @Inventory2 )
            AND [dbo].[tInventory_Good].[Branch] = @Branch1
            AND AccountYear = @AccountYear1
            AND vw_good.LEVEL2 BETWEEN @Level21 AND @Level22
            AND (FirstMojodi <> 0 OR FromStoreAmount <> 0 OR toStoreAmount <> 0 OR BuyAmount <> 0 OR BuyReturnAmount <> 0)
    ORDER BY CASE @SortOrder1
               WHEN 1 THEN [dbo].[vw_Good].[Code]
               WHEN 2 THEN Barcode
               WHEN 3 THEN [Name]
               WHEN 4 THEN Unit
               WHEN 5 THEN Mojodi
               WHEN 6 THEN Sellprice
               WHEN 7 THEN BuyPrice
               WHEN 8 THEN Counting1
             END

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[GetInventoryGood_AllMojodi_Report]
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
        DECLARE @Amountv INT   
        DECLARE @Amounts INT  
        DECLARE @feev INT  
        DECLARE @fees INT   
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
              FirstDateMojodi INT
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
              Amountv INT,
              Amounts INT,
              feev INT,
              fees INT,
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
                        ABS(Amountv) AS Amountv,
                        ABS(Amounts) AS Amounts,
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
                @SystemDay + ' ' + @SystemDate + ' ' + N' ساعت : '
                + @SystemTime AS Sysdate
        FROM    @tblReturnDateMojodi
    END

GO
