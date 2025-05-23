

--Script_V26_16_Fix10_bitActiveDifference
--اضافه کردن فیلد برای کنترل رکوردهاییکه باید در محاسبه
--مغایرت گیری شرکت کنند
--فقط کالاهاییکه فیلد فعال آنها تیک داشته باشد در مغایرت گیری شرکت می کنند
--سپس سند کسر و اضافه انبار برای این کالاها صادر می گردد
--آپدیت کردن قیمت خرید کالا با آخرین قیمت خرید آن کالا


IF COL_LENGTH('[tInventory_Good]','bitActiveDifference') IS NULL
	ALTER TABLE dbo.tInventory_Good ADD  bitActiveDifference BIT NULL 

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Update_tblTotal_tInventory_Good_By_Counting
(
	@Code		INT ,
	@InventoryNo INT,
	@Counting1	Float,
	@Counting2	Float,
	@Counting3	Float ,
	@Branch INT	,
	@AccountYear	SMALLINT ,
	@bitActiveDifference BIT 
	
)

AS
	
    UPDATE dbo.tInventory_Good

	SET    Counting1    = @Counting1 ,
	       Counting2 = @Counting2 ,
	       Counting3  = @Counting3 ,
	       [Time] = dbo.setTimeFormat(getdate()) ,
	       [Date] = dbo.Shamsi(GETDATE()) ,
	       bitActiveDifference = @bitActiveDifference
	Where GoodCode = @Code And InventoryNo = @InventoryNo And Branch = @branch And AccountYear = @accountYear
	

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_bitActiveDifference]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_bitActiveDifference]
GO


CREATE  PROCEDURE dbo.Update_bitActiveDifference
AS 

    UPDATE dbo.tInventory_Good
		SET bitActiveDifference = 0


GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO






ALTER  proc Update_tblTotal_tGood_By_Prams 
(
	@Level1 int ,
        @strSelectedLevels nvarchar(4000) , 
	@InventoryNo int ,
        @Branch int ,
	@AccountYear SMALLINT,
	@CheckNotZeroMojodi BIT,
	@CheckFirstMojodi	BIT,
	@CountingNo	INT
)
	
as
	UPDATE tInventory_Good
		SET tInventory_Good.CountDifference=(CASE @CountingNo 
							WHEN 1 THEN ISNULL(T.Counting1,0)
							WHEN 2 THEN ISNULL(T.Counting2,0)
							WHEN 3 THEN ISNULL(T.Counting3,0)
							ELSE ISNULL(T.Counting1,0)
							END)
							-T.Mojodi
	FROM
	(
		SELECT vw_Good.* , tInventory_Good.* 
		
		FROM 
			[dbo].[vw_Good] 
			Inner Join  
			dbo.tInventory_Good ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode 
		WHERE 
			(LEVEL1 = @Level1 OR @Level1=-1)
			And (InventoryNo = @InventoryNo OR @InventoryNo=-1)
			And (Branch = @Branch OR @Branch=-1)
			And (AccountYear = @AccountYear OR @AccountYear=-1)
			AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels='')
			AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
			AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
	)AS T
	WHERE T.GoodCode=tInventory_Good.GoodCode 
		AND T.InventoryNo=tInventory_Good.InventoryNo
		AND T.Branch=tInventory_Good.Branch
		AND T.AccountYear = tInventory_Good.AccountYear
		AND T.bitActiveDifference = 1

GO





if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_BuyPrice_by_LastPrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_BuyPrice_by_LastPrice
GO


CREATE  PROCEDURE dbo.Update_BuyPrice_by_LastPrice
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

