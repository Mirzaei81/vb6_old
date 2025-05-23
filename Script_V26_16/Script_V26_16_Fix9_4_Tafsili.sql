



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

