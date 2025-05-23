

--مقصد دوم حواله و درست کردن گزارشات آن و فاکتور خرید و حواله
--Destination &
--Script_V26_16_Fix9
--93/09/09


IF COL_LENGTH('tFacM','DestinationId') IS NULL
BEGIN
	ALTER TABLE dbo.tFacM
	ADD DestinationId INT NULL 
END

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    VIEW dbo.vw_FacMD_Good
AS
SELECT     dbo.tFacD.Amount, dbo.tFacD.GoodCode, dbo.tFacD.FeeUnit, dbo.tFacD.ServePlace, dbo.tFacD.DifferencesCodes, dbo.tFacD.DifferencesDescription, 
                      dbo.tFacD.Discount, dbo.tGood.Name, dbo.tGood.LatinName, dbo.tGood.Unit, dbo.tGood.Weight, dbo.tFacM.[No], dbo.tFacM.Status, 
                      dbo.tFacM.OrderType, dbo.tFacM.FacPayment, dbo.tFacD.Rate, dbo.tFacD.ChairName, dbo.tFacD.intInventoryNo, dbo.tFacD.DestInventoryNo, 
                      dbo.tFacD.ExpireDate, dbo.tFacM.AccountYear, dbo.tFacM.NvcDescription , dbo.tfacd.introw  , tGood.NumberOfUnit ,  tGood.MainType
			, dbo.tFacM.Branch,isnull(dbo.tfacm.TempAddress,'')as TempAddress , tUnitGood.[Description] , dbo.tGood.TaxBuy , TaxSale , DutyBuy , DutySale , ISNULL(GuestNo , '') AS GuestNo
			, tfacm.DestinationId
FROM         dbo.tFacM INNER JOIN
                      dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND dbo.tFacM.Branch = dbo.tFacD.Branch INNER JOIN
                      dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code INNER JOIN
                      dbo.[tUnitGood] ON [tGood].[Unit] = [tUnitGood].[Code]
                      
--WHERE     (dbo.tFacM.Branch = dbo.Get_Current_Branch())



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_FacMD_Good] (@No Bigint , @Status int , @intLanguage int , @AccountYear Smallint  , @Branch INT ) 

AS

Select Sum(vw_FacMD_Good.Amount)As Amount  , vw_FacMD_Good.GoodCode    , Max(vw_FacMD_Good.ServePlace) As Serveplace , Max(vw_FacMD_Good.DifferencesCodes) As DifferencesCodes  ,
	Max(vw_FacMD_Good.DifferencesDescription) As DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]  ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	MAX(vw_FacMD_Good.Rate) as rate ,vw_FacMD_Good.ChairName  , Max(vw_FacMD_Good.FeeUnit)As FeeUnit ,
	Max(vw_FacMD_Good.intinventoryNo)As intinventoryNo ,Max(vw_FacMD_Good.DestInventoryNo)As DestInventoryNo,Max(vw_FacMD_Good.[ExpireDate])As [ExpireDate] , Max( IsNull(vw_FacMD_Good.[NvcDescription], ''))As [NvcDescription]
        , case @intLanguage when 0 then Name 
			    when 1 then LatinName end as nvcName , vw_FacMD_Good.intRow
	,Max(vw_FacMD_Good.NumberOfUnit) As NumberOfUnit , vw_FacMD_Good.maintype,Max( IsNull(vw_FacMD_Good.[TempAddress], ''))As [TempAddress]
	, Max(vw_FacMD_Good.[Description]) AS UnitDescription , ISNULL(T.Mojodi , 0 ) AS Mojodi
	 , TaxBuy , TaxSale , DutyBuy , DutySale , ISNULL(DestinationId , 0) AS DestinationId
 	from vw_FacMD_Good 
	LEFT OUTER JOIN
	(SELECT InventoryNo , GoodCode , Branch , AccountYear ,Mojodi FROM dbo.tInventory_Good)T ON T.AccountYear = @AccountYear 
	AND t.Branch = @Branch AND t.GoodCode = vw_FacMD_Good.GoodCode AND t.InventoryNo = vw_FacMD_Good.intinventoryNo
 	 
	where No = @No  And  Status = @Status And  vw_FacMD_Good.AccountYear =  @AccountYear AND vw_FacMD_Good.Branch = @Branch
	 Group By vw_FacMD_Good.GoodCode     , vw_FacMD_Good.DifferencesCodes ,vw_FacMD_Good.DifferencesDescription   ,
	vw_FacMD_Good.Discount  , vw_FacMD_Good.Name  , vw_FacMD_Good.LatinName  ,  vw_FacMD_Good.Unit   , vw_FacMD_Good.Weight , 
	vw_FacMD_Good.[No]    ,  vw_FacMD_Good.OrderType  ,vw_FacMD_Good.Facpayment  ,
	vw_FacMD_Good.ChairName   , vw_FacMD_Good.ServePlace , vw_FacMD_Good.FeeUnit  , vw_FacMD_Good.intRow , vw_FacMD_Good.MainType
	,T.Mojodi  , TaxBuy , TaxSale , DutyBuy , DutySale , DestinationId
Order By  vw_FacMD_Good.intRow




GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tFacM_Destination]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_tFacM_Destination
GO


CREATE PROCEDURE [dbo].[Update_tFacM_Destination] (
	@intSerialNo BIGINT  ,
	@DestinationId INT 
	)

AS
IF @DestinationId = 0 SET @DestinationId = NULL

UPDATE dbo.tFacM SET
	DestinationId = @DestinationId

   WHERE intSerialNo = @intSerialNo



GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPub_Destination]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPub_Destination]
GO


CREATE TABLE [dbo].[tblPub_Destination]
(
DestinationId [int] NOT NULL IDENTITY(1,1) ,
[NvcDestination] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
) ON [PRIMARY]
GO


ALTER TABLE [dbo].[tblPub_Destination] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblPub_Destination] PRIMARY KEY  CLUSTERED 
	(
		[DestinationId]
	)  ON [PRIMARY] 
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insert_tblPub_Destination]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Insert_tblPub_Destination]
GO


CREATE PROCEDURE [dbo].[Insert_tblPub_Destination] 
(
	@NvcDestination nvarchar(50) ,
	@intStatus int out)
AS

BEGIN TRAN 
Insert Into dbo.tblPub_Destination
        ( NvcDestination 
        )
VALUES  ( @NvcDestination 
        )
if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus=@@IDENTITY
RETURN @intStatus

ErrHandler:
RollBack Tran
Set @intStatus = -1
RETURN @intStatus


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tblPub_Destination]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_tblPub_Destination]
GO

CREATE PROCEDURE [dbo].[Update_tblPub_Destination] (
	@DestinationId INT ,
	@NvcDestination nvarchar(50) ,
	@intStatus INT OUT 
	)

AS

UPDATE dbo.tblPub_Destination SET
	NvcDestination = @NvcDestination

   WHERE DestinationId = @DestinationId

Set @intStatus = 1
RETURN @intStatus


GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tblPub_Destination]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tblPub_Destination]
GO


CREATE PROCEDURE [dbo].[Get_All_tblPub_Destination] AS
select * from [dbo].[tblPub_Destination] ORDER BY DestinationId

GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_tblPub_Destination_ById]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_tblPub_Destination_ById]
GO


CREATE PROCEDURE [dbo].[Get_tblPub_Destination_ById] 
@DestinationId INT 
AS
select * from [dbo].[tblPub_Destination] WHERE  DestinationId = @DestinationId

GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Delete_tblPub_Destination]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Delete_tblPub_Destination]
GO

CREATE PROCEDURE [dbo].[Delete_tblPub_Destination](
	@DestinationId INT)
AS
	DELETE FROM dbo.tblPub_Destination WHERE DestinationId = @DestinationId

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FnFirstDateMojodi_ByDestinationId]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FnFirstDateMojodi_ByDestinationId]
GO


CREATE FUNCTION [dbo].[FnFirstDateMojodi_ByDestinationId]
    (
      @DateAfter VARCHAR(10),
      @AccountYear SMALLINT ,
      @DestinationId1 INT ,
      @DestinationId2 INT    
    )
RETURNS @ReturnTable TABLE
    (
      DestinationId INT ,
      GoodCode INT,
      FirstDateMojodi INT 
    )
AS 
BEGIN 

    INSERT  INTO @ReturnTable
            (
              DestinationId ,
              GoodCode,
              FirstDateMojodi                
            )


--DECLARE @DateAfter VARCHAR(10)
--DECLARE @Branch INT
--DECLARE @AccountYear SMALLINT   

--SET  @DateAfter = N'92/11/01'
--SET @inventory = 1
--SET @Branch  = 1
--SET @AccountYear = 1392

	select Y.DestinationId , Y.GoodCode , SUM(Y.Import) - SUM(Y.Export)  FROM 
	(SELECT 
			dbo.tFacM.DestinationId ,tblPub_Destination.NvcDestination 
			,CASE WHEN tfacd.intInventoryNo = 1 THEN  tfacd.Amount ELSE 0 END AS Import
			,CASE WHEN tfacd.intInventoryNo <> 1 THEN  tfacd.Amount ELSE 0 END AS Export
			,tfacd.GoodCode , tgood.Name 
			, (SELECT [Description] FROM dbo.tInventory WHERE InventoryNo = tfacd.intInventoryNo)  AS InventoryDescription
			, (SELECT [Description] FROM dbo.tInventory WHERE InventoryNo = tfacd.DestInventoryNo) AS DestInventoryDescription
			
		FROM dbo.tFacM
			INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
			INNER JOIN dbo.tblPub_Destination ON dbo.tblPub_Destination.DestinationId = dbo.tFacM.DestinationId
			INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
			
		WHERE 
			dbo.tFacM.Status = 6
			and AccountYear = @AccountYear
			AND dbo.tFacM.Date < @DateAfter
			AND dbo.tFacM.DestinationId >= @DestinationId1
			AND dbo.tFacM.DestinationId <= @DestinationId2
		)Y
	GROUP BY Y.DestinationId , Y.GoodCode
    RETURN
   END
--==========================================
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].GetGoods_ByDestinationId') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].GetGoods_ByDestinationId
GO


CREATE   Proc GetGoods_ByDestinationId (
   
      @SystemDate NVARCHAR(10) ,
      @SystemDay NVARCHAR(10) ,
      @SystemTime NVARCHAR(10) ,
      @Date1 NVARCHAR(50) ,
      @Date2 NVARCHAR(50) ,
      @AccountYear1 SMALLINT ,
      --@GoodLevel11 INT ,
      --@GoodLevel12 INT ,
      @DestinationId1 INT , 
      @DestinationId2 INT  
    )
AS 

		SELECT 
		'' AS SysDate ,
		X.DestinationId AS DestinationId,
		FirstDateMojodi AS Import ,
		0 AS Export ,
		GoodCode ,
		tgood.Name ,
		N'موجودی اول دوره' AS InventoryDescription ,
		'' AS DestInventoryDescription ,
		tblPub_Destination.NvcDestination AS NvcDestination ,
		'' AS Date ,
		0 AS intserialNo
		FROM 
		(SELECT GoodCode , DestinationId , FirstDateMojodi FROM dbo.FnFirstDateMojodi_ByDestinationId( @Date1,@AccountYear1 , @DestinationId1 , @DestinationId2))X
		--ON Y.DestinationId = X.DestinationId AND Y.GoodCode = X.GoodCode
		INNER JOIN dbo.tGood ON dbo.tGood.Code = x.GoodCode
		INNER JOIN dbo.tblPub_Destination ON dbo.tblPub_Destination.DestinationId = X.DestinationId

UNION 
	SELECT  
		Y.SysDate ,
		y.DestinationId  ,
		y.Import  ,
		y.Export    ,
		y.GoodCode    ,
		y.Name    ,
		y.InventoryDescription  ,
		y.DestInventoryDescription  ,
		y.NvcDestination  ,
		y.Date ,
		Y.intSerialNo
	FROM 
(
	SELECT @SystemDay + N' ' + @SystemDate + N' در ساعت' + @SystemTime AS SysDate ,
			dbo.tFacM.DestinationId ,tblPub_Destination.NvcDestination 
			,CASE WHEN (tfacd.intInventoryNo = 1 OR tfacd.intInventoryNo = 3 OR tfacd.intInventoryNo = 70) THEN  tfacd.Amount 
			     -- WHEN tfacd.intInventoryNo = 3 THEN  tfacd.Amount 
			     -- WHEN tfacd.intInventoryNo = 70 THEN  tfacd.Amount 
			      ELSE 0 END AS Import
			,CASE WHEN (tfacd.intInventoryNo = 1 OR tfacd.intInventoryNo = 3 OR tfacd.intInventoryNo = 70) THEN 0 ELSE tfacd.Amount  END AS Export
			,tfacd.GoodCode , tgood.Name 
			, (SELECT [Description] FROM dbo.tInventory WHERE InventoryNo = tfacd.intInventoryNo)  AS InventoryDescription
			, (SELECT [Description] FROM dbo.tInventory WHERE InventoryNo = tfacd.DestInventoryNo) AS DestInventoryDescription
			, dbo.tFacM.Date , tFacm.intSerialNo
		FROM dbo.tFacM
			INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
			INNER JOIN dbo.tblPub_Destination ON dbo.tblPub_Destination.DestinationId = dbo.tFacM.DestinationId
			INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
		WHERE 
			dbo.tFacM.Status = 6
			and AccountYear = @AccountYear1
			AND dbo.tFacM.Date >= @Date1
			AND dbo.tFacM.Date <= @Date2
			AND dbo.tFacM.DestinationId >= @DestinationId1
			AND dbo.tFacM.DestinationId <= @DestinationId2
			--AND dbo.tGood.Level1 >= @GoodLevel11			
			--AND dbo.tGood.Level1 <= @GoodLevel12			

		)Y
		ORDER BY Y.intSerialNo
GO




ALTER   proc Transport_tblTotal_tGood_By_Prams 
(
	@Level1 int ,
	@strSelectedLevels nvarchar(4000) , 
	@InventoryNo int ,
	@Branch int ,
	@AccountYear SMALLINT,
	@CheckNotZeroMojodi BIT,
	@CheckFirstMojodi	BIT,
	@CountingNo	INT,
	@ToOtherAccountYear SMALLINT
)
	
as
BEGIN TRAN

	DELETE tInventory_Good
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
			And (AccountYear = @ToOtherAccountYear OR @ToOtherAccountYear=-1)
			AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels=N'')
			AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
			AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
	)AS T
	WHERE T.GoodCode=tInventory_Good.GoodCode 
		AND T.InventoryNo=tInventory_Good.InventoryNo
		AND T.Branch=tInventory_Good.Branch
		AND T.AccountYear=tInventory_Good.AccountYear

if @@Error <> 0 
	Goto ErrHandler

	INSERT INTO tInventory_Good
		(
		      InventoryNo, Branch, GoodCode, FirstMojodi, Mojodi, MojodiControl, OrderPoint, MinValue, MaxValue, [Date], [Time], BuyAmount, SaleAmount, 
                      LossAmount, BuyReturnAmount, SaleReturnAmount, FromStoreAmount, toStoreAmount, AccountYear, Counting1, 
                      Counting2, Counting3, CountDifference
		)

		SELECT     InventoryNo, Branch, GoodCode,CASE @CountingNo 
								WHEN 0 THEN CAST(ISNULL(T.Mojodi  ,0) AS DECIMAL(20,3))
								WHEN 1 THEN ISNULL(T.Counting1,0)
								WHEN 2 THEN ISNULL(T.Counting2,0)
								WHEN 3 THEN ISNULL(T.Counting3,0)
								ELSE ISNULL(T.Mojodi,0)
								END
							AS FirstMojodi, 0, 0, 0, MinValue, MaxValue, [Date], [Time], 0, 0, 
	                      0, 0, 0, 0, 0, @ToOtherAccountYear, 0, 
	                      0, 0, 0
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
	
if @@Error <> 0 
	Goto ErrHandler

	Insert into dbo.tStation_Inventory_Good ( branch ,InventoryNo, AccountYear ,StationID,  GoodCode , Active)
	
	select Branch ,InventoryNo ,@ToOtherAccountYear ,StationID ,GoodCode ,Active 
		From tStation_Inventory_Good 
	        Where   inventoryno = @inventoryno and Branch = @Branch and AccountYear = @AccountYear

if @@Error <> 0 
	Goto ErrHandler


Commit Tran
Return

ErrHandler:
RollBack Tran
Return



GO

INSERT INTO dbo.tbltotal_ItemReports
        ( intReportId ,
          intGroupReportId ,
          ReportName ,
          LatinReportName ,
          Refrence_Sp
        )
VALUES  ( 96 , -- intReportId - int
          2 , -- intGroupReportId - int
          N'گزارش کالاها بر اساس مقصد' , -- ReportName - nvarchar(100)
          'RepGetGoods_ByDestinationId' , -- LatinReportName - varchar(50)
          'GetGoods_ByDestinationId'  -- Refrence_Sp - varchar(50)
        )

GO

DELETE FROM tblTotal_ItemReports_Details WHERE intReportId = 96
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
SELECT 96 ,
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
          FROM tblTotal_ItemReports_Details WHERE intReportId = 95 AND (Row = 1 OR Row = 2)
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
VALUES  ( 96 , -- intReportId - int
          3 , -- Row - tinyint
          N'از مقصد' , -- FromText - nvarchar(20)
          N'تا مقصد' , -- toText - nvarchar(20)
          'DestinationId' , -- ParameterName - varchar(20)
          5 , -- ParameterType - tinyint
          4 , -- parameterLengh - int
          1 , -- ObjectType - tinyint
          2 , -- Quantity - tinyint
          '' , -- MinValue - varchar(10)
          '' , -- MaxValue - varchar(10)
          'select * from tblPub_Destination ORDER BY DestinationId' , -- ComboQuery - text
          'DestinationId' , -- ComboFieldCode - varchar(50)
          N'NvcDestination' , -- ComboFieldDescr - nvarchar(50)
          0  -- RighttoLeft - bit
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
VALUES  ( 450 , -- intObjectCode - int
          N'RepGetGoods_ByDestinationId' , -- ObjectId - nvarchar(50)
          N'گزارش کالاها بر اساس مقصد' , -- ObjectName - nvarchar(50)
          N'RepGetGoods_ByDestinationId' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          108  -- ObjectParent - int
        )

GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          450  -- intObjectCode - int
          )

GO



ALTER    PROCEDURE [dbo].[GetOrderGoodAmountInfo]
    (
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @level11 INT,
      @level12 INT,
      @level21 INT,
      @level22 INT,
      @Inventory1 INT,
      @AccountYear1  INT  
    )

AS 
    DECLARE @intLanguage INT 
    SET @intLanguage = 0
    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage = 0 
        SET @TimeTitle = N' ساعت : '
    ELSE 
        SET @TimeTitle = N'Time: '


    SELECT  [CompDes],
            [Code],
            [vw_Good].[Level1],
            [vw_Good].[Level2],
            CASE @intLanguage
              WHEN 0 THEN [Name]
              ELSE [LatinName]
            END AS [Name],
            CASE @intLanguage
              WHEN 0 THEN [NamePrn]
              ELSE [LatinNamePrn]
            END AS [NamePrn],
            CASE @intLanguage
              WHEN 0 THEN [Level1Description]
              ELSE [Level1LatinDescription]
            END AS [Level1Name],
            CASE @intLanguage
              WHEN 0 THEN [Level2Description]
              ELSE [Level2LatinDescription]
            END AS [Level2Name]	,	 
            [OrderPoint],
            [MinValue],
            [MaxValue],
            [ProductCompany],
            [UnitDescription],
            [TypeDescription],
            @SystemDay + ' ' + @SystemDate  AS Sysdate,
            tInventory_Good.[InventoryNo],
            tInventory.Description AS Inventoryname ,
            Mojodi
			, ISNULL((select feeunit from tfacd 
				inner join (SELECT MAX(tfacd.intSerialNo) AS maxint,goodcode
							FROM tfacd  INNER JOIN  [tFacM] ON [tFacD].[intSerialNo] = [tFacM].[intSerialNo]    
							where tfacm.status=1 and tfacm.accountyear=@AccountYear1
					        GROUP BY goodcode )k ON tfacd.goodcode=k.goodcode and tfacd.intserialno=k.maxint
				WHERE [tfacd].[GoodCode]=[vw_Good].[Code]),[vw_Good].BuyPrice) AS feeunit
	

    FROM    [dbo].[vw_Good]
            INNER JOIN tInventory_Level1 ON tInventory_Level1.Level1 = vw_Good.Level1
            INNER JOIN tInventory ON tInventory.Branch = tInventory_Level1.Branch
                                     AND tInventory.InventoryNo = tInventory_Level1.InventoryNo
            INNER JOIN tInventory_Good ON vw_Good.Code = tInventory_Good.GoodCode

    WHERE   [vw_Good].[Level1] >= @level11
            AND [vw_Good].[Level1] <= @level12
            AND [vw_Good].[Level2] >= @level21
            AND [Level2] <= @level22
            AND [GoodType] <> 2
            AND [GoodType] <> 4
            AND tInventory_Good.[InventoryNo] = @Inventory1
            AND tInventory.Branch = tInventory_Good.Branch
            AND tInventory.InventoryNo = tInventory_Good.InventoryNo
            AND tInventory_Good.AccountYear = @AccountYear1
  		    AND [tInventory_Good].mojodi <= OrderPoint 
  		    AND OrderPoint > 0

ORDER BY Name

GO






UPDATE dbo.tObjects
SET ObjectName = N'گزارش فروش ساعتي درصدي'
WHERE ObjectId = 'RepPercentInvoicePerHour'

GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO



ALTER  PROCEDURE [dbo].[AverageCalculateBuyPrice] 
(@GoodCode as int ,  @DateAfter Nvarchar(10) , @DateBefore Nvarchar(10) , @Flag INT) 
 AS
DECLARE @AccountYear SMALLINT
SET @AccountYear = CAST('13' + SUBSTRING(@DateAfter ,1,2) AS SMALLINT)
--PRINT @AccountYear
DECLARE @Branch INT 
SET @Branch = dbo.Get_Current_Branch()
--PRINT @Branch
Declare @FeeUnit INT

IF @Flag = 0 
BEGIN

DECLARE @BuyTotal FLOAT
DECLARE @BuyPriceTotal BIGINT
DECLARE @BuyReturnTotal FLOAT
DECLARE @BuyReturnPriceTotal BIGINT
DECLARE @FirstMojodiTotal FLOAT 
 
	SELECT @BuyTotal = CASE WHEN M1.Status = 1 THEN  ISNULL(SUM(D1.Amount) ,0) ELSE 0 END ,
			@BuyPriceTotal = CASE WHEN M1.Status = 1 THEN  ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) ELSE 0 END,
			@BuyReturnTotal = CASE WHEN M1.Status = 4 THEN  ISNULL(SUM(D1.Amount) ,0) ELSE 0 END,
			@BuyReturnPriceTotal = CASE WHEN M1.Status = 4 THEN  ISNULL(SUM(D1.FeeUnit * D1.Amount ) ,0) ELSE 0 END
				 FROM [dbo].[tFacM] M1
					INNER JOIN [dbo].[tFacD] D1 ON [M1].[Branch] = [D1].[Branch]
							AND [M1].[intSerialNo] = [D1].[intSerialNo]
					WHERE  
						    M1.[Date] >= @DateAfter
						AND M1.[Date] <= @DateBefore
						AND (M1.Status = 1 OR M1.Status = 4)
						AND D1.GoodCode = @GoodCode
						AND M1.AccountYear = @AccountYear
						AND M1.Branch = @Branch
						AND Recursive = 0
						GROUP BY Status , GoodCode
	SELECT @FirstMojodiTotal = ISNULL(SUM(tInventory_Good.FirstMojodi) ,0) 
				 FROM dbo.tInventory_Good
					WHERE  
						 tInventory_Good.GoodCode = @GoodCode
						AND tInventory_Good.AccountYear = @AccountYear
						AND tInventory_Good.Branch = @Branch



--PRINT @BuyTotal 
--PRINT @BuyPriceTotal 
--PRINT @BuyReturnTotal 
--PRINT @BuyReturnPriceTotal 
--PRINT @FirstMojodiTotal  


DECLARE @BuyPrice INT 
SET @BuyPrice = (Select IsNull(BuyPrice,1) From tGood 	Where  tGood.Code = @Goodcode )
SET @FeeUnit = CASE WHEN (ISNULL(@FirstMojodiTotal,0) + ISNULL(@BuyTotal,0) - ISNULL(@BuyReturnTotal ,0) ) = 0 THEN @BuyPrice  
			ELSE 
			CAST(
			((ISNULL(@FirstMojodiTotal,0) * @BuyPrice) + ISNULL(@BuyPriceTotal,0) - ISNULL(@BuyReturnPriceTotal ,0))  
			/ (ISNULL(@FirstMojodiTotal,0) + ISNULL(@BuyTotal,0) - ISNULL(@BuyReturnTotal ,0) ) 
			AS BIGINT ) END

		--UPDATE dbo.tInventory_Good
		--SET BuyPriceAverage = @FeeUnit
		--WHERE AccountYear = @AccountYear
		--AND Branch = @Branch AND GoodCode = @GoodCode

	--SET @FeeUnit = (Select (IsNull(Sum(FeeUnit * Amount) ,0)/ISNULL(Sum(Amount),1)) From tFacM inner join tfacd On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch
	--			Where tfacm.Status = 1 and Recursive = 0 And Date >= @DateAfter And Date <= @DateBefore And tfacd.GoodCode = @Goodcode ) 
	--IF @FeeUnit = 0 
	--	SET @FeeUnit = (Select IsNull(BuyPrice,1) From tGood 	Where  tGood.Code = @Goodcode ) 
END
ELSE
	SET @FeeUnit = (Select IsNull(BuyPrice,1) From tGood 	Where  tGood.Code = @Goodcode ) 


Select @FeeUnit As AverageBuyPrice


GO







SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--درست كردن فاكتور خريد رستوراني 
--Version  V26_15_Fix5   &  V26_15
--Reports name  :  A5\BuyFactor_A5   & A4\BuyFactor_A4
--دقت شود ريپورت جديد برداشته شود
--92/04/23


ALTER    VIEW dbo.VwBuy_new
AS
SELECT DISTINCT 
                      dbo.tFacM.Branch , dbo.tFacM.[No], dbo.tFacM.[Date], dbo.tFacM.SumPrice, dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.Recursive, dbo.tFacM.StationId, 
                      dbo.tFacM.ServePlace AS masterserveplace, dbo.tFacD.ServePlace, dbo.tFacM.OrderType, dbo.tFacM.Status, dbo.tFacD.GoodCode, dbo.tGood.Weight, 
                      dbo.tFacD.FeeUnit, dbo.tFacM.ShiftNo, dbo.tFacD.Amount * dbo.tFacD.FeeUnit AS FeeTotal, dbo.tFacM.intSerialNo, dbo.tFacM.FacPayment, 
                      dbo.tFacM.InCharge, dbo.tServePlace.[Description] AS FactorServePlace, 
                      dbo.tServePlace.LatinDescription AS FactorLatinServePlace, dbo.tFacM.Customer, dbo.tFacM.Owner, dbo.tFacM.RegDate, dbo.tGood.NamePrn, 
                      dbo.tGood.LatinNamePrn, ISNULL(dbo.tSupplier.Address, '' ) AS Address,

-- 	        CASE ISNULL(dbo.tSupplier.Unit,0) WHEN 0 THEN ''
-- 					      ELSE  	  N' :  واحد '+CAST(dbo.tSupplier.Unit AS NVARCHAR(50)) 
-- 	        END AS Unit,
-- 
-- 	        CASE ISNULL(dbo.tSupplier.InternalNo,0) WHEN 0 THEN ''
-- 						   ELSE  N' : داخلي  '+ CAST(dbo.tSupplier.InternalNo AS NVARCHAR(50)) 
-- 	         END AS InternalNo,
-- 
-- 	        CASE ISNULL(dbo.tSupplier.Flour,0) WHEN 0 THEN ''
-- 					       ELSE N' : طبقه '+CAST(dbo.tSupplier.Flour AS NVARCHAR(50)) 
-- 	         END AS Flour,

	       -- CASE ISNULL((dbo.tPer.nvcFirstName + dbo.tPer.nvcSurName), N'0') 
                --      WHEN N'0' THEN N' ' ELSE dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName END AS GarsonName, 
                 --     CASE dbo.tPer.Gender WHEN 0 THEN N'خانم' WHEN 1 THEN N'آقاي' END AS GarsonGender, 
                  --    CASE ISNULL(LTRIM(RTRIM(dbo.tFacD.DifferencesDescription)), '') 
                  --    WHEN '' THEN '-' ELSE dbo.tFacD.DifferencesDescription END AS DifferencesDescription,  dbo.tFacM.TableNo As TableCode,dbo.tTable.[Name] As TableDesc,
                     CASE CAST(dbo.tFacM.BascoleNo AS NVARCHAR(50)) 
                      WHEN NULL THEN '' WHEN '0' THEN '' ELSE CAST(dbo.tFacM.BascoleNo AS NVARCHAR(50)) + SPACE(3) + N'   :ترازو' END AS BascoleNo, 
                      CASE dbo.tSupplier.Tel1 WHEN NULL THEN '' WHEN '' THEN '' ELSE dbo.tSupplier.Tel1 + SPACE(5) + N'   : تلفن' END AS Tel1, CASE dbo.tSupplier.Tel2 WHEN NULL
                       THEN '' WHEN '' THEN '' ELSE dbo.tSupplier.Tel2 + SPACE(5) + N'   : تلفن ' END AS Tel2,
 	          CASE dbo.tSupplier.[Name] + SPACE(3) + dbo.tSupplier.family WHEN NULL      THEN dbo.tSupplier.WorkName WHEN '' THEN dbo.tSupplier.WorkName 
                       ELSE    CASE dbo.tSupplier.Sex WHEN 0 THEN N' خانم  ' + dbo.tSupplier.[Name] + '  ' + dbo.tSupplier.family 
		          ELSE  N' آقاي '  + dbo.tSupplier.[Name] + '  ' + dbo.tSupplier.family
                                     End 
                       END AS family, 
                      dbo.tFacM.DiscountTotal , dbo.tFacM.CarryFeeTotal, 
                      dbo.tFacM.ServiceTotal,  dbo.tFacM.PackingTotal, 
                      CASE dbo.tFacD.Amount WHEN 0.0 THEN NULL ELSE CAST(dbo.tFacD.Amount AS DECIMAL(10, 3)) END AS WeightTotal, 
                      CASE dbo.tFacD.Amount WHEN 0.0 THEN NULL ELSE CAST(dbo.tFacD.Amount AS DECIMAL(10, 3)) END AS Amount, dbo.tSupplier.membershipid , tFacm.Balance ,
	        dbo.tGood.Unit As UnitType ,dbo.tFacM.AccountYear , dbo.tUnitGood.[Description] As UnitDesc , dbo.tStatusType.NvcDescription AS StatusName
	       -- ,CASE dbo.tSupplier.Sex WHEN 0 THEN N'خانم' WHEN 1 THEN N'آقاي' END AS CustGender
	       , tfacm.TaxTotal , tfacm.DutyTotal
		,(SELECT [Description] FROM [tInventory] WHERE [tInventory].[InventoryNo]=tfacd.[DestInventoryNo]) AS DestInventoryName
		,(SELECT [Description] FROM [tInventory] WHERE [tInventory].[InventoryNo]=tfacd.intInventoryNo) AS InttInventoryName
 		,([tPer].[nvcFirstName]+' '+[tPer].[nvcSurName]) AS FullName
	       , ISNULL(tblPub_Destination.NvcDestination , '') AS DestinationName
FROM           dbo.tFacM INNER JOIN
                      dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND dbo.tFacM.Branch = dbo.tFacD.Branch  INNER JOIN
                      dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code INNER JOIN
                      dbo.tUnitGood ON dbo.tGood.Unit = dbo.tUnitGood.Code INNER JOIN
                      dbo.tServePlace ON dbo.tFacM.ServePlace = dbo.tServePlace.intServePlace INNER JOIN
					dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo INNER JOIN 
                      dbo.tPrinting ON dbo.tServePlace.intServePlace = dbo.tPrinting.ServePlace LEFT OUTER JOIN
                      tSupplier ON tfacM.Owner = tSupplier.code and  tfacM.Branch = tSupplier.Branch  --LEFT OUTER JOIN
--                      tPer ON dbo.tFacM.Incharge = dbo.tPer.PPNO and dbo.tFacM.Branch = dbo.tPer.Branch   LEFT OUTER JOIN
--                      tTable ON dbo.tFacM.TableNo = dbo.tTable.[No] and  dbo.tFacM.Branch = dbo.tTable.[Branch]
					INNER JOIN tuser ON [tUser].[UID]=[tFacM].[User] 
					INNER JOIN tper ON [tPer].[pPno]=[tUser].[pPno]
					LEFT outer JOIN dbo.tblPub_Destination ON dbo.tblPub_Destination.DestinationId = dbo.tFacM.DestinationId
Where 		dbo.tFacM.Branch = dbo.Get_Current_Branch()





GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   VIEW VwBuy_Multipart
AS
SELECT DISTINCT 
                	dbo.VwBuy_new.*, tprinting.PrinterNo, 	
		CASE dbo.tPrinting.Arm 
			WHEN 1 THEN dbo.tprinters.Arm 
			ELSE NULL 
		END AS Arm, 
                      	CASE 
			WHEN dbo.tPrinting.Linefeed >= 1 THEN dbo.Repeater(dbo.tprinters.LineFeed, dbo.tPrinting.Linefeed) 
			ELSE NULL 
		END AS Linefeed, 
                      	CASE dbo.tPrinting.Cutter 
			WHEN 1 THEN dbo.tprinters.Cut 
			ELSE NULL 
		END AS Cutter, 
		tprinting.barcode, tprinting.DirectRpt,dbo.tServePlace.[Description] AS ItemServePlace, t2.[Description], tprinting.printformat, noticedescription AS NoticeDescription1, 
                      	NoticeLatinDescription, t2.LatinDescription, dbo.tServePlace.LatinDescription AS ItemLatinServePlace
FROM         dbo.VwBuy_new 
		INNER JOIN
                      	tPrinting ON (dbo.VwBuy_new.Status = dbo.tPrinting.Status AND dbo.VwBuy_new.ServePlace = tprinting.ServePlace AND dbo.VwBuy_new.StationId = tprinting.StationId) 
		INNER JOIN
                      	tprinters ON tprinting.PrinterNo = tprinters.printerNo 
		INNER JOIN
                      	tServePlace ON dbo.VwBuy_new.ServePlace = tServePlace.intServePlace 
		INNER JOIN
                      	tServePlace t2 ON VwBuy_new.masterServePlace = t2.intServePlace 
		INNER JOIN
                      	tprintformat ON tprinting.printformat = tprintformat.printformat 
		LEFT OUTER JOIN
                      	tnoticedescription ON tprintformat.noticeno = tnoticedescription.noticeno
Where 		tPrinting.Branch = dbo.Get_Current_Branch() And 	tprinters.Branch = dbo.Get_Current_Branch()





GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE dbo.Get_BuyInfo(

	@intLanguage	INT = 0,
	@intFacNo 	INT,
	@PrintFormat 	INT,
	@StationId 	INT,
	@Status		INT,
	@intPrinterNo   INT,
	@Mode	 	INT = 2,
	@AccountYear	Smallint = NULL ,
	@PartitionId INT = NULL ,
	@Branch INT = NULL
)
AS
IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()

If @AccountYear Is Null 
	Set @AccountYear = dbo.get_AccountYear() 

	    	SELECT Distinct dbo.[VwBuy_Multipart].[No], dbo.[VwBuy_Multipart].[Date], 
			dbo.[VwBuy_Multipart].[SumPrice], dbo.[VwBuy_Multipart].[Time], 
			dbo.[VwBuy_Multipart].[User], dbo.[VwBuy_Multipart].[Recursive], 
			dbo.[VwBuy_Multipart].[StationId], dbo.[VwBuy_Multipart].[masterserveplace], 
			dbo.[VwBuy_Multipart].[ServePlace], dbo.[VwBuy_Multipart].[OrderType], 
			dbo.[VwBuy_Multipart].[Status], dbo.[VwBuy_Multipart].[GoodCode], 
			dbo.[VwBuy_Multipart].[Weight], dbo.[VwBuy_Multipart].[FeeUnit], 
			dbo.[VwBuy_Multipart].[ShiftNo], dbo.[VwBuy_Multipart].[FeeTotal], 
			dbo.[VwBuy_Multipart].[intSerialNo], dbo.[VwBuy_Multipart].[FacPayment], 
			--dbo.[VwBuy_Multipart].[InCharge], 

			dbo.[VwBuy_Multipart].[Customer], dbo.[VwBuy_Multipart].[Owner], 
			dbo.[VwBuy_Multipart].[RegDate], 

			--dbo.[VwBuy_Multipart].[GarsonName], [VwBuy_Multipart].[GarsonGender], 
			--dbo.[VwBuy_Multipart].[DifferencesDescription], [VwBuy_Multipart].[TableDesc], 
			dbo.[VwBuy_Multipart].[BascoleNo], [VwBuy_Multipart].[Tel1], 
			dbo.[VwBuy_Multipart].[Tel2], [VwBuy_Multipart].[family], 
			dbo.[VwBuy_Multipart].[DiscountTotal], [VwBuy_Multipart].[CarryFeeTotal], 
			dbo.[VwBuy_Multipart].[ServiceTotal], [VwBuy_Multipart].[PackingTotal], 
			dbo.[VwBuy_Multipart].[WeightTotal], [VwBuy_Multipart].[Amount], 
			dbo.[VwBuy_Multipart].[membershipid], [VwBuy_Multipart].[PrinterNo], 
			dbo.[VwBuy_Multipart].[Arm], [VwBuy_Multipart].[Linefeed], 
			dbo.[VwBuy_Multipart].[Cutter], 

			dbo.[VwBuy_Multipart].[printformat],
       		             dbo.[VwBuy_Multipart].[DirectRpt] , dbo.[VwBuy_Multipart].[Balance] ,
			VwBuy_Multipart.Address /*+ ' '+ VwBuy_Multipart.Flour + ' ' + VwBuy_Multipart.Unit 
			+ ' '+VwBuy_Multipart.InternalNo*/ AS CustomerAddress,



			CASE @intLanguage 
				WHEN 0 THEN
					CASE @Mode 
						WHEN 1 THEN N'چاپ مجدد'
						WHEN 4 THEN N'اصلاحي'
						ELSE ''
					END
				WHEN 1 THEN 
					CASE @Mode 
						WHEN 1 THEN N'Repeated Print'
						WHEN 4 THEN N'Edited'
                           			ELSE ''
					END
			END AS ReportHeder,

			CASE @intLanguage 
				WHEN 0 THEN
					CASE dbo.VwBuy_Multipart.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'مرجوعي'
					END
				WHEN 1 THEN 
					CASE dbo.VwBuy_Multipart.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'Reffered'
					END
			END AS RecursievAlert,

		    	CASE @intLanguage 	
				WHEN 0 THEN dbo.VwBuy_Multipart.NamePrn
				WHEN 1 THEN dbo.VwBuy_Multipart.LatinNamePrn
			END AS GoodName,


			CASE @intLanguage 	
				WHEN 0 THEN VwBuy_Multipart.ItemServePlace
				WHEN 1 THEN VwBuy_Multipart.ItemLatinServePlace
			END AS ItemServePlaceDesc,

			CASE @intLanguage 	
				WHEN 0 THEN VwBuy_Multipart.NoticeDescription1
				WHEN 1 THEN VwBuy_Multipart.NoticeLatinDescription
			END AS NoticeDescription,

			CASE @intLanguage 	
				WHEN 0 THEN VwBuy_Multipart.FactorServePlace
				WHEN 1 THEN VwBuy_Multipart.FactorLatinServePlace
			END AS FactorServeDescription,

	      	          CASE VwBuy_Multipart.barcode 
				WHEN 1 THEN  (SELECT TOP 1  dbo.BarcodeGenerator(dbo.VwBuy_Multipart.ServePlace,@intFacNo) where [No]= @intFacNo and Status = 2 )
				ELSE '' END AS Barcode ,

			dbo.[VwBuy_Multipart].[UnitType] , dbo.[VwBuy_Multipart].[UnitDesc] ,dbo.[VwBuy_Multipart].StatusName

			, VwBuy_Multipart.Taxtotal , VwBuy_Multipart.DutyTotal
			,[DestInventoryName],[InttInventoryName],[dbo].[VwBuy_Multipart].FullName
			, VwBuy_Multipart.DestinationName 

		FROM dbo.VwBuy_Multipart

		WHERE 	No=@intFacNo 	
			AND (PrintFormat  = @PrintFormat OR @PrintFormat =0 )
			--AND GoodCode NOT IN (SELECT GoodCode  FROM tPrinterGood WHERE intPrinterFormat = @PrintFormat )
			AND ( dbo.VwBuy_Multipart.StationId = @StationId OR @Mode  =  0 )
			AND dbo.VwBuy_Multipart.status =@Status 
			And VwBuy_Multipart.AccountYear = @AccountYear
			AND   VwBuy_Multipart.PrinterNo=@intPrinterNo
			AND   VwBuy_Multipart.Branch = @Branch


GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   Procedure dbo.Insert_Supplier  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@State int ,  
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,
	@TotalRemainingAmount INT ,   
	@Code Bigint out  
)  

as  

Begin Tran  

DECLARE @Branch INT 
 Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode ) --AND (Branch = @Branch ) )  
    end   
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId) <> 0 -- AND Branch = @Branch ) <> 0   
  Goto ErrHandler   

--Set @Code = (Select  isnull(Max(Code),0) + 1 from tSupplier where code > 0)  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

insert Into dbo.tSupplier  
(   
	--Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	State ,  
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	Discount,   
	[Description],   
	[Date],   
	[Time],   
	[User], 
	TotalRemainingAmount , 
	Branch  
)  
values  
(   
	--@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@State,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@Discount,   
	@Description,   
	@Date,   
	@Time,   
	@User , 
	@TotalRemainingAmount , 
	@Branch  
)  
if @@Error <> 0   
 goto ErrHandler  

Set @Code = @@Identity  
 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] )   

update tSupplier set address=dbo.addressedit(address) where code=@code  -- AND Branch = @Branch
   

Commit Tran
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code  
--Select @Code


GO


