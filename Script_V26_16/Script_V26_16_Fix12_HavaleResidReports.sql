
--Script_V26_16_Fix12_HavaleResidReports.sql
--امکان چاپ و پرینت از سندهای خرید و حواله و رسید و رسید موقت در گزارشات
--در قسمت گزارشات کالا و انبار
--گزارش کلی در محدوده تاریخ
--گزارش از دیتایل سند در محدوده تاریخ و با امکان انتخاب شماره سند
--نام ریپورت ها 
--RepAllAssignment.rpt  گزارش کلی سندهای صادره
--RepAssignmentDetail.rpt  گزارش جزئیات سندهای صادره
--94/03/16


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
		  79 ,
          3 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          'SELECT * FROM dbo.tStatusType ' ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM dbo.tblTotal_ItemReports_Details WHERE intReportId = 2 AND Row = 9
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
		  80 ,
          3 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          'SELECT * FROM dbo.tStatusType ' ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM dbo.tblTotal_ItemReports_Details WHERE intReportId = 2 AND Row = 9
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
		  80 ,
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
FROM dbo.tblTotal_ItemReports_Details WHERE intReportId = 26 AND Row = 2
GO


UPDATE dbo.tbltotal_ItemReports SET ReportName = N'گزارش کلی سندهای صادره' WHERE intReportId = 79
UPDATE dbo.tbltotal_ItemReports SET ReportName = N'گزارش جزئیات سندهای صادره' WHERE intReportId = 80

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

-----------------------------------گزارش حواله هاي صادره------------------


ALTER  PROC Get_AllAssignement(		
	@SystemDate NVARCHAR(20) ,
	@SystemDay NVARCHAR(20) ,
	@SystemTime NVARCHAR(20) ,
	@Date1 NVARCHAR(8),
	@Date2 NVARCHAR(8),
	@Inventory1 INT ,
	@Inventory2 INT ,
	@Status1 INT ) 
AS
BEGIN


	SELECT DISTINCT dbo.tFacM.No,
			dbo.tFacM.Date,
			dbo.tFacM.SumPrice,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE  dbo.tFacD.intInventoryNo = dbo.tInventory.InventoryNo ) AS DepartureInventory,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE dbo.tFacD.DestInventoryNo = dbo.tInventory.InventoryNo ) AS DestinationInventory,
			(SELECT nvcFirstName+' '+nvcSurName FROM tper JOIN dbo.tUser ON dbo.tPer.pPno = dbo.tUser.pPno WHERE dbo.tUser.UID=dbo.tFacM.[User]) AS [User],
			dbo.tFacM.NvcDescription
			,@SystemDate AS SystemDate,
			@SystemDay AS SystemDay,
			@SystemTime AS SystemTime , dbo.tStatusType.NvcDescription AS SanadType
	 FROM dbo.tFacM
		INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tStatusType ON dbo.tStatusType.intStatusNo = dbo.tFacM.Status
	 WHERE [Status] = @Status1
			AND dbo.tFacD.intInventoryNo >=@Inventory1
			AND (dbo.tFacD.DestInventoryNo <=@Inventory2  OR dbo.tFacD.DestInventoryNo is NULL )
			AND dbo.tFacm.Date<=@Date2
			AND dbo.tFacM.Date>=@Date1
	ORDER BY dbo.tFacM.No
END


GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


-------------------------------گزارش حواله هاي صادره جزييات-------------
ALTER   PROC	Get_AssignmentDetail(   
	@SystemDate NVARCHAR(20) ,
	@SystemDay NVARCHAR(20) ,
	@SystemTime NVARCHAR(20) ,
	@Date1 NVARCHAR(8),
	@Date2 NVARCHAR(8),
	@Inventory1 INT,
	@Inventory2 INT ,
	@Status1 INT ,
	@intSerialNo1 INT ,
	@intSerialNo2 INT  )
AS	
BEGIN

	SELECT  dbo.tFacM.No, dbo.tFacD.intRow , dbo.tFacM.Date ,
			dbo.tGood.Name AS GoodName,
			SUM(dbo.tFacD.Amount) AS Amount,
			dbo.tFacD.FeeUnit ,
			dbo.tGood.BuyPrice AS BuyPrice,
			dbo.tGood.FinalPrice AS FinalPrice,
			SUM(dbo.tFacD.Amount*tfacd.FeeUnit) AS SumPrice,
			SUM(dbo.tFacD.Amount* dbo.tGood.BuyPrice) AS SumBuyPrice,
			SUM(dbo.tFacD.Amount*dbo.tGood.FinalPrice) AS SumFinalPrice
			,@SystemDate AS SystemDate,
		    @SystemDay AS SystemDay,
		    @SystemTime AS SystemTime ,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE  dbo.tFacD.intInventoryNo = dbo.tInventory.InventoryNo ) AS DepartureInventory,
			(SELECT [Description]
			 FROM  dbo.tInventory WHERE dbo.tFacD.DestInventoryNo = dbo.tInventory.InventoryNo ) AS DestinationInventory
			 , dbo.tStatusType.NvcDescription
	FROM dbo.tFacM 
		   INNER JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch 
			AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
			INNER JOIN dbo.tStatusType ON dbo.tStatusType.intStatusNo = dbo.tFacM.Status
		JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	WHERE dbo.tFacD.intInventoryNo >= @Inventory1
			AND (dbo.tFacD.DestInventoryNo <= @Inventory2 OR tfacd.DestInventoryNo IS NULL )
			AND dbo.tFacM.Date>=@Date1
			AND dbo.tFacM.Date<=@Date2
			AND dbo.tFacM.Status=@Status1
			AND dbo.tFacM.No >= @intSerialNo1
			AND dbo.tFacM.No <= @intSerialNo2
	GROUP BY dbo.tFacM.No,
			 dbo.tFacD.GoodCode,
			 dbo.tGood.Name,
			 dbo.tFacD.FeeUnit,
--			 dbo.tFacD.Amount,
			 dbo.tGood.BuyPrice,
			 dbo.tGood.FinalPrice ,
			 dbo.tFacD.intInventoryNo ,
			 dbo.tFacD.DestInventoryNo ,
			 dbo.tStatusType.NvcDescription ,
			 dbo.tFacD.intRow ,
			 dbo.tfacM.Date
	ORDER BY dbo.tFacM.No

END
--
--
--EXEC Get_AssignmentDetail '','','',N'89/02/06',N'89/02/06',1,101

GO






