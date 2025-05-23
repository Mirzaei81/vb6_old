

--Script_V26_16_Fix12_MarkedGoods.sql
--  فانکشن کالاهای اصلی را از فاکتور دریافت می کند
--میتوان چک کرد اگر در پرینتری فقط کالای مارک دار داشته باشد فقط از آن پرینتر چاپ شود
--برای پیتزا خاتون و سایر مشتریان هم قابل استفاده است
-- در پروسیچر [Get_InvoiceInfo_IsFilter] 
--باید مقدار متغیر @IsMarkedGoods = 1 باشد

--94/03/08

IF COL_LENGTH('[tGood]','BitMarkedGood') IS NULL
	ALTER TABLE dbo.tGood
	ADD BitMarkedGood BIT NOT NULL  DEFAULT(0)

GO

--UPDATE tgood SET BitMarkedGood = 0
--GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  VIEW [dbo].[VwInvoice_Multipart]   
AS
SELECT DISTINCT 
                      dbo.[tGoodLevel1].code AS levelcode1,dbo.tGoodLevel1.Description as leveldesc1,dbo.[tGoodLevel2].code AS levelcode2,dbo.tGoodLevel2.Description as leveldesc2,
                      dbo.tGoodLevel2.LatinDescription as LatinLeveldesc2,dbo.tUnitGood.Description AS UnitDesc ,dbo.tFacD.intRow,dbo.tFacM.[No], dbo.tFacM.[Date], dbo.tFacM.SumPrice, dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.Recursive, dbo.tPrinting.StationId, 
                      dbo.tFacM.ServePlace AS masterserveplace, dbo.tFacD.ServePlace, dbo.tFacM.OrderType, dbo.tFacM.Status, dbo.tFacD.GoodCode, dbo.tGood.Weight, 
                      dbo.tFacD.FeeUnit, dbo.tFacM.ShiftNo, dbo.tFacD.Amount * dbo.tFacD.FeeUnit AS FeeTotal, dbo.tFacM.intSerialNo, dbo.tFacM.FacPayment, 
                      dbo.tFacM.InCharge, dbo.tServePlace.[Description] AS FactorServePlace, 
                      dbo.tServePlace.LatinDescription AS FactorLatinServePlace, dbo.tFacM.Customer, dbo.tFacM.Owner, dbo.tFacM.RegDate, dbo.tGood.NamePrn, 
                      dbo.tGood.LatinNamePrn, dbo.tCust.Address,dbo.tCust.Mastercode,tPer_1.nvcSurName AS UserName,

	        CASE ISNULL(dbo.tCust.Unit,'') WHEN '' THEN ''
					      ELSE  	  N' :  واحد '+CAST(dbo.tCust.Unit AS NVARCHAR(50)) 
	        END AS Unit,
		
	        CASE ISNULL(dbo.tCust.InternalNo,'') WHEN '' THEN ''
						   ELSE  N' : داخلي  '+ CAST(dbo.tCust.InternalNo AS NVARCHAR(50)) 
	         END AS InternalNo,
	
	        CASE ISNULL(dbo.tCust.Flour,'') WHEN '' THEN ''
					       ELSE N' : طبقه '+CAST(dbo.tCust.Flour AS NVARCHAR(50)) 
	         END AS Flour,
	        
	        CASE ISNULL((dbo.tPer.nvcFirstName + dbo.tPer.nvcSurName), N'0') 
                      WHEN N'0' THEN N' ' ELSE dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName END AS GarsonName, 
                      CASE dbo.tPer.Gender WHEN 0 THEN N'خانم' WHEN 1 THEN N'آقاي' END AS GarsonGender, 
                      CASE ISNULL(LTRIM(RTRIM(dbo.tFacD.DifferencesDescription)), '') 
                      WHEN '' THEN '-' ELSE dbo.tFacD.DifferencesDescription END AS DifferencesDescription,  dbo.tFacM.TableNo As TableCode,dbo.tTable.[Name] As TableDesc,
                     CASE CAST(dbo.tFacM.BascoleNo AS NVARCHAR(50)) 
                      WHEN NULL THEN '' WHEN '0' THEN '' ELSE CAST(dbo.tFacM.BascoleNo AS NVARCHAR(50)) + SPACE(3) + N'   :ترازو' END AS BascoleNo, 
                      CASE dbo.tCust.Tel1 WHEN NULL THEN '' WHEN '' THEN '' ELSE dbo.tCust.Tel1 + SPACE(5) + N'   : تلفن' END AS Tel1, CASE dbo.tCust.Tel2 WHEN NULL
                       THEN '' WHEN '' THEN '' ELSE dbo.tCust.Tel2 + SPACE(5) + N'   : تلفن ' END AS Tel2,
 			CASE 
			WHEN (dbo.tCust.MasterCode is null And dbo.tCust.WorkName ='' and dbo.tCust.Name <> '')
				then   Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name ELSE N' آقاي '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name END
			WHEN (dbo.tCust.MasterCode is null And dbo.tCust.WorkName ='' and dbo.tCust.Name = '')
				then  Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' '  ELSE N' آقاي '  +  dbo.tcust.Family + ' '  END
			When (dbo.tCust.MasterCode is null And dbo.tCust.WorkName <>'')
				then dbo.tCust.WorkName
			WHEN (dbo.tCust.MasterCode is not null And tCust_1.WorkName <>'' and dbo.tCust.Name <> '')
				then tCust_1.WorkName + '_' +  Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name ELSE N' آقاي '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name END
			WHEN (dbo.tCust.MasterCode is not null And tCust_1.WorkName <>'' and dbo.tCust.Name = '')
				then tCust_1.WorkName + '_' +  Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' '  ELSE N' آقاي '  +  dbo.tcust.Family + ' '  END

			End as family	 ,        -- CASE dbo.tCust.[Name] + SPACE(3) + dbo.tCust.family WHEN NULL      THEN dbo.tCust.WorkName WHEN '' THEN dbo.tCust.WorkName 
                      -- ELSE    CASE dbo.tCust.Sex WHEN 0 THEN N' خانم  ' + dbo.tCust.[Name] + '  ' + dbo.tCust.family 
		        --  ELSE  N' آقاي '  + dbo.tCust.[Name] + '  ' + dbo.tCust.family
                 --   End 
  -- END AS family, 
                      dbo.tFacM.DiscountTotal , dbo.tFacM.CarryFeeTotal, 
                      dbo.tFacM.ServiceTotal,  dbo.tFacM.PackingTotal, 
                      CASE dbo.tFacD.Amount WHEN 0.0 THEN NULL ELSE CAST(dbo.tFacD.Amount AS DECIMAL(10, 3)) END AS WeightTotal, 
                      CASE dbo.tFacD.Amount WHEN 0.0 THEN NULL ELSE CAST(dbo.tFacD.Amount AS DECIMAL(10, 3)) END AS Amount, dbo.tcust.membershipid , tFacm.Balance ,
	                  dbo.tGood.Unit As UnitType , dbo.tfacd.Rate , dbo.tfacd.ChairName , dbo.tGood.MainType ,dbo.tFacM.Branch , dbo.tFacM.AccountYear , tGood.NumberOfUnit
			  , tprinting.PrinterNo ,
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
			tprinting.barcode, tprinting.DirectRpt,dbo.tServePlace.[Description],dbo.tServePlace.[Description] AS ItemServePlace,  tprinting.printformat, noticedescription AS NoticeDescription1, 
                         NoticeLatinDescription, dbo.tServePlace.LatinDescription, dbo.tServePlace.LatinDescription AS ItemLatinServePlace 
				,dbo.tPrinting.PermittedModes , dbo.tfacm.NvcDescription , TaxBuy , TaxSale , DutyBuy , DutySale , dbo.tFacM.TaxTotal , dbo.tFacM.DutyTotal
				, dbo.tPrinting.PartitionId , dbo.tPartitions.PartitionDescription , tfacM.GuestNo , tfacm.TempNo
				, ISNULL(dbo.tblTotal_Order.Date, N'') AS OrderDate, ISNULL(dbo.tblTotal_Order.Time, N'') AS OrderTime
				, GoodNamePrn2 , GoodNamePrn3 , BitMarkedGood
FROM        dbo.tFacM 
			INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND dbo.tFacM.Branch = dbo.tFacD.Branch  
			INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code 
			INNER join  dbo.tUnitGood on dbo.tgood.unit=dbo.tunitgood.code
			inner join dbo.tGoodLevel2 on dbo.tGoodLevel2.code=dbo.tGood.level2
			inner join dbo.tGoodLevel1 on dbo.tGoodLevel1.Code=dbo.tGoodLevel2.Level1Code
			inner JOIN dbo.tServePlace ON dbo.tFacM.ServePlace = dbo.tServePlace.intServePlace 
			INNER JOIN  tPrinting ON (dbo.tfacm.ServePlace = tprinting.ServePlace AND dbo.tfacm.Status = tprinting.Status AND tfacm.Branch = dbo.tPrinting.Branch) 
			INNER JOIN dbo.tPartitions ON dbo.tPrinting.PartitionId = dbo.tPartitions.PartitionID
			INNER JOIN	tprinters ON tprinting.PrinterNo = tprinters.printerNo AND dbo.tPrinting.Branch = dbo.tPrinters.Branch 
			INNER JOIN	tprintformat ON tprinting.printformat = tprintformat.printformat 
			LEFT OUTER JOIN	tnoticedescription ON tprintformat.noticeno = tnoticedescription.noticeno
			left outer JOIN tCUst ON tfacM.Customer = tcust.code  
			LEFT OUTER JOIN tPer ON dbo.tFacM.Incharge = dbo.tPer.PPNO    
			LEFT OUTER JOIN tTable ON dbo.tFacM.TableNo = dbo.tTable.[No] and  dbo.tFacM.Branch = dbo.tTable.[Branch]
			LEFT OUTER JOIN  dbo.tCust tCust_1  on dbo.tCust.MasterCode = tCust_1.Code --and dbo.tCust.Branch = tCust_1.Branch
			LEFT OUTER JOIN dbo.tUser AS tUser_1 ON dbo.tFacM.[User] = tUser_1.UID  
			inner join dbo.tPer AS tPer_1 ON tUser_1.PPNO = tPer_1.PPNO  
            LEFT OUTER JOIN  dbo.tblTotal_Order ON dbo.tFacM.intSerialNo = dbo.tblTotal_Order.intSerialNo AND dbo.tFacM.Branch = dbo.tblTotal_Order.Branch

--Where 		dbo.tFacM.Branch = dbo.Get_Current_Branch()



GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FN_IF_GetMarkedGood]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FN_IF_GetMarkedGood]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


-- این فانکشن کالاهای مارک دار را از فاکتور دریافت کرده و چک می کند تا اگر کالای مارک دار تنها بود فقط در پرینتر مشخص شده چاپ شود
CREATE  Function [dbo].[FN_IF_GetMarkedGood]

(
	@intFacNo 	INT,
	@AccountYear	Smallint ,
	@Branch INT = NULL
)

RETURNS  int  
	
As

BEGIN

DECLARE @IsContinue INT

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()

	SELECT @IsContinue = COUNT(dbo.[VwInvoice_Multipart].[GoodCode])
		FROM dbo.VWInvoice_MultiPart
		  WHERE     VWInvoice_MultiPart.[No]=@intFacNo 	
			AND VWInvoice_MultiPart.status =2 
			And VWInvoice_MultiPart.AccountYear = @AccountYear
			AND VWInvoice_MultiPart.Branch=@Branch
			AND VWInvoice_MultiPart.BitMarkedGood  = 0
			AND VWInvoice_MultiPart.MainType  = 1

RETURN @IsContinue

END



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER PROCEDURE [dbo].[Get_InvoiceInfo_IsFilter] (

	
	@intFacNo 	INT,
	@PrintFormat 	INT,
	@StationId 	INT,
	@Status		INT,
	@intPrinterNo   INT,
	@AccountYear	Smallint ,
    @Mode	INT   ,
    @PartitionId INT 
)
AS


DECLARE @Branch INT
IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
--==============================================================================
--فانکشن برای اینکه کالاهای غیراصلی را فقط از فیش مشتری و فاکتور فروش چاپ کند

--For Double Filter in All report except invoicefich & invoiceFactor
DECLARE @IsDoubleFilter BIT 
SET @IsDoubleFilter = 1   -- اکتیو کردن فیلتر کالاهای غیر اصلی

Declare  @ExistGood BIT
SET @ExistGood = 1
IF @IsDoubleFilter = 1
	SELECT @ExistGood = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END from dbo.[FN_IF_GetMainGood](@intFacNo,@PrintFormat,@StationId,@Status,@intPrinterNo,@Mode,@AccountYear,@PartitionId,@Branch)

--================================================================================
--فانکشن برای اینکه فقط کالاهای مارک دار را فقط ازپرینتر مشخص شده چاپ کند
DECLARE @IsMarkedGoods INT 
DECLARE @PrinterMarkedGoods INT 
Declare  @MarkedGoods INT 

SET @IsMarkedGoods  = 1   --  برای اکتیو کردن چاپ کالاهای مارک دار این متغیر را یک کنید
SET @PrinterMarkedGoods = 5   -- شماره پرینتر را مشخص کنید

IF @intPrinterNo <> @PrinterMarkedGoods
	SET @MarkedGoods = 0
ELSE
	SET @MarkedGoods = 1

IF @IsMarkedGoods = 1 AND @intPrinterNo <> @PrinterMarkedGoods
	SELECT @MarkedGoods = dbo.[FN_IF_GetMarkedGood](@intFacNo,@AccountYear,@Branch)   --If >= 1 then other printer does print 

--=======================================================
If @Mode <> 8 
BEGIN
SELECT  distinct   dbo.tFacM.intSerialNo,dbo.tFacM.ServePlace ,dbo.tFacd.DifferencesDescription , 
		   dbo.tFacD.GoodCode , tPrinting.Directrpt ,dbo.tFacd.Amount 
	FROM    dbo.tFacM     	 
	INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
	INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacD.GoodCode
    INNER JOIN dbo.tPrinting ON dbo.tPrinting.Status = tfacm.Status AND tfacm.ServePlace = dbo.tPrinting.ServePlace AND tfacm.Branch = dbo.tPrinting.Branch 
	INNER JOIN dbo.tPrintFormat ON dbo.tPrintFormat.PrintFormat = dbo.tPrinting.PrintFormat
	WHERE  
	dbo.tFacd.GoodCode not in (select dbo.tPrinterGood.goodcode from dbo.tPrinterGood WHERE intPrinterFormat = @PrintFormat)
	and dbo.tFacM.status = @status 
	And dbo.tPrinting.StationId = @StationId  
	And dbo.tPrinting.PartitionId = @PartitionId
	and dbo.tPrintFormat.PrintFormat = @PrintFormat
	and dbo.tFacM.[No] = @intFacNo 
	and dbo.tFacM.AccountYear = @AccountYear 
	and dbo.tPrinting.PrinterNo = @intPrinterNo
	AND dbo.tPrinting.PermittedModes & @Mode = @Mode
	AND (@IsDoubleFilter = 0 OR @PrintFormat = 1 OR @PrintFormat = 9 OR @ExistGood = 1)
	AND (@IsMarkedGoods = 0 OR @PrintFormat = 1 OR @PrintFormat = 9 OR @MarkedGoods >= 1)
END

ELSE

BEGIN
select ISNULL(T2.intSerialNo,T1.intSerialNo) AS intSerialNo,ISNULL(T2.ServePlace,T1.ServePlace) AS ServePlace,
       ISNULL(T1.DifferencesDescription,T2.DifferencesDescription)AS DifferencesDescription,
       ISNULL(T1.GoodCode,T2.GoodCode) AS GoodCode  ,ISNULL(T1.Directrpt,T2.Directrpt) AS Directrpt ,
       ISNULL(T1.Amount,0)-ISNULL(T2.Amount,0) AS Amount  
from
(SELECT  distinct   dbo.tFacM.intSerialNo,dbo.tFacM.ServePlace ,ISNULL(dbo.tFacd.DifferencesDescription , '') AS DifferencesDescription ,
                    dbo.tFacD.GoodCode , tPrinting.Directrpt ,
		    dbo.tFacd.Amount 
	FROM    dbo.tFacM     	 
	INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
        INNER JOIN dbo.tPrinting ON dbo.tPrinting.Status = tfacm.Status AND tfacm.ServePlace = dbo.tPrinting.ServePlace AND tfacm.Branch = dbo.tPrinting.Branch
	INNER JOIN dbo.tPrintFormat ON dbo.tPrintFormat.PrintFormat = dbo.tPrinting.PrintFormat
	where 
	dbo.tFacd.GoodCode not in (select dbo.tPrinterGood.goodcode from dbo.tPrinterGood WHERE intPrinterFormat = @PrintFormat)
	and dbo.tFacM.status = @status 
	And dbo.tPrinting.StationId = @StationId  
	And dbo.tPrinting.PartitionId = @PartitionId
	and dbo.tPrintFormat.PrintFormat = @PrintFormat
	and dbo.tFacM.[No] = @intFacNo 
	and dbo.tFacM.AccountYear = @AccountYear 
	and dbo.tPrinting.PrinterNo= @intPrinterNo
	AND dbo.tPrinting.PermittedModes & @Mode = @Mode
	)T1


full outer join
	(SELECT  distinct   dbo.tRepFacEditM.intSerialNo,dbo.tRepFacEditM.ServePlace ,ISNULL(dbo.tFacd2.DifferencesDescription,'') AS DifferencesDescription ,
                            dbo.tFacd2.GoodCode , tPrinting.Directrpt , dbo.tFacd2.Amount 
	FROM    dbo.tRepFacEditM     	 
	INNER JOIN dbo.tFacD2 ON dbo.tFacD2.Branch = dbo.tRepFacEditM.Branch AND dbo.tFacD2.Code = dbo.tRepFacEditM.Code
        INNER JOIN dbo.tPrinting ON dbo.tPrinting.Status = tRepFacEditM.Status AND tRepFacEditM.ServePlace = dbo.tPrinting.ServePlace AND tRepFacEditM.Branch = dbo.tPrinting.Branch
	INNER JOIN dbo.tPrintFormat ON dbo.tPrintFormat.PrintFormat = dbo.tPrinting.PrintFormat
	where 
	dbo.tFacd2.GoodCode not in (select dbo.tPrinterGood.goodcode from dbo.tPrinterGood WHERE intPrinterFormat = @PrintFormat)
	and dbo.tRepFacEditM.status = @status 
	And dbo.tPrinting.StationId = @StationId  
	And dbo.tPrinting.PartitionId = @PartitionId
	and dbo.tPrintFormat.PrintFormat = @PrintFormat
	and dbo.tRepFacEditM.[No] = @intFacNo 
	and dbo.tRepFacEditM.AccountYear = @AccountYear 
	and dbo.tPrinting.PrinterNo = @intPrinterNo
        AND dbo.tRepFacEditM.Code=(Select Max(Code) from dbo.tRepFacEditM 
        				where [No]= @intFacNo AND dbo.tRepFacEditM.status = @Status And dbo.tRepFacEditM.AccountYear = @AccountYear)
	AND dbo.tPrinting.PermittedModes & @Mode = @Mode
	)T2
		ON  
		T1.intSerialNo = T2.intSerialNo And 
		T1.GoodCode = T2.GoodCode And 
		T1.ServePlace = T2.ServePlace And
		T1.DifferencesDescription = T2.DifferencesDescription
		Where ( ISNULL(T1.Amount,0)-ISNULL(T2.Amount,0) <> 0  )
		AND (@IsDoubleFilter = 0 OR @PrintFormat = 1 OR @PrintFormat = 9 OR @ExistGood = 1)
		AND (@IsMarkedGoods = 0 OR @PrintFormat = 1 OR @PrintFormat = 9 OR @MarkedGoods >= 1)

END


GO
