


--Script_V26_16_Fix12_VwBuy.sql
--درست كردن فاكتور خريد , رسید و حواله زمانیکه از ایستگاه دیگر پرینت گرفته می شود  
--برای همه ورژن های رستورانی
--92/04/03


ALTER     VIEW dbo.VwBuy_new
AS
SELECT DISTINCT 
                      dbo.tFacM.Branch , dbo.tFacM.[No], dbo.tFacM.[Date], dbo.tFacM.SumPrice, dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.Recursive,  
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



ALTER    VIEW VwBuy_Multipart
AS
SELECT DISTINCT 
                	dbo.VwBuy_new.*, tprinting.StationId, tprinting.PrinterNo, 	
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
                      	tPrinting ON (dbo.VwBuy_new.Status = dbo.tPrinting.Status AND dbo.VwBuy_new.ServePlace = tprinting.ServePlace ) 
		INNER JOIN
                      	tprinters ON tprinting.PrinterNo = tprinters.printerNo 
		INNER JOIN	    dbo.tPartitions ON dbo.tPrinting.PartitionId = dbo.tPartitions.PartitionID
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




