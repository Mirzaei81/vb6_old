

--Script_V26_16_Fix8_BuyFactorPrint.SQL
--جلوگیری از تکرار رکورد هنگام چاپ در حالت موارد تکراری
--برای فاکتورخرید و حواله
--93/08/05

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Get_BuyInfo(

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

