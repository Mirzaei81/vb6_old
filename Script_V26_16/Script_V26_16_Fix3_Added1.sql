


--@Date Funtion remarked before for Miladi Date

ALTER  Proc Get_New_FacM_No ( @Status int, @AccountYear smallint, @Branch INT )


as
DECLARE @No INT 
DECLARE @TempNo INT 
DECLARE @ShiftNo INT 
DECLARE @Date NVARCHAR(10)  ---problem with miladi
	SET @ShiftNo= dbo.Get_Shift(GETDATE())     
	SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      
 
	set @No = (Select isnull(max([No]),0)+ 1 as No From tFacM  Where  Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear)
	--set @TempNo = (Select isnull(max([TempNo]),0)+ 1 as No From tFacM  Where  Status = @Status and Branch =  @Branch AND Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())  AND shiftNo = @ShiftNo)
	set @TempNo = (Select isnull(max([TempNo]),0)+ 1 as No From tFacM  Where  Status = @Status and Branch =  @Branch AND Date = @Date  AND shiftNo = @ShiftNo)

SELECT @No AS No , @TempNo AS TempNo


GO






ALTER PROCEDURE [dbo].[Get_InvoiceInfo] (

	@intLanguage INT ,
	@intFacNo 	INT,
	@PrintFormat 	INT,
	@StationId 	INT,
	@Status		INT,
	@intPrinterNo   INT,
	@Mode		INT ,
	@AccountYear	Smallint ,
	@PartitionId INT ,
	@Branch INT = NULL 

)
AS
IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()

Declare @Customer Int
Declare @DateFrom  Nvarchar(20) 
Declare @DateTo NvarChar(20)
Declare @familycount int 
Declare @CurrentBuy int 
Declare @MaxBuy int
Declare @Maxguest int
Declare @Member Bit
Declare @Central Bit
Declare @MainTypeNo int


Declare @intserialNo int
Declare @ChequeAmount int
Declare @Credit Bigint
Declare @CreditBuy Bigint
Declare @RecievedAmount Bigint
Declare @CurrentRecievedAmount Bigint
Declare @NvcDescription Nvarchar(100)
Declare @TempAddress Nvarchar(255)
Declare @PrePayment Bigint
declare @m1 nvarchar(255)
declare @m2 nvarchar(255)
declare @m3 nvarchar(255)
declare @m4 nvarchar(255)
exec Get_FaleHafez @m1 out , @m2 out ,@m3 out , @m4 out

declare @nvcGreatSpeech nvarchar(500)
exec Get_tblGreatSpeech @nvcGreatSpeech out 

Set   @IntSerialNo = (Select intSerialNo From tFacm Where [No]=@intFacNo And Status=@Status  And AccountYear = @AccountYear  And Branch =  @Branch)
SET   @Customer = (SELECT top 1 Customer FROM tFacM WHERE Status=@Status And [No] = @intFacNo And AccountYear = @AccountYear And Branch =  @Branch order by intserialno desc)
Set   @NvcDescription = (Select NvcDescription From tFacm Where [No]=@intFacNo And Status=@Status  And AccountYear = @AccountYear  And Branch =  @Branch)
Set   @TempAddress = (Select TempAddress From tFacm Where [No]=@intFacNo And Status=@Status  And AccountYear = @AccountYear  And Branch =  @Branch)


Set   @CurrentBuy = 0
Set   @Credit = 0
Set   @DateFrom = Left(dbo.Shamsi(GETDATE()),6) + '01' 
Set   @DateTo = dbo.Shamsi(GETDATE()) 
Set   @Member = (Select  Member From tCust Where Code = @Customer )
Set   @Central =  (Select  Central From tCust Where Code = @Customer )
set   @familycount = (Select FamilyNo From tCust Where Code = @Customer )
set   @MaxBuy = dbo.Get_MaxBuy()
set   @Maxguest = dbo.Get_MaxGuest()
Set   @CurrentBuy =  (Select Count(Customer) From tFacm Where Customer = @Customer And Recursive <> 1 And [Date] >= @DateFrom And [Date] <= @DateTo )
Set   @MainTypeNo = (Select Sum(Amount) From tFacd inner join tGood On tfacd.GoodCode = tGood.Code Where tGood.MainType =1 And tFacd.intSerialNo = @intserialNo )

Set   @ChequeAmount = (Select Sum(intChequeAmount) From tFacCheque Where  IntSerialNo  = @IntSerialNo And Branch =  @Branch Group By intserialNo)
SET   @CurrentRecievedAmount=(SELECT   isnull(sum(Bestankar),0 )    FROM         dbo.tblAcc_Recieved  WHERE  tblAcc_Recieved.intserialNo = @IntSerialNo    And AccountYear = @AccountYear ) + ((SELECT   isnull(sum(intAmount),0 )    FROM         dbo.[tFacCash]  WHERE  Branch = @Branch AND intserialNo = @intserialno )) +  ((SELECT   isnull(sum(intAmount),0 )    FROM         dbo.[tFacCard]  WHERE  Branch = @Branch AND intserialNo = @intserialno ))  + ((SELECT   isnull(sum(intAmount),0 )    FROM         dbo.[tFacCredit]  WHERE  Branch = @Branch AND intserialNo = @intserialno ))
SET   @PrePayment=(SELECT   isnull(sum(Bestankar),0 )     FROM         dbo.tblAcc_Recieved  WHERE    IntSerialNo = @IntSerialNo   And AccountYear = @AccountYear )

DECLARE @SumBuy BIGINT
SET   @SumBuy = 0
DECLARE @SumDaryaft BIGINT
SET   @SumDaryaft = 0
SET   @Credit = 0
SET   @CreditBuy = 0
SET   @RecievedAmount = 0

IF @Customer > 0
	BEGIN 
		Set   @Credit =  (Select ISNULL(Credit ,0) From tCust Where Code = @Customer )
		SET   @CreditBuy=( SELECT  isnull(sum(sumPrice),0 )    FROM         dbo.tFacM  WHERE     dbo.tFacm.Balance = 0  AND Recursive = 0 and Customer = @Customer And Status=2  And AccountYear = @AccountYear  And Branch =  @Branch)
		SET   @RecievedAmount=(SELECT   isnull(sum(Bestankar),0 )    FROM         dbo.tblAcc_Recieved  WHERE     RecieveType = 3 And Code_Bes = @Customer  And AccountYear = @AccountYear )
		SET   @SumBuy=( SELECT isnull(sum(sumPrice),0 )  FROM  dbo.tFacM  WHERE Customer = @Customer And Status=2  And AccountYear = @AccountYear  And Branch =  @Branch AND Recursive = 0)
		SELECT @SumDaryaft = SUM(Z.Daryaft) FROM  
		(SELECT ISNULL(SUM(tFacCash.intAmount), 0) + ISNULL(SUM(tblAcc_Recieved.Bestankar), 0)   AS Daryaft  
			FROM  dbo.tFacM
			LEFT OUTER JOIN dbo.tFacCash ON dbo.tFacCash.Branch = dbo.tFacM.Branch AND dbo.tFacCash.intSerialNo = dbo.tFacM.intSerialNo
			LEFT OUTER JOIN dbo.tblAcc_Recieved ON dbo.tblAcc_Recieved.Branch = dbo.tFacM.Branch AND dbo.tblAcc_Recieved.intSerialNo = dbo.tFacM.intSerialNo
				  WHERE     tfacm.AccountYear = @AccountYear
						AND Status = 2
						AND Recursive = 0
						AND tfacm.Branch = @Branch
						AND Customer = @Customer 
						--AND tfacm.intSerialNo <> @IntSerialNo
		UNION ALL  
		SELECT   ISNULL(SUM(tblAcc_Recieved.Bestankar) , 0) AS Daryaft 
			FROM  dbo.tblAcc_Recieved
				  WHERE     tblAcc_Recieved.AccountYear = @AccountYear
						AND tblAcc_Recieved.Branch >= @Branch
						AND Code_Bes >= @Customer 
						AND tblAcc_Recieved.intSerialNo IS NULL 
			)Z
	END 

IF @Mode <> dbo.GetNumericValue('ManipulateMode') 
	BEGIN
	    	SELECT distinct dbo.[VwInvoice_Multipart].levelcode1,dbo.[VwInvoice_Multipart].levelcode2,
	    	dbo.[VwInvoice_Multipart].leveldesc1,dbo.[VwInvoice_Multipart].leveldesc2,
	    	dbo.[VwInvoice_Multipart].LatinLeveldesc2,dbo.[VwInvoice_Multipart].unitdesc,dbo.[VwInvoice_Multipart].[intRow],
			dbo.[VwInvoice_Multipart].[No], dbo.[VwInvoice_Multipart].[Date], 
			dbo.[VwInvoice_Multipart].[SumPrice], dbo.[VwInvoice_Multipart].[Time], 
			dbo.[VwInvoice_Multipart].[User],dbo.[VwInvoice_Multipart].[UserName], dbo.[VwInvoice_Multipart].[Recursive], 
			dbo.[VwInvoice_Multipart].[StationId], dbo.[VwInvoice_Multipart].[masterserveplace], 
			dbo.[VwInvoice_Multipart].[ServePlace], dbo.[VwInvoice_Multipart].[OrderType], 
               		dbo.[VwInvoice_Multipart].[Status], dbo.[VwInvoice_Multipart].[GoodCode], 
			dbo.[VwInvoice_Multipart].[Weight], dbo.[VwInvoice_Multipart].[FeeUnit], 
			dbo.[VwInvoice_Multipart].[ShiftNo], dbo.[VwInvoice_Multipart].[FeeTotal], 
			dbo.[VwInvoice_Multipart].[intSerialNo], dbo.[VwInvoice_Multipart].[FacPayment], 
			dbo.[VwInvoice_Multipart].[InCharge], 
			dbo.[VwInvoice_Multipart].[Customer], dbo.[VwInvoice_Multipart].[Owner], 
			dbo.[VwInvoice_Multipart].[RegDate], 
			dbo.[VwInvoice_Multipart].[GarsonName], [VwInvoice_Multipart].[GarsonGender], 
			dbo.[VwInvoice_Multipart].[DifferencesDescription], [VwInvoice_Multipart].[TableDesc], 
			dbo.[VwInvoice_Multipart].[BascoleNo], [VwInvoice_Multipart].[Tel1], 
			dbo.[VwInvoice_Multipart].[Tel2], [VwInvoice_Multipart].[family], 
			dbo.[VwInvoice_Multipart].[DiscountTotal], [VwInvoice_Multipart].[CarryFeeTotal], 
			dbo.[VwInvoice_Multipart].[ServiceTotal], [VwInvoice_Multipart].[PackingTotal], 
			dbo.[VwInvoice_Multipart].[WeightTotal], [VwInvoice_Multipart].[Amount], 
			dbo.[VwInvoice_Multipart].[membershipid], [VwInvoice_Multipart].[PrinterNo], 
			dbo.[VwInvoice_Multipart].[Arm], [VwInvoice_Multipart].[Linefeed], 
			dbo.[VwInvoice_Multipart].[Cutter], 
			dbo.[VwInvoice_Multipart].[printformat],
       		        dbo.[VwInvoice_Multipart].[DirectRpt] , dbo.[VwInvoice_Multipart].[Balance] ,
			VWInvoice_MultiPart.Address + ' '+ VWInvoice_MultiPart.Flour + ' ' + VWInvoice_MultiPart.Unit 
			+ ' '+VWInvoice_MultiPart.InternalNo AS CustomerAddress,
	

	
			CASE @intLanguage 
				WHEN 0 THEN
					CASE @Mode 
						WHEN 1 THEN N'ç«Å „Ãœœ'
						WHEN 4 THEN N'«’·«ÕÌ'
						WHEN 8 THEN  N' €ÌÌ—« '
						ELSE ''
					END
				WHEN 1 THEN 
					CASE @Mode 
						WHEN 1 THEN N'Repeated Print'
						WHEN 4 THEN N'Edited'
                           				WHEN 8 THEN N'Change'
						ELSE ''
					END
			END AS ReportHeder,
			
			CASE @intLanguage 
				WHEN 0 THEN
					CASE dbo.VwInvoice_Multipart.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'„—ÃÊ⁄Ì'
					END
				WHEN 1 THEN 
					CASE dbo.VwInvoice_Multipart.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'Reffered'
					END
			END AS RecursievAlert,
	
		    	CASE @intLanguage 	
				WHEN 0 THEN dbo.VwInvoice_MultiPart.NamePrn
				WHEN 1 THEN dbo.VwInvoice_MultiPart.LatinNamePrn
			END AS GoodName,
			
	
			CASE @intLanguage 	
				WHEN 0 THEN VWInvoice_MultiPart.ItemServePlace
				WHEN 1 THEN VWInvoice_MultiPart.ItemLatinServePlace
			END AS ItemServePlaceDesc,
	
			CASE @intLanguage 	
				WHEN 0 THEN VWInvoice_MultiPart.NoticeDescription1
				WHEN 1 THEN VWInvoice_MultiPart.NoticeLatinDescription
			END AS NoticeDescription,
	
			CASE @intLanguage 	
				WHEN 0 THEN VWInvoice_MultiPart.FactorServePlace
				WHEN 1 THEN VWInvoice_MultiPart.FactorLatinServePlace
			END AS FactorServeDescription,

	      	          CASE VWInvoice_MultiPart.barcode 
				WHEN 1 THEN  (SELECT TOP 1  dbo.BarcodeGenerator(dbo.VWInvoice_MultiPart.ServePlace,@intFacNo) where [No]= @intFacNo and Status = 2 And AccountYear = @AccountYear )
				ELSE '' END AS Barcode ,
                                        
			dbo.[VwInvoice_Multipart].[UnitType], dbo.VwInvoice_MultiPart.LatinNamePrn As LatinGoodname
 			, dbo.[VwInvoice_Multipart].[Rate] , dbo.[VwInvoice_Multipart].[ChairName],dbo.[VwInvoice_Multipart].[MainType]
 			, VwInvoice_MultiPart.NumberOfUnit ,dbo.[VwInvoice_Multipart].NvcDescription , dbo.VwInvoice_MultiPart.NamePrn  
			, TaxBuy , DutyBuy , TaxSale , DutySale 
			, dbo.VwInvoice_Multipart.TaxTotal , dbo.VwInvoice_Multipart.DutyTotal , dbo.numbertoharf(dbo.VwInvoice_Multipart.sumprice , 0 ) AS SumpriceHarf
			, VWInvoice_MultiPart.PartitionId , VWInvoice_MultiPart.PartitionDescription , VWInvoice_MultiPart.GuestNo , VWInvoice_MultiPart.TempNo
			, VWInvoice_MultiPart.OrderDate , VWInvoice_MultiPart.OrderTime 
			, dbo.[VwInvoice_Multipart].GoodNamePrn2 , dbo.[VwInvoice_Multipart].GoodNamePrn3
			, @CurrentBuy As CurrentBuy ,@Maxguest As MaxGuest ,@MaxBuy As MaxBuy ,@familycount As FamilyCount,@Central As Central ,@Member As Member , @MainTypeNo As MainTypeNo
			, @ChequeAmount as ChequeAmount,@CreditBuy as CreditBuy,@RecievedAmount as RecievedAmount , @NvcDescription As NvcDescription ,@TempAddress AS TempAddress
			, @Credit As Credit , @CurrentRecievedAmount as CurrentRecievedAmount , ISNULL(@PrePayment , 0) as PrePayment
            , @m1 as m1 , @m2 as m2 , @m3 as m3 , @m4 as m4,N' www.FGARYA.com-‰—„ «›“«—  ”„— -02188554455 ' as CompanyName
            , @SumBuy AS SumBuy , @SumDaryaft AS SumDaryaft , RTRIM(LTRIM(@nvcGreatSpeech)) AS nvcGreatSpeech
            
		FROM dbo.VWInvoice_MultiPart
		
		  WHERE     VWInvoice_MultiPart.[No]=@intFacNo 	
			AND VWInvoice_MultiPart.PrintFormat  = @PrintFormat 
			AND GoodCode NOT IN (SELECT GoodCode  FROM tPrinterGood WHERE intPrinterFormat = @PrintFormat )
			AND VWInvoice_MultiPart.StationId = @StationId 
			AND VWInvoice_MultiPart.PartitionId = @PartitionId
			AND VWInvoice_MultiPart.status =@Status 
			And VWInvoice_MultiPart.AccountYear = @AccountYear
			AND VWInvoice_MultiPart.PrinterNo=@intPrinterNo
			AND VWInvoice_MultiPart.permittedModes & @Mode = @Mode
	Order By   GoodCode Asc
	END
Else if @Mode = dbo.GetNumericValue('ManipulateMode')   ---And @PrintFormat = 3

	BEGIN
        Select T3.* 
		, @CurrentBuy As CurrentBuy ,@Maxguest As MaxGuest ,@MaxBuy As MaxBuy ,@familycount As FamilyCount,@Central As Central ,@Member As Member , @MainTypeNo As MainTypeNo
		, @ChequeAmount as ChequeAmount,@CreditBuy as CreditBuy,@RecievedAmount as RecievedAmount , @NvcDescription As NvcDescription ,@TempAddress AS TempAddress
		, @Credit As Credit , @CurrentRecievedAmount as CurrentRecievedAmount , ISNULL(@PrePayment , 0) as PrePayment
        , @m1 as m1 , @m2 as m2 , @m3 as m3 , @m4 as m4,N' www.FGARYA.com-‰—„ «›“«—  ”„— -02188554455 ' as CompanyName
        , @SumBuy AS SumBuy , @SumDaryaft AS SumDaryaft , RTRIM(LTRIM(@nvcGreatSpeech)) AS nvcGreatSpeech
 	From
	(Select  ISNULL(T1.levelcode1 , T2.levelcode1) as levelcode1, ISNULL(T1.levelcode2 , T2.levelcode2) as levelcode2,
		ISNULL(T1.leveldesc1 , T2.leveldesc1) as leveldesc1, ISNULL(T1.leveldesc2 , T2.leveldesc2) as leveldesc2,
		ISNULL(T1.LatinLeveldesc2 , T2.LatinLeveldesc2) as LatinLeveldesc2,
		ISNULL(T1.unitdesc , T2.unitdesc) as unitdesc,ISNULL(T1.intRow , T2.intRow) as intRow,
		ISNULL(T1.No , T2.No) as No,ISNULL(T1.Date , T2.Date) as Date,ISNULL(T1.SumPrice , T2.SumPrice) as SumPrice,
		ISNULL(T1.Time , T2.Time) as Time,ISNULL(T1.[User] , T2.[User]) as [User],ISNULL(T1.UserName , T2.UserName) as UserName,ISNULL(T1.Recursive , T2.Recursive) as Recursive,
		ISNULL(T1.StationId , T2.StationId) as StationId,ISNULL(T1.masterserveplace , T2.masterserveplace) as masterserveplace,
		ISNULL(T1.ServePlace , T2.ServePlace) as ServePlace,ISNULL(T1.OrderType , T2.OrderType) as OrderType,ISNULL(T1.Status , T2.Status) as Status,
		ISNULL(T1.GoodCode , T2.GoodCode) as GoodCode,ISNULL(T1.Weight , T2.Weight) as Weight,ISNULL(T1.FeeUnit , T2.FeeUnit) as FeeUnit,
		ISNULL(T1.ShiftNo , T2.ShiftNo) as ShiftNo,ISNULL(T1.FeeTotal ,T2.FeeTotal) as FeeTotal,ISNULL(T1.intSerialNo , T2.intSerialNo) as intSerialNo,
		ISNULL(T1.FacPayment , T2.FacPayment)as FacPayment,ISNULL(T1.InCharge , T2.InCharge) as InCharge,
		--ISNULL(T1.FactorServePlace , T2.FactorServePlace) as FactorServePlace,ISNULL(T1.FactorLatinServePlace , T2.FactorLatinServePlace) as FactorLatinServePlace,
		ISNULL(T1.Customer , T2.Customer)as Customer,ISNULL(T1.Owner , T2.Owner) as Owner,ISNULL(T1.RegDate , T2.RegDate) as RegDate,
		--ISNULL(T1.Address , T2.Address) as Address,
		--ISNULL(T1.Unit , T2.Unit) as Unit,ISNULL(T1.InternalNo , T2.InternalNo) as InternalNo,ISNULL(T1.Flour , T2.Flour) as Flour ,
		ISNULL(T1.GarsonName , T2.GarsonName) as GarsonName,ISNULL(T1.GarsonGender , T2.GarsonGender) as GarsonGender,		
		ISNULL(T1.DifferencesDescription , T2.DifferencesDescription) as DifferencesDescription,ISNULL(T1.TableDesc , T2.TableDesc) as TableDesc,
		ISNULL(T1.BascoleNo , T2.BascoleNo) as BascoleNo,ISNULL(T1.Tel1 , T2.Tel1) as Tel1,ISNULL(T1.Tel2 , T2.Tel2) as Tel2,	
		ISNULL(T1.family , T2.family) as family,ISNULL(T1.DiscountTotal , T2.DiscountTotal) as DiscountTotal,ISNULL(T1.CarryFeeTotal , T2.CarryFeeTotal) as CarryFeeTotal,	
		ISNULL(T1.ServiceTotal , T2.ServiceTotal) as ServiceTotal,ISNULL(T1.PackingTotal , T2.PackingTotal) as PackingTotal,
		ISNULL(T1.WeightTotal , T2.WeightTotal) as WeightTotal,ISNULL(T1.Amount , 0) - ISNULL(T2.Amount , 0) as Amount,ISNULL(T1.membershipid , T2.membershipid) as membershipid,
		ISNULL(T1.PrinterNo , T2.PrinterNo) as PrinterNo,ISNULL(T1.Arm,T2.Arm) as Arm,ISNULL(T1.Linefeed , T2.Linefeed) as Linefeed,
		ISNULL(T1.Cutter , T2.Cutter) as Cutter , --ISNULL(T1.ItemServePlace , T2.ItemServePlace) as ItemServePlace,
		--ISNULL(T1.Description , T2.Description) as Description,
		ISNULL(T1.printformat , T2.printformat) as printformat,
		--ISNULL(T1.NoticeDescription1 , T2.NoticeDescription1) as NoticeDescription1,ISNULL(T1.NoticeLatinDescription , T2.NoticeLatinDescription) as NoticeLatinDescription,
		--ISNULL(T1.LatinDescription , T2.LatinDescription) as LatinDescription,ISNULL(T1.ItemLatinServePlace , T2.ItemLatinServePlace) as ItemLatinServePlace,
		ISNULL(T1.DirectRpt , T2.DirectRpt) as DirectRpt, ISNULL(T1.Balance , T2.Balance) as Balance ,
		ISNULL(T1.CustomerAddress , T2.CustomerAddress) as CustomerAddress,
		ISNULL(T1.ReportHeder , T2.ReportHeder) as ReportHeder,ISNULL(T1.RecursievAlert , T2.RecursievAlert) as RecursievAlert,
		ISNULL(T1.GoodName , T2.GoodName) as GoodName ,  -- ISNULL(T1.ServePlaceDesc , T2.ServePlaceDesc) as ServePlaceDesc,
		ISNULL(T1.ItemServePlaceDesc , T2.ItemServePlaceDesc) as ItemServePlaceDesc,ISNULL(T1.NoticeDescription , T2.NoticeDescription) as NoticeDescription,
		ISNULL(T1.FactorServeDescription , T2.FactorServeDescription) as FactorServeDescription , ISNULL(T1.Barcode , T2.Barcode) as Barcode	,
		ISNULL(T1.UnitType , T2.UnitType) as UnitType ,ISNULL(T1.LatinNamePrn , T2.LatinNamePrn) as LatinNamePrn ,
		ISNULL(T1.Rate , T2.Rate) as Rate ,
		ISNULL(T1.ChairName , T2.ChairName) as ChairName ,ISNULL(T1.MainType , T2.MainType) as MainType ,
		ISNULL(T1.NumberOfUnit , T2.NumberOfUnit) as NumberOfUnit
		, T1.nvcDescription , ISNULL(T1.NamePrn ,T2.NamePrn ) AS NamePrn , T1.TaxBuy , T1.DutyBuy , T1.TaxSale , T1.DutySale
		, T1.TaxTotal , T1.DutyTotal , dbo.numbertoharf(T1.sumprice , 0 ) AS SumpriceHarf
		, T1.PartitionId , T1.PartitionDescription , T1.GuestNo , T1.TempNo
		, T1.OrderDate , T1.OrderTime , T1.GoodNamePrn2 , T1.GoodNamePrn3
		FROM
	    	(SELECT distinct dbo.[VwInvoice_Multipart].levelcode1,dbo.[VwInvoice_Multipart].levelcode2,
	    	dbo.[VwInvoice_Multipart].leveldesc1,dbo.[VwInvoice_Multipart].leveldesc2,
	    	dbo.[VwInvoice_Multipart].LatinLeveldesc2,dbo.[VwInvoice_Multipart].unitdesc ,
	    	dbo.[VwInvoice_Multipart].[intRow], 
			dbo.[VwInvoice_Multipart].[No], dbo.[VwInvoice_Multipart].[Date], 
			dbo.[VwInvoice_Multipart].[SumPrice], dbo.[VwInvoice_Multipart].[Time], 
			dbo.[VwInvoice_Multipart].[User],dbo.[VwInvoice_Multipart].[UserName], dbo.[VwInvoice_Multipart].[Recursive], 
			dbo.[VwInvoice_Multipart].[StationId], dbo.[VwInvoice_Multipart].[masterserveplace], 
			dbo.[VwInvoice_Multipart].[ServePlace], dbo.[VwInvoice_Multipart].[OrderType], 
			dbo.[VwInvoice_Multipart].[Status], dbo.[VwInvoice_Multipart].[GoodCode], 
			dbo.[VwInvoice_Multipart].[Weight], dbo.[VwInvoice_Multipart].[FeeUnit], 
			dbo.[VwInvoice_Multipart].[ShiftNo], dbo.[VwInvoice_Multipart].[FeeTotal], 
			dbo.[VwInvoice_Multipart].[intSerialNo], dbo.[VwInvoice_Multipart].[FacPayment], 
			dbo.[VwInvoice_Multipart].[InCharge], 
			
			dbo.[VwInvoice_Multipart].[Customer], dbo.[VwInvoice_Multipart].[Owner], 
			dbo.[VwInvoice_Multipart].[RegDate],
			
			dbo.[VwInvoice_Multipart].[GarsonName], [VwInvoice_Multipart].[GarsonGender], 
			dbo.[VwInvoice_Multipart].[DifferencesDescription], [VwInvoice_Multipart].[TableDesc], 
			dbo.[VwInvoice_Multipart].[BascoleNo], [VwInvoice_Multipart].[Tel1], 
			dbo.[VwInvoice_Multipart].[Tel2], [VwInvoice_Multipart].[family], 
			dbo.[VwInvoice_Multipart].[DiscountTotal], [VwInvoice_Multipart].[CarryFeeTotal], 
			dbo.[VwInvoice_Multipart].[ServiceTotal], [VwInvoice_Multipart].[PackingTotal], 
			dbo.[VwInvoice_Multipart].[WeightTotal], [VwInvoice_Multipart].[Amount], 
			dbo.[VwInvoice_Multipart].[membershipid], [VwInvoice_Multipart].[PrinterNo], 
			dbo.[VwInvoice_Multipart].[Arm], [VwInvoice_Multipart].[Linefeed], 
			dbo.[VwInvoice_Multipart].[Cutter], 

			
			dbo.[VwInvoice_Multipart].[printformat],
       		             dbo.[VwInvoice_Multipart].[DirectRpt] ,  dbo.[VwInvoice_Multipart].[Balance] ,
			VWInvoice_MultiPart.Address + ' '+ VWInvoice_MultiPart.Flour + ' ' + VWInvoice_MultiPart.Unit 
			+ ' '+VWInvoice_MultiPart.InternalNo AS CustomerAddress,
	
	
			CASE @intLanguage 
				WHEN 0 THEN
					CASE @Mode 
						WHEN 1 THEN N'ç«Å „Ãœœ'

						WHEN 4 THEN N'«’·«ÕÌ'
						WHEN 8 THEN N' €ÌÌ—« '
						Else ''
					END
				WHEN 1 THEN 
					CASE @Mode 
						WHEN 1 THEN N'Repeated Print'
						WHEN 4 THEN N'Edited'
						WHEN 8 THEN N'Change'
						Else ''
					END
			END AS ReportHeder,
			
			CASE @intLanguage 
				WHEN 0 THEN
					CASE dbo.VwInvoice_Multipart.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'„—ÃÊ⁄Ì'
					END
				WHEN 1 THEN 
					CASE dbo.VwInvoice_Multipart.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'Reffered'
					END
			END AS RecursievAlert,
	
		    	CASE @intLanguage 	
				WHEN 0 THEN dbo.VwInvoice_MultiPart.NamePrn
				WHEN 1 THEN dbo.VwInvoice_MultiPart.LatinNamePrn

			END AS GoodName,
			
	
			CASE @intLanguage 	
				WHEN 0 THEN VWInvoice_MultiPart.ItemServePlace
				WHEN 1 THEN VWInvoice_MultiPart.ItemLatinServePlace
			END AS ItemServePlaceDesc,
	
			CASE @intLanguage 	
				WHEN 0 THEN VWInvoice_MultiPart.NoticeDescription1
				WHEN 1 THEN VWInvoice_MultiPart.NoticeLatinDescription
			END AS NoticeDescription,
	
			CASE @intLanguage 	
				WHEN 0 THEN VWInvoice_MultiPart.FactorServePlace
				WHEN 1 THEN VWInvoice_MultiPart.FactorLatinServePlace
			END AS FactorServeDescription,

		CASE VWInvoice_MultiPart.barcode 
				WHEN 1 THEN  (SELECT TOP 1  dbo.BarcodeGenerator(dbo.VWInvoice_MultiPart.ServePlace,@intFacNo) where [No]= @intFacNo and Status = 2 )
				ELSE '' END AS Barcode ,

		dbo.[VwInvoice_Multipart].[UnitType] , dbo.VwInvoice_MultiPart.LatinNamePrn 
		, dbo.[VwInvoice_Multipart].[Rate]  , dbo.[VwInvoice_Multipart].[ChairName], dbo.[VwInvoice_Multipart].[MainType], dbo.[VwInvoice_Multipart].NumberOfUnit
		, dbo.[VwInvoice_Multipart].NvcDescription , dbo.VwInvoice_MultiPart.NamePrn 
		, TaxBuy , DutyBuy , TaxSale , DutySale
		, dbo.VwInvoice_Multipart.TaxTotal , dbo.VwInvoice_Multipart.DutyTotal 
		, dbo.VwInvoice_Multipart.PartitionId , dbo.VwInvoice_Multipart.PartitionDescription 
		, dbo.VwInvoice_Multipart.GuestNo 
		, VWInvoice_MultiPart.TempNo
		, VWInvoice_MultiPart.OrderDate , VWInvoice_MultiPart.OrderTime 
		, dbo.[VwInvoice_Multipart].GoodNamePrn2 , dbo.[VwInvoice_Multipart].GoodNamePrn3
	FROM dbo.VWInvoice_MultiPart 
		
		WHERE 	No = @intFacNo 	
			AND (PrintFormat  = @PrintFormat OR @PrintFormat =0 )
			AND GoodCode NOT IN (SELECT GoodCode  FROM tPrinterGood WHERE intPrinterFormat = @PrintFormat )
			AND VWInvoice_MultiPart.StationId = @StationId 
			AND VWInvoice_MultiPart.PartitionId = @PartitionId
			AND dbo.VWInvoice_MultiPart.status =@Status And VWInvoice_MultiPart.AccountYear = @AccountYear
			AND   VWInvoice_MultiPart.PrinterNo=@intPrinterNo 
			AND   VWInvoice_MultiPart.permittedModes & @Mode = @Mode
			) T1
		Full Outer Join
		(SELECT dbo.[VwInvoice_Multipart2].levelcode1,dbo.[VwInvoice_Multipart2].levelcode2,
	    	dbo.[VwInvoice_Multipart2].leveldesc1,dbo.[VwInvoice_Multipart2].leveldesc2,
	    	dbo.[VwInvoice_Multipart2].LatinLeveldesc2,dbo.[VwInvoice_Multipart2].unitdesc,dbo.[VwInvoice_Multipart2].[intRow], 
			dbo.[VwInvoice_Multipart2].[No], dbo.[VwInvoice_Multipart2].[Date], 
			dbo.[VwInvoice_Multipart2].[SumPrice], dbo.[VwInvoice_Multipart2].[Time], 
			dbo.[VwInvoice_Multipart2].[User], dbo.[VwInvoice_Multipart2].[UserName],dbo.[VwInvoice_Multipart2].[Recursive], 
			dbo.[VwInvoice_Multipart2].[StationId], dbo.[VwInvoice_Multipart2].[masterserveplace], 
			dbo.[VwInvoice_Multipart2].[ServePlace], dbo.[VwInvoice_Multipart2].[OrderType], 
			dbo.[VwInvoice_Multipart2].[Status], dbo.[VwInvoice_Multipart2].[GoodCode], 
			dbo.[VwInvoice_Multipart2].[Weight], dbo.[VwInvoice_Multipart2].[FeeUnit], 
			dbo.[VwInvoice_Multipart2].[ShiftNo], dbo.[VwInvoice_Multipart2].[FeeTotal], 
			dbo.[VwInvoice_Multipart2].[intSerialNo], dbo.[VwInvoice_Multipart2].[FacPayment], 
			dbo.[VwInvoice_Multipart2].[InCharge], 			
			dbo.[VwInvoice_Multipart2].[Customer], dbo.[VwInvoice_Multipart2].[Owner], 
			dbo.[VwInvoice_Multipart2].[RegDate], 
			dbo.[VwInvoice_Multipart2].[GarsonName], [VwInvoice_Multipart2].[GarsonGender], 
			dbo.[VwInvoice_Multipart2].[DifferencesDescription], (Select [Name] From  dbo.ttable Where [VwInvoice_Multipart2].[TableDesc] = ttable.[No]) as TableDesc , 
			dbo.[VwInvoice_Multipart2].[BascoleNo], [VwInvoice_Multipart2].[Tel1], 
			dbo.[VwInvoice_Multipart2].[Tel2], [VwInvoice_Multipart2].[family], 
			dbo.[VwInvoice_Multipart2].[DiscountTotal], [VwInvoice_Multipart2].[CarryFeeTotal], 
			dbo.[VwInvoice_Multipart2].[ServiceTotal], [VwInvoice_Multipart2].[PackingTotal], 
			dbo.[VwInvoice_Multipart2].[WeightTotal], [VwInvoice_Multipart2].[Amount], 
			dbo.[VwInvoice_Multipart2].[membershipid], [VwInvoice_Multipart2].[PrinterNo], 
			dbo.[VwInvoice_Multipart2].[Arm], [VwInvoice_Multipart2].[Linefeed], 
			dbo.[VwInvoice_Multipart2].[Cutter], 
			
			dbo.[VwInvoice_Multipart2].[printformat], 
       		             dbo.[VwInvoice_Multipart2].[DirectRpt] ,  dbo.[VwInvoice_Multipart2].[Balance] ,
			VwInvoice_Multipart2.Address + ' '+ VwInvoice_Multipart2.Flour + ' ' + VwInvoice_Multipart2.Unit 
			+ ' '+VwInvoice_Multipart2.InternalNo AS CustomerAddress,
	
			CASE @intLanguage 
				WHEN 0 THEN
					CASE @Mode 
						WHEN 1 THEN N'ç«Å „Ãœœ'
						WHEN 4 THEN N'«’·«ÕÌ'
						WHEN 8 THEN N' €ÌÌ—« '
						ELSE ''
					END
				WHEN 1 THEN 
					CASE @Mode 

						WHEN 1 THEN N'Repeated Print'						WHEN 4 THEN N'Edited'
						WHEN 8 THEN N'Change'
						ELSE ''
					END
			END AS ReportHeder,
			
			CASE @intLanguage 
				WHEN 0 THEN
					CASE dbo.VwInvoice_Multipart2.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'„—ÃÊ⁄Ì'
					END
				WHEN 1 THEN 
					CASE dbo.VwInvoice_Multipart2.Recursive 	
						WHEN  0 THEN ''
						WHEN 1 THEN N'Reffered'
					END
			END AS RecursievAlert,
	
		    	CASE @intLanguage 	
				WHEN 0 THEN dbo.VwInvoice_Multipart2.NamePrn
				WHEN 1 THEN dbo.VwInvoice_Multipart2.LatinNamePrn
			END AS GoodName,
			

			CASE @intLanguage 	
				WHEN 0 THEN VwInvoice_Multipart2.ItemServePlace
				WHEN 1 THEN VwInvoice_Multipart2.ItemLatinServePlace
			END AS ItemServePlaceDesc,
	
			CASE @intLanguage 	
				WHEN 0 THEN VwInvoice_Multipart2.NoticeDescription1
				WHEN 1 THEN VwInvoice_Multipart2.NoticeLatinDescription
			END AS NoticeDescription,
	
			CASE @intLanguage 	
				WHEN 0 THEN VwInvoice_Multipart2.FactorServePlace
				WHEN 1 THEN VwInvoice_Multipart2.FactorLatinServePlace
			END AS FactorServeDescription,

		CASE VwInvoice_Multipart2.barcode 
				WHEN 1 THEN  (SELECT dbo.BarcodeGenerator(dbo.VwInvoice_Multipart2.ServePlace,@intFacNo))
				ELSE '' END AS Barcode ,

		dbo.[VwInvoice_Multipart2].[UnitType] , dbo.VwInvoice_Multipart2.LatinNamePrn 
		, dbo.[VwInvoice_Multipart2].[Rate] ,dbo.[VwInvoice_Multipart2].[ChairName], dbo.[VwInvoice_Multipart2].[MainType],@CreditBuy as CreditBuy,@RecievedAmount as RecievedAmount , @Credit As Credit
		,dbo.[VwInvoice_Multipart2].NumberOfUnit , dbo.VwInvoice_MultiPart2.NamePrn 
		, TaxBuy , DutyBuy , TaxSale , DutySale
		, VwInvoice_Multipart2.TaxTotal , DutyTotal
		FROM dbo.VwInvoice_Multipart2
		
		WHERE 	[No]=@intFacNo 	
			AND (PrintFormat  = @PrintFormat OR @PrintFormat =0 OR ((@PrintFormat NOT IN (SELECT PrintFormat FROM dbo.VWInvoice_MultiPart WHERE IntSerialNo=@IntSerialNo)
				AND (PrintFormat = (SELECT TOP 1 PrintFormat FROM dbo.VWInvoice_MultiPart WHERE IntSerialNo=@IntSerialNo)))) )
			AND GoodCode NOT IN (SELECT GoodCode  FROM tPrinterGood WHERE intPrinterFormat = @PrintFormat )
			AND VWInvoice_MultiPart2.StationId = @StationId 
			AND VWInvoice_MultiPart2.PartitionId = @PartitionId
			AND dbo.VwInvoice_Multipart2.status =@Status And VWInvoice_MultiPart2.AccountYear = @AccountYear
			AND   VwInvoice_Multipart2.PrinterNo=@intPrinterNo
			AND   VWInvoice_MultiPart2.permittedModes & @Mode = @Mode
 			And Code = (Select Max(Code) from VwInvoice_Multipart2 where [No]=@intFacNo AND dbo.VwInvoice_Multipart2.status =@Status And VWInvoice_MultiPart2.AccountYear = @AccountYear))T2
		on 
		T1.intSerialNo = T2.intSerialNo And 
		T1.GoodCode = T2.GoodCode And 
		T1.ServePlace = T2.ServePlace And
		T1.DifferencesDescription = T2.DifferencesDescription
		Where ISNULL(T1.Amount,0)-ISNULL(T2.Amount,0) <> 0 
           )T3   
	Order By   GoodCode Asc
	END





GO
