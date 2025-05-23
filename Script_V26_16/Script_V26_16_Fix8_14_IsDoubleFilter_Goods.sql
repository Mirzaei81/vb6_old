

--Script_V26_16_Fix8_14_IsDoubleFilter_Goods.sql
--  فانکشن کالاهای اصلی را از فاکتور دریافت می کند
--میتوان چک کرد اگر در فیش های آشپزخانه کالای اصلی نداشته باشد کالاهای فرعی هم چاپ نشود
--برای پیتزا خاتون و سایر مشتریان هم قابل استفاده است
-- در پروسیچر [Get_InvoiceInfo_IsFilter] 
--باید مقدار متغیر @IsDoubleFilter = 1 باشد

--93/08/23

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FN_IF_GetMainGood]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[FN_IF_GetMainGood]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


-- این فانکشن کالاهای اصلی را از فاکتور دریافت می کند
CREATE  Function [dbo].[FN_IF_GetMainGood]

(
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

RETURNS  @ReturnTable TABLE(
 GoodCode BIGINT 
)	
As

BEGIN

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()

Declare @intserialNo int

Set   @IntSerialNo = (Select intSerialNo From tFacm Where [No]=@intFacNo And Status=@Status  And AccountYear = @AccountYear  And Branch =  @Branch)


IF @Mode <> dbo.GetNumericValue('ManipulateMode') 
	BEGIN
			INSERT INTO @ReturnTable
			
	    	SELECT distinct 
	    	dbo.[VwInvoice_Multipart].[GoodCode]
		FROM dbo.VWInvoice_MultiPart
		  WHERE     VWInvoice_MultiPart.[No]=@intFacNo 	
			AND VWInvoice_MultiPart.PrintFormat  = @PrintFormat 
			AND VWInvoice_MultiPart.StationId = @StationId 
			AND VWInvoice_MultiPart.PartitionId = @PartitionId
			AND VWInvoice_MultiPart.status =@Status 
			And VWInvoice_MultiPart.AccountYear = @AccountYear
			AND VWInvoice_MultiPart.PrinterNo=@intPrinterNo
			AND VWInvoice_MultiPart.permittedModes & @Mode = @Mode
			AND VWInvoice_MultiPart.GoodCode not in (select dbo.tPrinterGood.goodcode from dbo.tPrinterGood WHERE intPrinterFormat = @PrintFormat)
			AND MainType = 1

	Order By   GoodCode Asc
END
Else if @Mode = dbo.GetNumericValue('ManipulateMode')   ---And @PrintFormat = 3
	BEGIN

    	INSERT INTO @ReturnTable
	    Select T3.GoodCode From
		(Select  
	    	ISNULL(T1.GoodCode , T2.GoodCode) AS GoodCode ,
	    	ISNULL(T1.MainType , T2.MainType) AS MainType
		FROM
	    	(SELECT distinct 
	    	dbo.[VwInvoice_Multipart].[GoodCode] ,
			dbo.[VwInvoice_Multipart].[intSerialNo],
			dbo.[VwInvoice_Multipart].[ServePlace],
			dbo.[VwInvoice_Multipart].[DifferencesDescription],
			dbo.[VwInvoice_Multipart].[Amount] ,
			dbo.[VwInvoice_Multipart].MainType
			FROM dbo.VWInvoice_MultiPart 
		
		WHERE 	No = @intFacNo 	
			AND (PrintFormat  = @PrintFormat OR @PrintFormat =0 )
			AND VWInvoice_MultiPart.StationId = @StationId 
			AND VWInvoice_MultiPart.PartitionId = @PartitionId
			AND dbo.VWInvoice_MultiPart.status =@Status And VWInvoice_MultiPart.AccountYear = @AccountYear
			AND   VWInvoice_MultiPart.PrinterNo=@intPrinterNo 
			AND   VWInvoice_MultiPart.permittedModes & @Mode = @Mode
			) T1
		Full Outer Join
		(SELECT 
	    	dbo.[VwInvoice_Multipart2].[GoodCode] ,
			dbo.[VwInvoice_Multipart2].[intSerialNo],
			dbo.[VwInvoice_Multipart2].[ServePlace],
			dbo.[VwInvoice_Multipart2].[DifferencesDescription],
			dbo.[VwInvoice_Multipart2].[Amount] ,
			dbo.[VwInvoice_Multipart2].MainType
			FROM dbo.VWInvoice_MultiPart2 
		WHERE 	[No]=@intFacNo 	
			AND (PrintFormat  = @PrintFormat OR @PrintFormat =0 OR ((@PrintFormat NOT IN (SELECT PrintFormat FROM dbo.VWInvoice_MultiPart WHERE IntSerialNo=@IntSerialNo)
				AND (PrintFormat = (SELECT TOP 1 PrintFormat FROM dbo.VWInvoice_MultiPart WHERE IntSerialNo=@IntSerialNo)))) )
			AND VWInvoice_MultiPart2.StationId = @StationId 
			AND VWInvoice_MultiPart2.PartitionId = @PartitionId
			AND dbo.VwInvoice_Multipart2.status =@Status 
			And VWInvoice_MultiPart2.AccountYear = @AccountYear
			AND  VwInvoice_Multipart2.PrinterNo=@intPrinterNo
			AND  VWInvoice_MultiPart2.permittedModes & @Mode = @Mode
 			And Code = (Select Max(Code) from VwInvoice_Multipart2 where [No]=@intFacNo AND dbo.VwInvoice_Multipart2.status =@Status And VWInvoice_MultiPart2.AccountYear = @AccountYear))T2
		on 
		T1.intSerialNo = T2.intSerialNo And 
		T1.GoodCode = T2.GoodCode And 
		T1.ServePlace = T2.ServePlace And
		T1.DifferencesDescription = T2.DifferencesDescription
		Where ISNULL(T1.Amount,0)-ISNULL(T2.Amount,0) <> 0 
           ) T3   
	WHERE T3.MainType = 1
	AND T3.GoodCode not in (select dbo.tPrinterGood.goodcode from dbo.tPrinterGood WHERE intPrinterFormat = @PrintFormat)
	Order By   GoodCode Asc
	END
	

RETURN 

END



GO



--SELECT * FROM  from dbo.[FN_IF_GetMainGood](@intFacNo,@PrintFormat,@StationId,@Status,@intPrinterNo,@Mode,@AccountYear,@PartitionId,@Branch)
--SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END AS ExistDoods FROM  dbo.[FN_IF_GetMainGood](265,3,1,2,1,1,1393,1,1)


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Get_InvoiceInfo_IsFilter] (

	
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

--For Double Filter in All report except invoicefich & invoiceFactor
DECLARE @IsDoubleFilter BIT 
SET @IsDoubleFilter = 0

DECLARE @Branch INT
IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()
Declare  @ExistGood BIT
SET @ExistGood = 1
IF @IsDoubleFilter = 1
	SELECT @ExistGood = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END from dbo.[FN_IF_GetMainGood](@intFacNo,@PrintFormat,@StationId,@Status,@intPrinterNo,@Mode,@AccountYear,@PartitionId,@Branch)


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

END


GO


