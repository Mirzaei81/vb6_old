

--Script_V26_16_Fix12
--اضافه کردن فیلد برای کنترل رکوردهاییکه باید در محاسبه
--مغایرت گیری شرکت کنند
--فقط کالاهاییکه فیلد فعال آنها تیک داشته باشد در مغایرت گیری شرکت می کنند
--سپس سند کسر و اضافه انبار برای این کالاها صادر می گردد
--آپدیت کردن قیمت خرید کالا با آخرین قیمت خرید آن کالا
--جستجوی ارسالی ها با اشتراک و نام و تلفن
--93/12/05
--Script_V26_16_Fix12_1_FirstMojodiControl

 --گزاشتن دسترسی روی  موجودی اولیه   
--کد پرسنلی پیک 4 رقمی شده و بارکد آن اصلاح شد
--برای همه ورژن ها قابل استقاده است
--93/12/22
--Script_V26_16_Fix12_2_InsertCustomerInNetwork
--ثبت اشتراک همزمان در شبکه
--V26_16_Fix5 برای کلیه ورژن های بعد از 
--93/12/23

INSERT  INTO dbo.tblPub_Script2
        ( Version ,
          Script ,
          LastScriptNo ,
          [Date] ,
          FixNumber
        )
VALUES  ( 26 ,
          16 ,
          0 ,
          dbo.shamsi(GETDATE()) ,
          12
        )
GO


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



--Script_V26_16_Fix12_1_FirstMojodiControl

 --گزاشتن دسترسی روی  موجودی اولیه   
--کد پرسنلی پیک 4 رقمی شده و بارکد آن اصلاح شد
--برای همه ورژن ها قابل استقاده است
--93/12/22


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 328 , -- intObjectCode - int
          N'FirstMojodiControl' , -- ObjectId - nvarchar(50)
          N'کنترل موجودی اولیه' , -- ObjectName - nvarchar(50)
          N'FirstMojodiControl' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          328  -- intObjectCode - int
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
VALUES  ( 329 , -- intObjectCode - int
          N'frmOrder' , -- ObjectId - nvarchar(50)
          N'فرم سفارشات' , -- ObjectName - nvarchar(50)
          N'frmOrder' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          329  -- intObjectCode - int
          )
GO



--Script_PaykBarcode
--کد پرسنلی پیک 4 رقمی شده و بارکد آن اصلاح شد
--برای همه ورژن ها قابل استقاده است


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  FUNCTION dbo.PersonelBarcodeGenerator
(
	@JobID INT,
	@PPNO   INT
)
RETURNS  NVARCHAR(20)

AS

BEGIN


	DECLARE @strJobID    NVARCHAR(10)
	DECLARE @strPPNO     NVARCHAR(10)
	DECLARE @Tmp         NVARCHAR(20)
	DECLARE @ZeroCount   INT


	SET @ZeroCount = 2 - LEN(CAST(@JobID AS NVARCHAR(10)))
	SET @strJobID  = (SELECT dbo.Repeater('0',@ZeroCount)) + CAST(@JobID AS NVARCHAR(10)) 

	SET @ZeroCount = 4 - LEN(CAST(@PPNO AS NVARCHAR(10)))
	SET @strPPNO   = (SELECT dbo.Repeater('0',@ZeroCount)) + CAST(@PPNO AS NVARCHAR(10)) 

	--SET @Tmp = @strJobID + (SELECT dbo.Repeater('0',7)) + @strPPNO   --12 number is correct
	SET @Tmp = @strJobID + (SELECT dbo.Repeater('0',6)) + @strPPNO
	SET @Tmp = '*' + @TMP + '*'
 	RETURN(@Tmp)
END



GO



--Script_V26_16_Fix12_2_InsertCustomerInNetwork
--ثبت اشتراک همزمان در شبکه
--V26_16_Fix5 برای کلیه ورژن های بعد از 
--93/12/23

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER      Procedure dbo.Insert_Cust  
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
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
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
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Code Bigint out 

)  

as  

Begin Tran  

--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  
if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
	 begin  
	   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  )  --AND (Branch = @Branch )
	   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode   )--AND (Branch = @Branch )
	 end   
else   

	if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId ) <> 0 --AND Branch = @Branch)   
		--Goto ErrHandler   
		BEGIN  -- baraye sabt hamzaman eshterak
			SELECT  @MembershipId =  ISNULL(MAX(MembershipId) ,0) + 1 from  dbo.tCust WHERE Branch = @Branch   --AND (Branch = @Branch )
		END 

Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  Branch = @Branch )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

if @nvcRFID = N''  
  SET @nvcRFID=N'-999'  

insert Into dbo.tCust  
(   
	Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Assansor,   
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
	CarryFee,   
	PaykFee,   
	Distance,   
	Credit,   
	Discount,   
	BuyState,   
	[Description],   
	[Date],   
	[Time],   
	[User],  
	FamilyNo ,  
	Member ,  
	State ,  
	Central ,  
	Branch,  
	nvcRFID,  
	sellprice ,
	EconomicCode ,
	nvcBirthDate ,
	TotalRemainingAmount
	
)  
values  
(   
	@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Assansor,   
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
	@CarryFee,   
	@PaykFee,   
	@Distance,   
	@Credit,   
	@Discount,   
	@BuyState,   
	@Description,   
	@Date,   
	@Time,   
	@User ,  
	@FamilyNo ,  
	@Member ,  
	@State ,  
	@Central ,  
	@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate ,
	@TotalRemainingAmount
	
)  
if @@Error <> 0   
 goto ErrHandler  

--Set @Code = @@Identity  
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
  and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address)  
 , nvcRFID=CAST(Branch AS NVARCHAR(1))+CAST(Code AS NVARCHAR(8))  
  where code=@code  AND Branch = @Branch 


Update [tCust]
Set [Name] = Replace(  [Name] , N'ك' , N'ک'  ) 
Update [tCust]
Set [Name] = Replace(  [Name]  , N'ي' , N'ی' ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Family] = Replace(  [Family]  , N'ي' , N'ی' ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
Update [tCust]
Set WorkName = Replace(  WorkName  , N'ي' , N'ی' ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
Update [tCust]
Set [Address] = Replace(  [Address]  , N'ي' , N'ی' ) 


Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO






SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    VIEW dbo.vw_NotPaidFactors      
AS      
SELECT     tfacm.intSerialNo, tfacm.[No], tfacm.Status, tfacm.Owner, tfacm.Customer, tfacm.DiscountTotal, tfacm.SumPrice, tfacm.CarryFeeTotal,      
  tfacm.Recursive, tfacm.FacPayment, tfacm.InCharge, tfacm.OrderType, tfacm.ServePlace, tfacm.StationID, tfacm.ServiceTotal, tfacm.PackingTotal,      
  tfacm.BascoleNo, tfacm.ShiftNo, tfacm.TableNo, tfacm.[Date], tfacm.[Time], tfacm.[User], tfacm.RegDate, tfacm.Branch, tfacm.Balance, tfacm.AccountYear, tfacm.NvcDescription, tfacm.RefFacM,       


  CASE dbo.tCust.[Name] + ' ' + dbo.tCust.Family WHEN ' ' THEN tCust.WorkName ELSE dbo.tCust.[Name] + ' ' +       
  dbo.tCust.Family END AS [Full Name],       
   dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, dbo.tPer.Job, dbo.tCust.MembershipId AS Code, dbo.tCust.Address, dbo.tCust.Credit      
    ,dbo.tServePlace.[Description] AS ServePlaceName,dbo.tCust.distance,CASE LTRIM(RTRIM(dbo.tFacM.NvcDescription)) WHEN N'پيغام' THEN 1       
  WHEN N'' THEN 1 ELSE -1 END AS intWarn,ISNULL(LTRIM(RTRIM(dbo.tFacM.TempAddress)),'') AS TempAddress,      
  (DATEPART(hour, GETDATE()) * 60 + DATEPART(minute, GETDATE()))-      
  (CAST(SUBSTRING(dbo.tFacM.[Time], 1, 2) AS int) * 60 + CAST(SUBSTRING(dbo.tFacM.[Time], 4, 2) AS int))  AS RemainMinute      
  ,t.DateSend,t.TimeSend      
  ,      
  (DATEPART(hour, GETDATE()) * 60 + DATEPART(minute, GETDATE()))-      
  (CAST(SUBSTRING(t.TimeSend, 1, 2) AS int) * 60 + CAST(SUBSTRING(t.TimeSend, 4, 2) AS int))  AS RemainMinuteSend
    ,ISNULL(LTRIM(RTRIM(dbo.[tCust].Mobile)),'') AS Mobile        
, ISNULL(tfacm.GuestNo , '') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.No) AS TempNo	
    , tshift.Description AS ShiftDescription , dbo.tCust.Tafsili 
	, Tel1 + Tel2 + Tel3 + Tel4 +Mobile AS TelNumber
FROM         dbo.tFacM       
  INNER JOIN dbo.tServePlace ON dbo.tServePlace.intServePlace= dbo.tfacm.ServePlace      
  INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code
  LEFT OUTER JOIN dbo.tPer ON dbo.tFacM.InCharge = dbo.tPer.pPno and dbo.tper.ActDeact=1 --AND dbo.tFacM.Branch = dbo.tPer.Branch 
  LEFT OUTER JOIN dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code --AND (dbo.tFacM.Branch = dbo.tCust.Branch OR dbo.tCust.Branch IS NULL)      

 left outer  JOIN (SELECT MAX(RegDate) AS DateSend,MAX(RegTime) AS TimeSend,intserialno FROM thistory       
        WHERE ActionCode=4 GROUP BY intserialno) t      
 ON [tfacm].[intSerialNo] = t.[intSerialNo]        
 WHERE     (dbo.tFacM.Balance = 0 And Status =2 and Recursive=0)   --for list peik in recived    




GO



ALTER  VIEW dbo.VwTotal_NotDelivers
AS
SELECT     dbo.tCust.MembershipId As Code,dbo.tFacM.intSerialNo, dbo.tFacM.[No], 
                      CASE dbo.tCust.Family + dbo.tCust.[Name] WHEN '' THEN tCust.WorkName ELSE dbo.tCust.[Name] + ' ' + dbo.tCust.Family END AS [Full Name], 
                      dbo.tFacM.SumPrice, dbo.tFacM.[Time], dbo.tCust.Address, dbo.tFacM.[Date], dbo.tfacm.ServePlace, 
                      dbo.tServePlace.[Description] AS ServePlaceName,dbo.tFacM.AccountYear,
		 (DATEPART(hour, GETDATE()) * 60 + DATEPART(minute, GETDATE()))-
		(CAST(SUBSTRING(dbo.tFacM.[Time], 1, 2) AS int) * 60 + CAST(SUBSTRING(dbo.tFacM.[Time], 4, 2) AS int))  AS RemainMinute
		,dbo.tCust.distance,CASE LTRIM(RTRIM(dbo.tFacM.NvcDescription)) WHEN N'پيغام' THEN 1 
		WHEN N'' THEN 1 ELSE -1 END AS intWarn,LTRIM(RTRIM(dbo.tFacM.TempAddress)) AS TempAddress
		, Tel1 + Tel2 + Tel3 + Tel4 +Mobile AS TelNumber
FROM         dbo.tFacM LEFT OUTER JOIN
                      tCust ON dbo.tCust.Branch = dbo.tFacM.Branch AND dbo.tCust.Code = dbo.tFacM.Customer INNER JOIN
                      dbo.tServePlace ON dbo.tServePlace.intServePlace = dbo.tfacm.ServePlace
WHERE     (dbo.tFacM.Incharge IS NULL OR
                      dbo.tFacM.Incharge = '') AND dbo.tFacM.Status = 2 AND dbo.tFacM.facPayment = 0 AND dbo.tFacM.TableNo IS NULL AND dbo.tfacm.Recursive <> 1 AND
                       (dbo.tfacm.ServePlace = 2 OR
                      dbo.tfacm.ServePlace = 4) AND dbo.tFacM.Branch = dbo.Get_Current_Branch()




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   PROCEDURE dbo.GetTotal_Delivers 
	@AccountYear SMALLINT,
	@Job INT
AS
Select Code,intSerialNo,[no],[Full Name],SumPrice,[Time],Address,[Date],ServePlace,ServePlaceName,N'سفارش' as DeliverStatus,
	CAST(RemainMinute/60  AS VARCHAR(4))+':'+CAST(RemainMinute%60  AS VARCHAR(4)) AS RemainTime
	,N'ارسال نشده' as RemainTimesend,N'ارسال نشده' as Timesend
	,distance,intWarn,isnull(TempAddress,'') as TempAddress , '' AS InchargeName , '' AS Incharge
	, ISNULL(VwTotal_NotDelivers.TelNumber , N'') AS TelNumber
	from VwTotal_NotDelivers WHERE AccountYear = @AccountYear
UNION
Select  Code,intSerialNo,[no],[Full Name],SumPrice,[Time],Address,[Date],ServePlace,ServePlaceName,N'ارسال شده' as DeliverStatus,
	CAST(RemainMinute/60  AS VARCHAR(4))+':'+CAST(RemainMinute%60  AS VARCHAR(4)) AS RemainTime
	,CAST(RemainMinutesend/60  AS VARCHAR(4))+':'+CAST(RemainMinutesend%60  AS VARCHAR(4)) AS RemainTimesend,Timesend
	,distance,intWarn,isnull(TempAddress,'')as TempAddress , ISNULL(nvcFirstName , '') + ' ' + ISNULL(nvcSurName, '') AS InchargeName , InCharge
	, ISNULL(vw_NotPaidFactors.TelNumber , N'') AS TelNumber
	from vw_NotPaidFactors Where Job = @Job And Balance = 0
ORDER BY [No] , [Date] ,[Time]



GO
