

--این اسکریپت فقط یک رکورد برای اصلاحی ها با زمان و مبلغ اولی و آخری بر می گرداند
--Script_V26_16_Fix10_1_EditedFich
--93/10/21


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO

ALTER  PROCEDURE [dbo].[Get_EditedFactors_Print] (
@SystemDate  	NVARCHAR(20),
@SystemDay   	NVARCHAR(20),
@SystemTime  	NVARCHAR(20),
@DateAfter Nvarchar(20) , 
@DateBefore Nvarchar(20)

)
 AS

SELECT    
		      @DateBefore  AS DateBefore, @DateAfter AS DateAfter ,
	   	      @SystemDay + ' ' + @SystemDate +' '+N' ساعت : ' + @SystemTime AS Sysdate  ,
		      dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tFacM.Status, 
                      dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.OrderType, 
                      dbo.tFacM.ServePlace, dbo.tFacM.StationID,  
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName + ' ' +  dbo.tPer.nvcSurName As FullName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description ,
		      ISNULL(T.Time , 0) AS Time1 , ISNULL(T.SumPrice , 0) AS Price1 , ISNULL(T.intSerialNo ,0) AS MinCode  -- , dbo.tRepFacEditM.SumPrice As Price1
FROM         dbo.tFacM  INNER JOIN          
			 dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
             dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
             dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  and  dbo.tFacM.Branch = dbo.tShift.Branch
				LEFT OUTER JOIN (SELECT code , Branch , intSerialNo , SumPrice , Time FROM dbo.tRepFacEditM )T 
						ON T.Branch = dbo.tFacM.Branch AND T.intSerialNo = tFacM.intSerialNo AND T.Code = (Select Min(Code) FROM dbo.tRepFacEditM WHERE T.Branch = dbo.tRepFacEditM.Branch AND T.intSerialNo = dbo.tRepFacEditM.intSerialNo) 									
WHERE      ISNULL(T.intSerialNo ,0) > 0  AND
			 dbo.tFacm.[Date] >= @DateAfter And dbo.tFacm.[Date] <= @DateBefore 
			And dbo.tFacm.Status =2


order By dbo.tFacM.intSerialNo desc
GO





SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE dbo.Insert_CustomerFast


	@MembershipId BIGINT  ,     
	@Name nVarChar(50),   
	@Family nVarChar(50),    
	@Address nvarchar(150),  
	@Tel1 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Description nVarChar(200), 
	@User int ,   
	@Code Bigint out 

as  

 BEGIN TRAN  

	declare @MasterCode int
	set @MasterCode=null

	declare @Owner int    
	set @Owner=0

	declare @Sex int
	set @Sex=1 

	declare @WorkName nVarChar(50)
	set @WorkName=N''

	declare @InternalNo nVarChar(50)
	set @InternalNo=N''
 
	declare @Unit nVarChar(50)
	set @Unit=N''
  
	declare @City int
	set @City=1
	
	declare @ActKind int
	set @ActKind=1
  
	declare @ActDeAct bit
	set @ActDeAct=1
	
	declare @Prefix int
	set @Prefix=1
	  
	declare @Assansor bit   
	set @Assansor=0

	declare @PostalCode nVarChar(50)   
	set @PostalCode=N''

	declare @Tel2 nVarChar(50) 
	set @Tel2=N'' 
	declare @Tel3 nVarChar(50) 
	set @Tel3=N'' 
	declare @Tel4 nVarChar(50) 
	set @Tel4=N'' 

	declare @Fax nVarChar(50)
	set @Fax=N''  

	declare @Email nVarChar(50)
	set @Email=N''
  
	declare @Flour nVarChar(50)
	set @Flour=N''

	declare @CarryFee Float
	set  @CarryFee=0
  
	declare @PaykFee Float
	set  @PaykFee=0
  
	declare @Distance int
	set @Distance=1
   
	declare @Credit Float
	set @Credit=0
   
	declare @Discount Float
	set  @Discount=0
 
	declare @BuyState int
	set @BuyState=15
   
	declare @FamilyNo int 
	set  @FamilyNo=0 
	declare @Member Bit 
	set @Member=1
  
	declare @State int 
	set @State=1
  
	declare @Central BIT 
	set   @Central=1
	
	declare @Sellprice smallint
	set   @Sellprice=1

	declare @EconomicCode NVARCHAR(20) 
	set @EconomicCode=N''

	declare @nvcRFID NVARCHAR(20)
	set @nvcRFID=N''

	declare @nvcBirthDate NVARCHAR(10)
	set @nvcBirthDate=N''


Declare @Branch Int  
Set @Branch = dbo.Get_Current_Branch()  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  
 
Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  ( Branch = @Branch  Or Branch Is NULL ) )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

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
	nvcBirthDate
	
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
	@WorkName ,   
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
	@nvcBirthDate
	
)  

if @@Error <> 0   
 goto ErrHandler  

--SET @Code = @@IDENTITY
--SET @Code = 200


Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code


GO
