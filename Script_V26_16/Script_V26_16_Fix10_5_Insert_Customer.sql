

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    Procedure dbo.Insert_Cust  
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
  Goto ErrHandler   

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



Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO
