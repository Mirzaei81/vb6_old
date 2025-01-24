

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
