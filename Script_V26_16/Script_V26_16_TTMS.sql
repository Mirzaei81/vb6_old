
--Script_V26_16_TTMS
--اضافه شدن فرم محاسبه مالیات و عوارض بر ارزش افزوده
--گذاشتن دسترسی برای فرم ارزش افزوده
-- اضافه شدن کد اقتصادی و شناسه (کد) ملی به تامین کنندگان
--برای ارسال گزارش دقیق به دارایی
--94/02/05


ALTER TABLE dbo.tSupplier
ADD EconomicCode NVARCHAR(20) NULL ,
    NationalCode NVARCHAR(20) NULL 
GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 330 , -- intObjectCode - int
          N'frmTTMS' , -- ObjectId - nvarchar(50)
          N'گزارشات ارزش افزوده' , -- ObjectName - nvarchar(50)
          N'frmTTMS' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          102  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          330  -- intObjectCode - int
          )
          
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER     Procedure dbo.Insert_Supplier  
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
	@State int ,  
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
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
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,
	@TotalRemainingAmount INT ,   
	@EconomicCode nVarChar(20) = NULL ,
	@NationalCode nVarChar(20) = NULL ,
	@Code Bigint out  
	
)  

as  

Begin Tran  

DECLARE @Branch INT 
 Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode ) --AND (Branch = @Branch ) )  
    end   
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId) <> 0 -- AND Branch = @Branch ) <> 0   
  Goto ErrHandler   

--Set @Code = (Select  isnull(Max(Code),0) + 1 from tSupplier where code > 0)  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

insert Into dbo.tSupplier  
(   
	--Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	State ,  
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
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
	Discount,   
	[Description],   
	[Date],   
	[Time],   
	[User], 
	TotalRemainingAmount , 
	Branch ,
	EconomicCode ,
	NationalCode
	 
)  
values  
(   
	--@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@State,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
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
	@Discount,   
	@Description,   
	@Date,   
	@Time,   
	@User , 
	@TotalRemainingAmount , 
	@Branch ,
	@EconomicCode ,
	@NationalCode 
)  
if @@Error <> 0   
 goto ErrHandler  

Set @Code = @@Identity  
 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] )   

update tSupplier set address=dbo.addressedit(address) where code=@code  -- AND Branch = @Branch
   

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 

Commit Tran
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code  
--Select @Code


GO
SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  Procedure dbo.Update_Supplier  
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
	@State int ,   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
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
	@Discount Float,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint , 
	@TotalRemainingAmount  INT , 
	@EconomicCode nVarChar(20) = NULL ,
	@NationalCode nVarChar(20) = NULL ,  
	@Updated Bigint OUT  

)  

as  

Begin Tran  
--IF @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tSupplier where  Code = @MasterCode )--  AND (Branch = @Branch ) )  
    end  
else   

 if (Select top 1 isnull(Code , 0) from tSupplier where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code ) <> 0 --  AND (Branch = @Branch ) ) <> 0    
  Goto ErrHandler   
 else  

  Update dbo.tSupplier     
   Set MembershipId = @MembershipId   

  Where MasterCode = @Code   --AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tSupplier  

 Set MembershipId = @MembershipId ,  
 MasterCode  = @MasterCode ,    
 Owner = @Owner ,  
 Name = @Name ,  
 Family = @Family ,  
 Sex = @Sex ,  
 WorkName = @WorkName ,   
 InternalNo = @InternalNo ,  
 Unit = @Unit ,  
 State = @State ,  
 City = @City ,  
 ActKind = @ActKind ,  
 ActDeAct = @ActDeAct ,  
 Prefix = @Prefix ,  
 Address = @Address ,  
 PostalCode = @PostalCode ,  
 Tel1 = @Tel1 ,  
 Tel2 = @Tel2 ,  
 Tel3 = @Tel3 ,  
 Tel4 = @Tel4 ,  
 Mobile = @Mobile ,  
 Fax = @Fax ,  
 Email = @Email ,  
 Flour = @Flour ,  
 Discount = @Discount ,  
 [Description] = @Description ,  
 [Date] = @Date ,  
 [Time] = @Time ,  
 [User] = @User ,
 TotalRemainingAmount = @TotalRemainingAmount ,
 EconomicCode = @EconomicCode ,
 NationalCode = @NationalCode
Where Code = @Code  -- AND (Branch = @Branch )  


if @@Error <> 0   
 goto ErrHandler  

 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] or dbo.tSupplier.[Branch] is Null)  

update tSupplier set address=dbo.addressedit(address) where code=@code  --AND (Branch = @Branch )

	Update dbo.tSupplier
	Set [Name] = Replace(  [Name]  , N'ك' , N'ک' ) 
	Update tSupplier
	Set [Name] = Replace(  [Name]  , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Family] = Replace(  [Family]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Family] = Replace(  [Family] , N'ي', N'ی'   ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName  , N'ك', N'ک'  ) 
	Update tSupplier
	Set WorkName = Replace(  WorkName , N'ي', N'ی'   ) 
	Update tSupplier
	Set [Address] = Replace(  [Address]  , N'ك', N'ک'  ) 
	Update tSupplier
	Set [Address] = Replace(  [Address] , N'ي', N'ی'   ) 


Commit Tran   
Set @Updated = @Code  
return @Updated  

ErrHandler:  
RollBack Tran  
Set @Updated = 0  
return @Updated



GO


