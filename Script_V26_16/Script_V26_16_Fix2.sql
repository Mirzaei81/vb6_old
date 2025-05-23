


--Script_V26_16_Fix2
--اضافه شدن مبلغ بدهی یا طلب اولیه به مشتریان و تامین کنندگان
--گذاشتن دسترسی برای تغییر مبلغ اولیه مشتریان

-- 92/11/05

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
          2
        )
GO
DELETE FROM dbo.tblAcc_DocumentHeader
go 

DELETE FROM tblAcc_Tafsili WHERE TafsiliId > 0
GO

UPDATE dbo.tPer SET Tafsili = NULL 
UPDATE dbo.tCust SET Tafsili = NULL 
UPDATE dbo.tSupplier SET Tafsili = NULL 
GO


ALTER TABLE dbo.tCust
ADD TotalRemainingAmount INT NOT NULL DEFAULT(0) ,
	SanadNo INT NULL 
	
GO

ALTER TABLE dbo.tSupplier
ADD TotalRemainingAmount INT NOT NULL DEFAULT(0) ,
	SanadNo INT NULL 
	
GO

INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 325 , -- intObjectCode - int
          N'ChangeTotalRemainingAmount' , -- ObjectId - nvarchar(50)
          N'تغییر بدهی یا طلب اولیه مشتریان' , -- ObjectName - nvarchar(50)
          N'ChangeTotalRemainingAmount' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          102  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1 ,-- intAccessLevel - int
          325  -- intObjectCode - int
          )
          
GO

UPDATE dbo.tObjects
SET ObjectId = 'frmAccount' ,
	ObjectName = N'سیستم یکپارچه حسابداری' ,
	objectLatinName = 'frmAccount' ,
	intObjectType = 1
WHERE intObjectCode = 343

GO

INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 467 , -- intObjectCode - int
          N'frmDefineAccount' , -- ObjectId - nvarchar(50)
          N'تعریف کد و نام حساب ها' , -- ObjectName - nvarchar(50)
          N'frmDefineAccount' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          470  -- ObjectParent - int
        )
        
GO

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          467  -- intObjectCode - int
          )
          
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tCust_SanadNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_tCust_SanadNo
GO


CREATE PROCEDURE dbo.Update_tCust_SanadNo
(
	@SanadNo	INT,
	@Code int

) 

AS

	UPDATE 	dbo.tCust
		SET 	SanadNo = @SanadNo
	    	          WHERE   Code = @Code 

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tSupplier_SanadNo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_tSupplier_SanadNo
GO


CREATE PROCEDURE dbo.Update_tSupplier_SanadNo
(
	@SanadNo	INT,
	@Code int

) 

AS

	UPDATE 	dbo.tSupplier
		SET 	SanadNo = @SanadNo
	    	          WHERE   Code = @Code 
	    	          
	    	          
GO


ALTER  Procedure dbo.Insert_Supplier  
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
	@Code Bigint out  
)  

as  

Begin Tran  

--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  
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
	Branch  
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
	NULL --@Branch  
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
   

Commit Tran
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code  
--Select @Code


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
	@Updated Bigint out  

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
 TotalRemainingAmount = @TotalRemainingAmount 
Where Code = @Code  -- AND (Branch = @Branch )  


if @@Error <> 0   
 goto ErrHandler  

 UPDATE dbo.tSupplier  
 SET Address = tmpCust.Address  
 FROM dbo.tSupplier  , dbo.tSupplier tmpCust  
 WHERE dbo.tSupplier.MasterCode = tmpCust.Code  
  --and (dbo.tSupplier.[Branch] = tmpCust.[Branch] or dbo.tSupplier.[Branch] is Null)  

update tSupplier set address=dbo.addressedit(address) where code=@code  --AND (Branch = @Branch )


Commit Tran   
Set @Updated = @Code  
return @Updated  

ErrHandler:  
RollBack Tran  
Set @Updated = 0  
return @Updated



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

--Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0)  

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

if @nvcRFID = N''  
  SET @nvcRFID=N'-999'  

insert Into dbo.tCust  
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
	NULL , --@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate ,
	@TotalRemainingAmount
	
)  
if @@Error <> 0   
 goto ErrHandler  

Set @Code = @@Identity  
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
  --and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address)  
 , nvcRFID=CAST(Branch AS NVARCHAR(1))+CAST(Code AS NVARCHAR(8))  
  where code=@code  --AND Branch = @Branch 



Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO




ALTER  Procedure dbo.Update_Cust  
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
	@Code Bigint ,  
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Updated Bigint out  

)  

as  

Begin Tran  
--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  ) --AND (Branch = @Branch ) )  
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode )   --AND (Branch = @Branch ) )  
 end  
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code  ) <> 0  -- AND (Branch = @Branch )    
  Goto ErrHandler   
 else  

  Update dbo.tCust     
   Set MembershipId = @MembershipId   

   Where MasterCode = @Code   --AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tCust  

 Set MembershipId = @MembershipId ,  
	MasterCode  = @MasterCode ,    
	Owner = @Owner ,  
	Name = @Name ,  
	Family = @Family ,  
	Sex = @Sex ,  
	WorkName = @WorkName ,   
	InternalNo = @InternalNo ,  
	Unit = @Unit ,  
	City = @City ,  
	ActKind = @ActKind ,  
	ActDeAct = @ActDeAct ,  
	Prefix = @Prefix ,  
	Assansor = @Assansor ,  
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
	CarryFee = @CarryFee ,  
	PaykFee = @PaykFee ,  
	Distance = @Distance ,  
	Credit = @Credit ,  
	Discount = @Discount ,  
	BuyState = @BuyState ,  
	[Description] = @Description ,  
	[Date] = @Date ,  
	[Time] = @Time ,  
	[User] = @User ,  
	FamilyNo = @FamilyNo ,  
	Member = @Member ,  
	State = @State ,  
	Central = @Central,  
	Sellprice=@Sellprice  ,
	EconomicCode = @EconomicCode ,
	nvcRFID = @nvcRFID ,
	nvcBirthDate = @nvcBirthDate ,
	TotalRemainingAmount = @TotalRemainingAmount
	
Where Code = @Code   --AND (Branch = @Branch )   

if @@Error <> 0   
 goto ErrHandler  


Set @Updated = @Code   
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
 -- and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address) where code=@code  --AND Branch = @Branch 
 


Commit Tran  
return @Updated  

ErrHandler:  
RollBack Tran  
return -1



GO




ALTER  view vw_Customers

as
SELECT     dbo.tCust.Code, dbo.tCust.MembershipId, dbo.tCust.MasterCode, dbo.tCust.Owner, dbo.tCust.Name, dbo.tCust.Family, dbo.tCust.Sex, 
                      dbo.tCust.WorkName, dbo.tCust.State, dbo.tCust.City,dbo.tCust.ActKind, dbo.tCust.ActDeAct, 
                      dbo.tCust.Prefix, dbo.tCust.Assansor, dbo.tCust.Address, dbo.tCust.PostalCode, dbo.tCust.Tel1, dbo.tCust.Tel2, dbo.tCust.Tel3, 
                      dbo.tCust.Tel4, dbo.tCust.Mobile, dbo.tCust.Fax, dbo.tCust.Email, dbo.tCust.CarryFee, dbo.tCust.PaykFee, dbo.tCust.Distance, dbo.tCust.Discount, 
                      dbo.tCust.BuyState, dbo.tCust.Credit, dbo.tCust.Description, dbo.tCust.[Date], dbo.tCust.[Time], dbo.tCust.[User], dbo.tCust.Unit, dbo.tCust.InternalNo, 
                      dbo.tCust.Flour, ISNULL(SUM(T.sumPrice), 0) AS Price , IsNull(T2.Bestankar,0) As Bestankar ,  CASE 
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

			End as FullName , case
			
			WHEN (dbo.tCust.MasterCode is null )
				Then
					dbo.tCust.Address
			WHEN (dbo.tCust.MasterCode is not null )
				Then
					tCust_1.Address + N' طبقه ' + isnull(dbo.tCust.Flour , '') + N' واحد ' + isnull(dbo.tCust.Unit , '')
	--				dbo.tCust.Address + ' ' + tCust_1.Address + N' طبقه ' + isnull(dbo.tCust.Flour , '') + N' واحد ' + isnull(dbo.tCust.Unit , '')
			end as FullAddress  , tCust.FamilyNo , tCust.Member , tCust.Central , tCust.SellPrice , tCust.[Branch] , tcust.Tafsili
FROM         dbo.tCust LEFT OUTER JOIN  dbo.tCust tCust_1  on dbo.tCust.MasterCode = tCust_1.Code and dbo.tCust.Branch = tCust_1.Branch
			  
			LEFT OUTER JOIN
                          (SELECT     Sum(IsNull(Bestankar,0)) As Bestankar , Code_Bes
                             FROM         dbo.tblAcc_Recieved
                             WHERE     dbo.tblAcc_Recieved.RecieveType = 3 And AccountYear = dbo.Get_AccountYear() Group By Code_Bes) T2  ON dbo.tCust.Code = T2.Code_Bes  
                      
			 LEFT OUTER JOIN
                          (SELECT    Case Status When 2 then  sumPrice When 5 then  -Sumprice end as sumprice, Customer  ,  Branch
                             FROM         dbo.tFacM
                             WHERE     dbo.tFacm.Balance = 0 And dbo.tFacm.FacPayment = 1 /* and dbo.tFacm.Branch = dbo.Get_Current_Branch()*/ ) T ON dbo.tCust.Code = T.Customer --and dbo.tCust.Branch = T.Branch
--WHERE [tCust].[Branch] = dbo.[Get_Current_Branch]()

GROUP BY dbo.tCust.Code, dbo.tCust.MembershipId, dbo.tCust.MasterCode, dbo.tCust.Owner, dbo.tCust.Name, dbo.tCust.Family, dbo.tCust.Sex, 
          dbo.tCust.WorkName, dbo.tCust.State, dbo.tCust.City, dbo.tCust.ActKind, dbo.tCust.ActDeAct, 
          dbo.tCust.Prefix, dbo.tCust.Assansor, dbo.tCust.Address, dbo.tCust.PostalCode, dbo.tCust.Tel1, dbo.tCust.Tel2, dbo.tCust.Tel3, 
          dbo.tCust.Tel4, dbo.tCust.Mobile, dbo.tCust.Fax, dbo.tCust.Email, dbo.tCust.CarryFee, dbo.tCust.PaykFee, dbo.tCust.Distance, dbo.tCust.Discount, 
          dbo.tCust.BuyState, dbo.tCust.Credit, dbo.tCust.Description, dbo.tCust.[Date], dbo.tCust.[Time], dbo.tCust.[User], dbo.tCust.Unit, dbo.tCust.InternalNo, 
          dbo.tCust.Flour , tCust_1.WorkName , tCust_1.Address , tCust.FamilyNo , tCust.Member , T2.Bestankar , tCust.Central , tCust.SellPrice , tCust.[Branch]
			, dbo.tCust.Tafsili


GO




ALTER   VIEW dbo.vw_FacM_Per
AS
SELECT  dbo.tFacM.StationID,
		dbo.tFacM.RegDate, 
		ISNULL(dbo.tFacM.InCharge, 0) AS InCharge, 
		ISNULL(dbo.tFacM.TableNo, 0) AS TableNo, 
		dbo.tFacM.[Time], 
        dbo.tPer.nvcFirstName, 
		dbo.tPer.nvcSurName, 
		dbo.tFacM.[No], 
		dbo.tFacM.Status, 
		dbo.tFacM.[User], 
		dbo.tFacM.intSerialNo, 
        dbo.tShift.Description AS ShiftDescription, 
		dbo.tShift.Code AS ShiftNo, 
		dbo.tFacM.Balance, 
		dbo.tFacM.FacPayment, 
		dbo.tFacM.ServePlace , 
		dbo.tFacM.AccountYear
		, CASE DeliveryPer.job WHEN 3 THEN ISNULL(DeliveryPer.nvcFirstName,'-') +' '+ISNULL(DeliveryPer.nvcSurName,'-') ELSE N'--' END AS DeliveryFullName 
		,dbo.tFacM.Branch
		, dbo.tFacM.BitHavaleResid
		,dbo.tFacM.transferAccounting 
		, tfacm.BitLock , tfacm.GuestNo , tfacm.TempNo , Refrence_Acc
FROM    dbo.tFacM 
		INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID 
							--AND dbo.tFacM.Branch = dbo.tUser.Branch 
		INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno 
							--AND dbo.tUser.Branch = dbo.tPer.Branch 
		INNER JOIN dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code 
							--AND dbo.tFacM.Branch = dbo.tShift.Branch
		LEFT OUTER JOIN dbo.tPer AS DeliveryPer ON tFacM.InCharge = DeliveryPer.pPno 
							--AND tFacM.Branch = DeliveryPer.Branch 
                      
--WHERE     (dbo.tFacM.Branch = dbo.Get_Current_Branch()) 





GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS OFF
GO


ALTER  PROCEDURE dbo.Update_tSupplier_tafsili
(
	@tafsili	INT,
	@Code int

) 

AS

	UPDATE 	dbo.tSupplier
		SET 	Tafsili = @Tafsili
	    	          WHERE   Code = @Code -- And  (Branch = dbo.Get_Current_Branch() or Branch Is Null)
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_SaleSummary]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_SaleSummary]
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
CREATE    PROCEDURE [dbo].[Get_SaleSummary]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
 SELECT (SUM(tvw.SumPrice) - SUM(PackingTotal)-SUM(CarryFeeTotal)-SUM(ServiceTotal)-SUM(DutyTotal)-SUM(TaxTotal)+SUM(DiscountTotal)) AS SumPriceTotal ,
        tvw.[Date] ,
        Branch ,
        SUM(DiscountTotal) AS Sumdiscount ,
        SUM(PackingTotal) AS SumPacking ,
        SUM(CarryFeeTotal) AS SumCarryFee ,
        SUM(ServiceTotal) AS SumService ,
        SUM(DutyTotal) AS DutyTotal ,
        SUM(TaxTotal) AS TaxTotal ,
        0 AS tafsili
 FROM   ( SELECT    dbo.tFacM.Branch ,
                    dbo.tFacM.[No] ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacM.[Time] ,
                    dbo.tFacM.[User] ,
                    CarryFeeTotal ,
                    DiscountTotal ,
                    StationID ,
                    ServiceTotal ,
                    PackingTotal ,
                    TaxTotal ,
                    DutyTotal ,
                    FacPayment ,
                    Balance ,
                    --( tfacd.Amount * tfacd.Feeunit ) AS SumPrice
                    dbo.tFacM.SumPrice
          FROM      dbo.tFacM
                    --INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                    --                    AND tfacm.Branch = tfacD.Branch
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch
 ORDER BY tvw.[Date] 
 
 
END

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Get_AccountDocument]
    (
      @Branch INT ,
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @Code INT ,
      @Uid INT 
    )
AS 
    IF ( @Code = 1 ) 
            BEGIN
                SELECT  tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
                        tPer.Tafsili AS PersonTafsili 
                       --, ISNULL(SUM(ISNULL(tFacCash.intAmount , 0)), 0)  AS sp
                       ,CASE WHEN dbo.tFacM.Status =2 THEN SUM(dbo.tFacM.SumPrice) ELSE -1 * SUM(dbo.tFacM.SumPrice) END AS sp
                FROM    tFacM
                        INNER JOIN tUser ON tUser.UID = tFacM.[User]
                        INNER JOIN tPer ON tUser.pPno = tPer.pPno
						INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
                WHERE    tFacM.Branch = @Branch 
                        AND tFacM.Recursive = 0 
                        AND (tFacM.Status = 2 OR tFacM.Status = 5 )
                        AND dbo.tFacM.transferAccounting=0  
						AND (tfacm.[User] = @Uid OR @Uid = 0)
						AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
                GROUP BY tFacM.[Date] ,tFacM.Status ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
                        tPer.Tafsili 
                HAVING  tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,tPer.Tafsili
            END

    --IF ( @Code = 5 ) 
    --        BEGIN
    --            SELECT  tFacM.[Date] ,
    --                    tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
    --                    tPer.Tafsili AS PersonTafsili ,
    --                    ISNULL(SUM(tFacCard.intAmount), 0) AS sp
    --            FROM    tUser
    --                    INNER JOIN tPer ON tUser.pPno = tPer.pPno
    --                    INNER JOIN tFacM ON tUser.UID = tFacM.[User]
    --                                       -- AND tUser.Branch = tFacM.Branch
    --                    INNER JOIN tFacCard ON tFacM.Branch = tFacCard.Branch
    --                                           AND tFacM.intSerialNo = tFacCard.intSerialNo
    --            WHERE    tFacM.Branch = @Branch 
    --                    AND  tFacM.Recursive = 0 
    --                    AND ( tFacM.Status = @Status
    --                          OR tFacM.Status = 8
    --                        )AND dbo.tFacM.transferAccounting=0
    --            GROUP BY tFacM.[Date] ,
    --                    tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
    --                    tPer.Tafsili
    --            HAVING   tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
    --            ORDER BY tFacM.[Date] ,
    --                    tPer.Tafsili
    --        END



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER PROC Update_transferAccounting
(
  @Branch INT ,
  @DateBefore NVARCHAR(8) ,
  @DateAfter NVARCHAR(8),
  @SanadNo INT ,
  @Uid INT 
)

AS
	UPDATE dbo.tFacM
	SET dbo.tFacM.transferAccounting=1	,
		dbo.tFacM.BitLock = 1 ,
		dbo.tFacM.Refrence_Acc = @SanadNo
	WHERE tfacm.Branch = @Branch
		AND tfacm.[Date] >= @DateBefore
		AND tfacm.[Date] <= @DateAfter
		AND [Recursive] = 0
		AND transferAccounting = 0
		AND (Status = 2 OR Status = 5) 
		--AND (Customer < 0 OR Customer IS NULL OR  (Customer > 0 AND Credit = 0))  --لازم نیست چون فاکتور مشتریان قبلا سند حسابداری خورده 
		AND (tfacm.[User] = @Uid OR @Uid = 0)

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[Get_SaleReturnSummary]
    (
	@Branch int,
	@DateBefore NVARCHAR(8) ,
	@DateAfter NVARCHAR(8) ,
	@Uid INT 
    )

AS 

BEGIN
 SELECT (SUM(tvw.SumPrice) - SUM(PackingTotal)-SUM(CarryFeeTotal)-SUM(ServiceTotal)-SUM(DutyTotal)-SUM(TaxTotal)+SUM(DiscountTotal)) AS SumPriceTotal ,
        tvw.[Date] ,
        Branch ,
        SUM(DiscountTotal) AS Sumdiscount ,
        SUM(PackingTotal) AS SumPacking ,
        SUM(CarryFeeTotal) AS SumCarryFee ,
        SUM(ServiceTotal) AS SumService ,
        SUM(DutyTotal) AS DutyTotal ,
        SUM(TaxTotal) AS TaxTotal ,
        0 AS tafsili
 FROM   ( SELECT    dbo.tFacM.Branch ,
                    dbo.tFacM.[No] ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacM.[Time] ,
                    dbo.tFacM.[User] ,
                    CarryFeeTotal ,
                    DiscountTotal ,
                    StationID ,
                    ServiceTotal ,
                    PackingTotal ,
                    TaxTotal ,
                    DutyTotal ,
                    FacPayment ,
                    Balance ,
                    --( tfacd.Amount * tfacd.Feeunit ) AS SumPrice
                    dbo.tFacM.SumPrice
          FROM      dbo.tFacM
					INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 5
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
                    AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
        ) tvw
 GROUP BY tvw.[Date] ,  Branch
 ORDER BY tvw.[Date] 
 
 
END


GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER    view [dbo].[vw_Suppliers]
as
SELECT  dbo.tSupplier.Code ,
        dbo.tSupplier.MembershipId ,
        dbo.tSupplier.MasterCode ,
        dbo.tSupplier.Owner ,
        dbo.tSupplier.Name ,
        dbo.tSupplier.Family ,
        dbo.tSupplier.Sex ,
        dbo.tSupplier.WorkName ,
        tSupplier.State ,
        dbo.tSupplier.City ,
        dbo.tSupplier.ActKind ,
        dbo.tSupplier.ActDeAct ,
        dbo.tSupplier.Prefix ,
        dbo.tSupplier.Address ,
        dbo.tSupplier.PostalCode ,
        dbo.tSupplier.Tel1 ,
        dbo.tSupplier.Tel2 ,
        dbo.tSupplier.Tel3 ,
        dbo.tSupplier.Tel4 ,
        dbo.tSupplier.Mobile ,
        dbo.tSupplier.Fax ,
        dbo.tSupplier.Email ,
        dbo.tSupplier.Discount ,
        dbo.tSupplier.Description ,
        dbo.tSupplier.[Date] ,
        dbo.tSupplier.[Time] ,
        dbo.tSupplier.[User] ,
        dbo.tSupplier.Unit ,
        dbo.tSupplier.InternalNo ,
        dbo.tSupplier.Flour ,
        ISNULL(SUM(T.sumPrice), 0) AS Price ,
        CASE WHEN ( dbo.tSupplier.MasterCode IS NULL
                    AND dbo.tSupplier.WorkName = ''
                    AND dbo.tSupplier.Name <> ''
                  ) THEN dbo.tSupplier.Family + ' ' + dbo.tSupplier.Name
             WHEN ( dbo.tSupplier.MasterCode IS NULL
                    AND dbo.tSupplier.WorkName = ''
                    AND dbo.tSupplier.Name = ''
                  ) THEN dbo.tSupplier.Family
             WHEN ( dbo.tSupplier.MasterCode IS NULL
                    AND dbo.tSupplier.WorkName <> ''
                  ) THEN dbo.tSupplier.WorkName
             WHEN ( dbo.tSupplier.MasterCode IS NOT NULL
                    AND tSupplier_1.WorkName <> ''
                    AND dbo.tSupplier.Name <> ''
                  )
             THEN tSupplier_1.WorkName + '_' + dbo.tSupplier.Family + ' '
                  + dbo.tSupplier.Name
             WHEN ( dbo.tSupplier.MasterCode IS NOT NULL
                    AND tSupplier_1.WorkName <> ''
                    AND dbo.tSupplier.Name = ''
                  ) THEN tSupplier_1.WorkName + '_' + dbo.tSupplier.Family
        END AS FullName ,
        CASE WHEN ( dbo.tSupplier.MasterCode IS NULL )
             THEN dbo.tSupplier.Address
             WHEN ( dbo.tSupplier.MasterCode IS NOT NULL )
             THEN dbo.tSupplier.Address + ' ' + tSupplier_1.Address
                  + N' طبقه ' + ISNULL(dbo.tSupplier.Flour, '') + N' واحد '
                  + ISNULL(dbo.tSupplier.Unit, '')
        END AS FullAddress ,
        dbo.tSupplier.Branch , ISNULL(dbo.tSupplier.Tafsili ,0) AS Tafsili
FROM    dbo.tSupplier
        LEFT OUTER JOIN dbo.tSupplier tSupplier_1 ON dbo.tSupplier.MasterCode = tSupplier_1.Code
                                                     --AND dbo.tSupplier.Branch = tSupplier_1.Branch
        LEFT OUTER JOIN ( SELECT    sumPrice ,
                                    Owner ,
                                    Branch
                          FROM      dbo.tFacM
                          WHERE     dbo.tFacm.Facpayment = 0
                                    -- AND dbo.tFacm.Branch = dbo.Get_Current_Branch()
                        ) T ON dbo.tSupplier.Code = T.Owner
                               --AND dbo.tSupplier.Branch = T.Branch
GROUP BY dbo.tSupplier.Code ,
        dbo.tSupplier.MembershipId ,
        dbo.tSupplier.MasterCode ,
        dbo.tSupplier.Owner ,
        dbo.tSupplier.Name ,
        dbo.tSupplier.Family ,
        dbo.tSupplier.Sex ,
        dbo.tSupplier.WorkName ,
        dbo.tSupplier.State ,
        dbo.tSupplier.City ,
        dbo.tSupplier.ActKind ,
        dbo.tSupplier.ActDeAct ,
        dbo.tSupplier.Prefix ,
        dbo.tSupplier.Address ,
        dbo.tSupplier.PostalCode ,
        dbo.tSupplier.Tel1 ,
        dbo.tSupplier.Tel2 ,
        dbo.tSupplier.Tel3 ,
        dbo.tSupplier.Tel4 ,
        dbo.tSupplier.Mobile ,
        dbo.tSupplier.Fax ,
        dbo.tSupplier.Email ,
        dbo.tSupplier.Discount ,
        dbo.tSupplier.Description ,
        dbo.tSupplier.[Date] ,
        dbo.tSupplier.[Time] ,
        dbo.tSupplier.[User] ,
        dbo.tSupplier.Unit ,
        dbo.tSupplier.InternalNo ,
        dbo.tSupplier.Flour ,
        tSupplier_1.WorkName ,
        tSupplier_1.Address ,
        dbo.tSupplier.Branch ,
        dbo.tSupplier.Tafsili


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   Proc Get_All_Supplier
 @Branch INT 
as

Select * from dbo.vw_Get_Supplier where code > 0 --AND ( Branch = @Branch OR @Branch IS NULL )
Order By Code

GO




ALTER View [dbo].[vw_DeliveryFactor]   AS

SELECT     tfacm.intSerialNo, tfacm.[No], tfacm.Status, tfacm.Owner, tfacm.Customer, tfacm.DiscountTotal, tfacm.SumPrice, tfacm.CarryFeeTotal,
	 tfacm.Recursive, tfacm.FacPayment, tfacm.InCharge, tfacm.OrderType, tfacm.ServePlace, tfacm.StationID, tfacm.ServiceTotal, tfacm.PackingTotal,
 	tfacm.BascoleNo, tfacm.ShiftNo, tfacm.TableNo, tfacm.[Date], tfacm.[Time], tfacm.[User], tfacm.RegDate, tfacm.Branch, tfacm.Balance, tfacm.AccountYear, tfacm.NvcDescription, tfacm.RefFacM,
	 CASE  dbo.tCust.[Name] + ' ' + dbo.tCust.Family  WHEN ' '  THEN tCust.WorkName
			ELSE dbo.tCust.[Name] + ' ' + dbo.tCust.Family 
			END AS [Full Name],  dbo.tPer.nvcFirstName , dbo.tPer.nvcSurName ,dbo.tPer.Job, dbo.tCust.MembershipId As Code, dbo.tCust.Address ,  dbo.tCust.Credit
    , ISNULL(tfacm.GuestNo , '') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.No) AS TempNo	
    , tshift.Description AS ShiftDescription , ISNULL(dbo.tCust.Mobile , '') AS Mobile , dbo.tCust.Tafsili
	FROM         dbo.tFacM left outer JOIN
      dbo.tPer ON dbo.tFacM.InCharge = dbo.tPer.pPno and dbo.tFacM.Branch = dbo.tPer.Branch Inner JOIN
      dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code and (dbo.tFacM.Branch = dbo.tCust.Branch or dbo.tCust.Branch Is Null)
	  INNER JOIN dbo.tShift ON dbo.tFacM.Branch = dbo.tShift.Branch
                       AND dbo.tfacm.ShiftNo = dbo.tShift.Code
	WHERE  	dbo.tFacM.FacPayment = 0  and dbo.tFacM.Recursive = 0  and dbo.tFacM.Balance = 0  and dbo.tFacM.Branch  = dbo.Get_Current_Branch()  and
		dbo.tFacm.intSerialNo in (Select distinct dbo.tFacD.intSerialNo from dbo.tFacD where  dbo.tFacD.Branch  = dbo.Get_Current_Branch() )



GO


ALTER  PROCEDURE dbo.GetCustomersInfo  

@AccountYear SmallInt   , @Branch INT 

AS  
 SELECT   dbo.tCust.MembershipId As Code, dbo.tFacM.intSerialNo ,dbo.tFacM.[No] ,  
  CASE  dbo.tCust.Family + dbo.tCust.[Name] WHEN ''  THEN tCust.WorkName  
                       ELSE dbo.tCust.[Name] + ' ' + dbo.tCust.Family  
  END AS [Full Name],  
  dbo.tFacM.SumPrice , dbo.tFacM.[Time] , dbo.tCust.Address ,dbo.tFacM.[Date] ,dbo.tfacm.ServePlace ,dbo.tServePlace.[Description] AS ServePlaceName  
  ,dbo.tCust.distance,CASE LTRIM(RTRIM(dbo.tFacM.NvcDescription)) WHEN N'پيغام' THEN 1   
  WHEN N'' THEN 1 ELSE -1 END AS intWarn,dbo.tFacM.NvcDescription,ISNULL(LTRIM(RTRIM(dbo.tFacM.TempAddress)),'') AS TempAddress  
  ,ISNULL(LTRIM(RTRIM(dbo.[tCust].Mobile)),'') AS Mobile
, ISNULL(tfacm.GuestNo , '') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.No) AS TempNo	
    , tshift.Description AS ShiftDescription , dbo.tCust.Tafsili
FROM   dbo.tFacM  
	Left Outer JOIN tCust ON dbo.tCust.Code = dbo.tFacM.Customer   
	INNER JOIN dbo.tServePlace ON dbo.tServePlace.intServePlace= dbo.tfacm.ServePlace  
    INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code

 WHERE   ( dbo.tFacM.Incharge IS NULL OR dbo.tFacM.Incharge = '' ) AND dbo.tFacM.Status = 2  
		AND dbo.tFacM.facPayment = 0    
		and dbo.tFacM.TableNo is null   
		--And dbo.tfacm.Customer > 0  
		And dbo.tfacm.Recursive <> 1  
		And (dbo.tfacm.ServePlace = 2 OR dbo.tfacm.ServePlace = 4)  
		And AccountYear = @AccountYear   
		AND dbo.tFacM.Branch =  @Branch  
 ORDER BY dbo.tFacM.[No] , dbo.tFacM.[Date] ,dbo.tFacM.[Time]  




GO






ALTER   VIEW dbo.vw_NotPaidFactors      
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






ALTER  Proc Get_vw_NotPaidFactors_By_Job (@Job int , @AccountYear Smallint)  
as  
 Select vw_NotPaidFactors.intSerialNo,vw_NotPaidFactors.No,vw_NotPaidFactors.Status,  
        vw_NotPaidFactors.Owner,vw_NotPaidFactors.Customer,vw_NotPaidFactors.DiscountTotal,  
        vw_NotPaidFactors.SumPrice,vw_NotPaidFactors.CarryFeeTotal,vw_NotPaidFactors.Recursive,  
        vw_NotPaidFactors.FacPayment,vw_NotPaidFactors.InCharge,vw_NotPaidFactors.OrderType,  
        vw_NotPaidFactors.ServePlace,vw_NotPaidFactors.StationID,vw_NotPaidFactors.ServiceTotal,  
        vw_NotPaidFactors.PackingTotal,vw_NotPaidFactors.BascoleNo,vw_NotPaidFactors.ShiftNo,  
        vw_NotPaidFactors.TableNo,t.Date,t.TIME,vw_NotPaidFactors.RegDate,vw_NotPaidFactors.[USER],  
           vw_NotPaidFactors.Branch,vw_NotPaidFactors.Balance,vw_NotPaidFactors.ServePlaceName,  
        vw_NotPaidFactors.AccountYear,vw_NotPaidFactors.NvcDescription,vw_NotPaidFactors.RefFacM,  
        vw_NotPaidFactors.[Full NAME],vw_NotPaidFactors.nvcFirstName,vw_NotPaidFactors.nvcSurName,  
        vw_NotPaidFactors.Job,vw_NotPaidFactors.Code,vw_NotPaidFactors.Address,vw_NotPaidFactors.Credit,  
        vw_NotPaidFactors.distance,vw_NotPaidFactors.intWarn,vw_NotPaidFactors.RemainMinute,  
        vw_NotPaidFactors.TempAddress,vw_NotPaidFactors.mobile ,  
        vw_NotPaidFactors.GuestNo , vw_NotPaidFactors.TempNo , vw_NotPaidFactors.ShiftDescription 
        , vw_NotPaidFactors.Tafsili
  from vw_NotPaidFactors  
  INNER  JOIN (SELECT MAX(RegDate) AS Date,MAX(RegTime) AS Time,intserialno FROM thistory   
        WHERE ActionCode=4 GROUP BY intserialno) AS t   
		ON [vw_NotPaidFactors].[intSerialNo] = t.[intSerialNo]   
   Where  Balance = 0 And FacPayment = 0 And AccountYear = @AccountYear
			AND vw_NotPaidFactors.InCharge > 0  --Job = @Job And 




GO




ALTER    Proc Get_vw_NotPaidFactors_By_Job_InCharge (@Job int , @InCharge Int, @AccountYear Smallint)  
as  
 Select vw_NotPaidFactors.intSerialNo,vw_NotPaidFactors.No,vw_NotPaidFactors.Status,  
        vw_NotPaidFactors.Owner,vw_NotPaidFactors.Customer,vw_NotPaidFactors.DiscountTotal,  
        vw_NotPaidFactors.SumPrice,vw_NotPaidFactors.CarryFeeTotal,vw_NotPaidFactors.Recursive,  
        vw_NotPaidFactors.FacPayment,vw_NotPaidFactors.InCharge,vw_NotPaidFactors.OrderType,  
        vw_NotPaidFactors.ServePlace,vw_NotPaidFactors.StationID,vw_NotPaidFactors.ServiceTotal,  
        vw_NotPaidFactors.PackingTotal,vw_NotPaidFactors.BascoleNo,vw_NotPaidFactors.ShiftNo,  
        vw_NotPaidFactors.TableNo,t.Date,t.TIME,vw_NotPaidFactors.RegDate,vw_NotPaidFactors.[USER],  
           vw_NotPaidFactors.Branch,vw_NotPaidFactors.Balance,vw_NotPaidFactors.ServePlaceName,  
        vw_NotPaidFactors.AccountYear,vw_NotPaidFactors.NvcDescription,vw_NotPaidFactors.RefFacM,  
        vw_NotPaidFactors.[Full NAME],vw_NotPaidFactors.nvcFirstName,vw_NotPaidFactors.nvcSurName,  
        vw_NotPaidFactors.Job,vw_NotPaidFactors.Code,vw_NotPaidFactors.Address,vw_NotPaidFactors.Credit,  
        vw_NotPaidFactors.distance,vw_NotPaidFactors.intWarn,vw_NotPaidFactors.RemainMinute,  
		vw_NotPaidFactors.TempAddress,vw_NotPaidFactors.[Mobile] ,  
         vw_NotPaidFactors.GuestNo , vw_NotPaidFactors.TempNo , vw_NotPaidFactors.ShiftDescription
         , vw_NotPaidFactors.Tafsili
 from vw_NotPaidFactors  
  INNER  JOIN (SELECT MAX(RegDate) AS Date,MAX(RegTime) AS Time,intserialno FROM thistory   
        WHERE ActionCode=4 GROUP BY intserialno) AS t   
  ON [vw_NotPaidFactors].[intSerialNo] = t.[intSerialNo]   
   Where Job = @Job And InCharge = @InCharge And Balance = 0 And FacPayment = 0 And AccountYear = @AccountYear
GO



