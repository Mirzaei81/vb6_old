ALTER PROCEDURE [dbo].[InsertPersonel]( 
	@PersonnelNumber nvarchar(50),
	@nvcFirstName nvarchar(50),
	@nvcSurName nvarchar(50),
	@Gender bit,
	@IdNumber nvarchar(50),
	@Job int,
	@InsuranceNo nvarchar(50) ,
	@Address nvarchar(300),
	@Tel nvarchar(30),
	@User int , 
	@UserName nvarchar(50) ,
	@Password nvarchar(50) ,
	@intAccessLevel int ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno int out

	)
 AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

set @Time = dbo.SetTimeFormat(getdate())

select @pPno = isnull(max(Ppno),0) + 1 from tper Where Branch = @Branch 
If @pPno < (@Branch * 1000 ) Set @pPno = (@Branch * 1000 )


begin Tran
insert into dbo.tper (
	pPno ,
	PersonnelNumber,
	nvcFirstName,
	nvcSurName,
	Gender ,
	IdNumber,
	Job ,
	InsuranceNo  ,
	Address ,
	Tel ,
	[Date] ,
	[Time] ,			
	[User] ,
	Branch,
	MaxCredit,
	ActDeAct

)
values(
	@pPno ,
	@PersonnelNumber,
	@nvcFirstName,
	@nvcSurName ,
	@Gender ,
	@IdNumber ,
	@Job ,
	@InsuranceNo ,
	@Address ,
	@Tel ,
	@Date,
	@Time ,
	@User ,
	@Branch,
	@MaxCredit,
	@ActDeAct
)
if @@Error <> 0 
		GOTO EventHandler	


--set @pPno=@@identity
DECLARE @UID INT
if @intAccessLevel<>0 and @UserName <> '' and @Password<>''

BEGIN

	select @Uid = isnull(max(Uid),0) + 1 from tUser Where Branch = @Branch 
	If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )

	insert into dbo.tUser 
	(
		[Uid] ,
	 UserName ,
	 [Password] ,
	 intAccessLevel ,
	 pPno ,
	 addUser , 
	 Branch, 
	 CountRePrint, 
	 CountInvoicePrint,
	 CountInvoiceEditable,
	 CountInvoiceRefferable
	)
 values (
	@UID ,					
	@UserName  ,
	@Password  ,
	@intAccessLevel ,
	@pPno , 
	@User ,
	@Branch,
	@CountRePrint,
	@CountInvoicePrint,
	@CountInvoiceEditable,
	@CountInvoiceRefferable
	)

if @@Error <> 0 
		GOTO EventHandler	
--SET @UID = @@IDENTITY		
END	



commit Tran



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1




GO

