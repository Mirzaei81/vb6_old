
SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO


ALTER  PROCEDURE Get_CurrentEditTime
@intserialNo INT ,
@Baranch INT 
 
AS

DECLARE @MinuteUseDiff INT
SELECT @MinuteUseDiff = 
( (CAST(SUBSTRING(dbo.shamsi(GETDATE()), 4, 2) AS INT) - 1 ) * 30
                + CAST(SUBSTRING(dbo.shamsi(GETDATE()), 7, 2) AS INT) ) * 1440
+ ( DATEPART(HOUR, GETDATE()) * 60 + DATEPART(minute, GETDATE()) ) 
 - ( (CAST(SUBSTRING(T.nvcDate, 4, 2) AS INT) - 1) * 30
                + CAST(SUBSTRING(T.nvcDate, 7, 2) AS INT) )  * 1440
           - ( CAST(SUBSTRING(T.nvcTime, 1, 2) AS INT) * 60
                + CAST(SUBSTRING(T.nvcTime, 4, 2) AS INT) ) 
from 
(
SELECT TOP 1 ISNULL(dbo.tRepFacEditM.Time , dbo.tFacM.Time) AS nvcTime , 
	         ISNULL(dbo.tRepFacEditM.RegDate , dbo.tFacM.RegDate) AS nvcDate 
 FROM dbo.tFacM LEFT OUTER JOIN dbo.tRepFacEditM ON dbo.tRepFacEditM.Branch = dbo.tFacM.Branch AND dbo.tRepFacEditM.intSerialNo = dbo.tFacM.intSerialNo
WHERE dbo.tFacM.intSerialNo = @intserialNo AND dbo.tFacM.Branch = @Baranch	
	And Code = (Select MIN(Code) from tRepFacEditM where intSerialNo = @intserialNo AND Branch = @Baranch)
) T

--SET @MinuteUseDiff = 25
SELECT @MinuteUseDiff AS  MinuteUseDiff,
 CASE WHEN @MinuteUseDiff < 0 THEN 0 WHEN @MinuteUseDiff > dbo.[Get_UserEditTime]() THEN 0 ELSE 1 END AS UserDiffTme ,
 CASE WHEN @MinuteUseDiff < 0 THEN 0 WHEN @MinuteUseDiff > dbo.[Get_ManagerEditTime]() THEN 0 ELSE 1 END AS ManagerDiffTme


GO




ALTER    PROCEDURE [dbo].[UpdatePersonel]( 
	@CurrentPPNO 		INT,
	@PersonnelNumber 	NVARCHAR(50),
	@nvcFirstName 		NVARCHAR(50),
	@nvcSurName	 	NVARCHAR(50),
	@Gender 		BIT,
	@IdNumber 		NVARCHAR(50),
	@Job 			INT,
	@InsuranceNo 		NVARCHAR(50) ,
	@Address 		NVARCHAR(300),
	@Tel 			NVARCHAR(30),
	@User 			INT , 
	@UID 			INT ,
	@UserName 		NVARCHAR(50) ,
	@Password 		NVARCHAR(50) ,
	@intAccessLevel 	INT ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno 			INT OUT
	       )
AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

SET @Time= dbo.SetTimeFormat(getdate())

BEGIN TRANSACTION

	UPDATE tPer
		SET PersonnelNumber 	= @PersonnelNumber,
		    nvcFirstName    	= @nvcFirstName,
		    nvcSurName	    	= @nvcSurName,
		    Gender	    	= @Gender,
		    IdNumber       	= @IdNumber,
		    Job		    	= @Job,
		    InsuranceNo     	= @InsuranceNo,
		    Address	    	= @Address,
		    Tel   	    	= @Tel,
		    [Date]	    	= @Date,
		    [Time]	    	= @Time,
		    [User]	    	= @User,
		    MaxCredit		=@MaxCredit,
		    ActDeAct 		=@ActDeAct ,
		    Branch			= @Branch
	WHERE       pPNO = @CurrentPPNO  


	IF @@ERROR <> 0 
		GOTO EventHandler	

	set @pPno = @CurrentPPNO

	IF @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID<>0
		UPDATE tUser
			SET 		UserName       	= @UserName,
	        	   		 	[Password]     	= @Password,
			    		intAccessLevel 	= @intAccessLevel,
			    		pPno           	= @pPno,
			    		addUser        	= @User,
					 CountRePrint		=@CountRePrint,
		  			 CountInvoicePrint	=@CountInvoicePrint,
					 CountInvoiceEditable		=@CountInvoiceEditable,
		  			 CountInvoiceRefferable	=@CountInvoiceRefferable ,
		  			 Branch					= @Branch
			WHERE   UID = @UID    
	else 
		if @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID=0
		BEGIN 
			select @Uid = isnull(max(Uid),0) + 1 from tUser --Where Branch = @Branch   
			If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )  
			insert into dbo.tUser (
						UID ,
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
			) values (	
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
			END 
	IF @@ERROR <> 0 
		GOTO EventHandler	



COMMIT TRANSACTION



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1



GO







