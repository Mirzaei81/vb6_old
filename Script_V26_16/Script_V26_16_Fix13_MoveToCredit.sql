
--Script_V26_16_Fix13_MoveToCredit.sql
--انتقال به حساب مشتریان اعتباری 
--مستقیما از روی فرم دریافت
--با زدن کلید اعتباری و چاپ
--94/05/31
 

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE Update_tfacm_Balance  
(
@No Bigint,
@Status int,
@Uid  int,
@AccountYear Smallint = NULL ,
@ds NVARCHAR(4000) = NULL ,-- For Ppc
@Branch INT  
)
AS
IF @AccountYear IS NULL
	SET @AccountYear = dbo.Get_AccountYear()
--DECLARE @Branch INT
--	SET @Branch = dbo.Get_Current_Branch()

BEGIN TRAN 

Declare @TableNo int
Declare @SumPrice BigInt
DECLARE @CountTableInUse int
SET @SumPrice = (SELECT tFacM.SumPrice FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch and AccountYear = @AccountYear)

DECLARE @IntSerialNo Bigint

SET @IntSerialNo = (Select IntSerialNo From tfacm Where [No] = @No  And Status = @Status and Branch =  @Branch and AccountYear = @AccountYear)

set @TableNo = (select dbo.tfacm.TableNo  from tfacm   Where [No] = @No And Status = @Status  And Branch = @Branch and AccountYear = @AccountYear ) 
SET @CountTableInUse=(SELECT COUNT(*)FROM tfacm WHERE dbo.tfacm.TableNo=@TableNo AND Status = @Status  And Branch = @Branch and AccountYear = @AccountYear AND tfacm.[Recursive]=0 AND tfacm.[Balance]=0)
If  @TableNo >0 
begin
	IF @CountTableInUse >= 1
		begin
		UPDATE tTable

		SET Empty=1 
		WHERE dbo.tTable.[No]=@TableNo   AND Branch = @Branch
		END 
		If dbo.Get_TableMonitoring() = 1 AND @CountTableInUse >= 1		---Table Monitoring
		Begin
		DECLARE @intTableUsedNo INT      
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.nvcEndTime=  dbo.SetTimeFormat(getdate())      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	
END 
	IF LEN(@ds) > 0
	   Update tfacm
			 set Balance = 1 , FacPayment = 1 , [User] = @Uid --, BitLock = 1
				 Where [No] = @No And Status = @Status  And Branch = @Branch and AccountYear = @AccountYear
	ELSE	--Move To Credit Cust 
	   Update tfacm 
		 set Balance = 0 , FacPayment = 1 , [User] = @Uid --, BitLock = 1
			 Where [No] = @No And Status = @Status  And Branch = @Branch and AccountYear = @AccountYear

    DECLARE @Date AS NVARCHAR(10)
    SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())
	DECLARE @FichDate AS NVARCHAR(10)
	SET @FichDate = (SELECT [Date] FROM tfacm Where [No] = @No And Status = @Status  And Branch = @Branch and AccountYear = @AccountYear)
	
	IF (@Status =  1 OR @Status = 2 )  
		BEGIN 
		--IF @Date = @FichDate
			DECLARE @Result INT     
			exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds ,  @Branch  , 0  , 1 , @Result OUT      
			IF @@ERROR <>0  
			 GoTo EventHandler  

		--ELSE 
		--	BEGIN 
		--	DECLARE @NewTime NVARCHAR(5)  
		--	SELECT  @NewTime = dbo.[SetTimeFormat](GETDATE())  
		--	DECLARE @RegDate NVARCHAR(20)  
		--	SELECT  @RegDate =   [dbo].[shamsi](GETDATE())


		--	INSERT  INTO dbo.[tblAcc_Recieved]
  --                  ( Code , [No] ,
  --                    [List] ,
  --                    [Date] ,
  --                    [RegDate] ,
  --                    [RegTime] ,
  --                    [UID] ,
  --                    [Description] ,
  --                    [Bestankar] ,
  --                    [Branch] ,
  --                    [RecieveType] ,
  --                    [Code_Bes] ,
  --                    [intSerialNo] ,
  --                    [AccountYear]
  --                  )
  --                  SELECT  ISNULL(MAX([tblAcc_Recieved].Code), 0) + 1 ,
		--					ISNULL(MAX([tblAcc_Recieved].[No]), 0) + 1 ,
  --                          1 ,
  --                          @Date ,
  --                          @RegDate ,
  --                          @NewTime ,
  --                          @Uid ,
  --                          N'دريافت بابت فاكتور ' + CAST( [tFacM].[No] AS NVARCHAR(7)) ,
  --                          [dbo].[tFacM].[SumPrice] ,
  --                          @Branch ,
  --                          3 , --5
  --                          [dbo].[tFacM].[Customer] ,
  --                          [dbo].[tFacM].[intSerialNo] ,
  --                          [dbo].[Get_AccountYear]()
  --                  FROM    [dbo].[tFacM]
		--			LEFT OUTER JOIN dbo.tblAcc_Recieved ON dbo.tFacM.Branch = dbo.tblAcc_Recieved.Branch
  --                  WHERE   [dbo].[tFacM].intSerialNo = @IntSerialNo
  --                  GROUP BY [dbo].[tFacM].[Date] ,
  --                          [dbo].[tFacM].[SumPrice] ,
  --                          [dbo].[tFacM].[intSerialNo] ,
		--					[dbo].[tFacM].[Customer] ,
		--					[dbo].[tFacM].[No]
		--		END 
			END 				


    Exec InsertHistory  @No , @Status , @Uid , 5  , @AccountYear , @Branch

Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @Branch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @Branch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @Branch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @Branch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @Branch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @Branch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @Branch

COMMIT TRAN

RETURN 

EventHandler:  
    ROLLBACK TRAN  

    RETURN 

GO
