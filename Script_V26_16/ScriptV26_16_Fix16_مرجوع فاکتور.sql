
--ScriptV26_16_Fix16_ãÑÌæÚ ÝÇ˜ÊæÑ.sql
--95/04/02

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE Update_tFacM_Recursive
(
@No  Bigint,
@Status int,
@Recursive int,
@Uid int,
@Balance Bit,
@FacPayment Bit ,
@AccountYear Smallint = NULL ,
@Branch INT 
)

AS
Declare @TableNo int
DECLARE @intTableUsedNo INT      
IF @AccountYear Is Null 
	SET @AccountYear = dbo.Get_AccounYear()

DECLARE @intSerialNo BIGINT

--DECLARE @Branch INT
--	SET @Branch = dbo.Get_Current_Branch()

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch and AccountYear = @AccountYear)

UPDATE tFacM
     SET Recursive= @Recursive
         WHERE tFacM.intSerialNo = @intserialNo And  Branch = @Branch 


DECLARE @intserialNo2 BIGINT
If @Status = 6 OR (@Status = 2 AND dbo.AutoHavale() = 1)
BEGIN 
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear)  
	IF @intSerialNo2 > 0
		UPDATE tFacM
			 SET Recursive= @Recursive
				 WHERE tFacM.intSerialNo = @intserialNo2 And  Branch = @Branch 
END 

If @Recursive = 1 
Begin

UPDATE tFacM
     SET FacPayment = 0 , Balance = 0
         WHERE tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear

SET @TableNo = ISNULL((Select TableNo From tfacm   Where  tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear) , 0)
  UPDATE tTable
       SET Empty = 1 
           WHERE dbo.tTable.[No] = @TableNo
	If dbo.Get_TableMonitoring() = 1 AND @TableNo > 0		---Table Monitoring
	Begin
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.bitIsValid = 0      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	


Exec DeleteFactorChildren @intSerialNo , @Branch

UPDATE dbo.tblAcc_Recieved SET Bestankar = 0 WHERE intSerialNo = @intSerialNo And  Branch = @Branch  

End

If @Recursive = 0

Begin
   Update tFacm 
       SET FacPayment = @FacPayment , Balance = @Balance
         WHERE tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear

	SET @TableNo = ISNULL((Select TableNo From tfacm   Where  tFacM.[No]=@No   AND Status = @Status And  Branch = @Branch and AccountYear = @AccountYear) , 0)
    UPDATE tTable
       SET Empty = 0 
           WHERE dbo.tTable.[No] = @TableNo
	If dbo.Get_TableMonitoring() = 1 AND @TableNo > 0		---Table Monitoring
	Begin
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@Branch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.bitIsValid = 1      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
 	END	

	IF @Balance = 1
	BEGIN 
	DELETE FROM tFacCash WHERE intSerialNo = @intSerialNo AND [Branch] = @Branch
	INSERT INTO tFacCash (intSerialNo, intAmount ,branch)
		SELECT @intSerialNo AS
	 intSerialNo, Sumprice,@Branch From tFacM  WHERE tFacM.[No]=@No   AND Status = 2 And  Branch = @Branch and AccountYear = @AccountYear

	END 
End

--Declare @Monitor1 Bit
--Declare @Monitor2 Bit

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())


--If @Monitor1 > 0 
--  exec Notify_to_Clients
--Else If @Monitor2 > 0 
--  exec Notify_to_Clients

If @Recursive = 0
   Exec InsertHistory  @No, @Status , @Uid , 8 ,@AccountYear , @Branch
Else if @Recursive = 1
   Exec InsertHistory  @No, @Status , @Uid , 3 ,@AccountYear , @Branch 

---------------------------------------Mojodi Control Online---------------------------------------------------------

Exec DeleteMojodiCalculate @Status , @intserialNo , @Recursive ,@AccountYear , @Branch
If (@Status = 6 OR (@Status = 2 AND dbo.AutoHavale() = 1) ) AND @intserialNo2 > 0
	EXEC DeleteMojodiCalculate 7, @intSerialNo2 , @Recursive, @AccountYear , @Branch

--------------------------------------------------------------------------------------------------------------------------------------

---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @intSerialNo , 3

--------------------------------------------------------------------------------------------------------------------------------------
GO
