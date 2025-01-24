
SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

--«’·«Õ Õ–› —”Ìœ Â‰ê«„ Õ–› ÕÊ«·Â 
--«’·«Õ „—ÃÊ⁄ —”Ìœ Â‰ê«„ „—ÃÊ⁄ ÕÊ«·Â 
--91/03/01

ALTER  PROCEDURE dbo.Delete_tFacmd	
(
	@No     Bigint ,
	@Status	int ,
	@AccountYear	Smallint ,
	@Branch INT ,
	@Result INT Out
)
As
DECLARE @intSerialNo BIGINT
DECLARE @intSerialNo2 BIGINT

Begin Tran

--DECLARE @Branch INT
-- SET @Branch = dbo.Get_Current_Branch()

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear)
IF @Status = 6  
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE intSerialNo = @intSerialNo)  


	Exec DeleteMojodiCalculate @Status , @intserialNo  , 1 , @AccountYear , @Branch
	IF @@ERROR <>0
		GoTo EventHandler

	Delete From dbo.tHistory Where intSerialNo = @intSerialNo and Branch =  @Branch
	Delete From dbo.tfacd  Where intSerialNo = @intSerialNo and Branch =  @Branch
	Delete From dbo.tfacd2  Where intSerialNo = @intSerialNo and Branch =  @Branch
	Delete From dbo.tRepfaceditm  Where intSerialNo = @intSerialNo and Branch =  @Branch
	Exec DeleteFactorChildren @intSerialNo , @Branch

    IF @Status = 6 AND @intSerialNo2 > 0
        BEGIN
			EXEC DeleteMojodiCalculate 7, @intSerialNo2 , 1, @AccountYear , @Branch
            DELETE  FROM dbo.[tFacM]
            WHERE   intSerialNo = @intSerialNo2
                    AND Branch = @Branch 
        END
    IF @@ERROR <> 0 
        GOTO EventHandler

	Delete From dbo.tfacm  Where intSerialNo = @intSerialNo and Branch =  @Branch
    IF @@ERROR <>0
	        GoTo EventHandler


COMMIT TRANSACTION
Set @Result = 1
Return @Result


EventHandler:
    ROLLBACK TRAN
    Set @Result = 0
    RETURN @Result




GO


