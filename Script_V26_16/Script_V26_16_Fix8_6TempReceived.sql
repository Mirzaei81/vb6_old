

--Script_V26_16_Fix8_TempReceived.sql
--ثبت رسید موقت در فاکتور خرید و امکان ویرایش آن
-- و تبدیل آن به فاکتور خرید هنگام ثبت مجدد در مود مشاهده
--93/08/08

UPDATE dbo.tStatusType SET Active = 1 WHERE intStatusNo = 9
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   Proc Get_tFacM_By_No_Status(@No bigint ,@Status int ,@AccountYear Smallint , @Branch Int )
as
If @Status <> dbo.Get_Numeric_Status(N'TempRecieved') 
	Select * --,isnull(t.Amount,0) AS TaxTotal 
	FROM  	dbo.tFacM  --LEFT OUTER JOIN 
-- (SELECT Amount,intserialno,branch FROM  [tFactorAdditionalServices] where [tFactorAdditionalServices].[intServiceNo]=2	) t
-- 	ON [tFacM].[Branch] = t.[Branch] AND [tFacM].[intSerialNo] = t.[intSerialNo]
	where [No] = @No and Status = @Status And AccountYear = @AccountYear And tfacm.Branch =  @Branch
		
Else if @Status = dbo.Get_Numeric_Status(N'TempRecieved')

	Select * 
	FROM  	dbo.tFacM  --LEFT OUTER JOIN 
	where [No] = @No and Status = @Status And AccountYear = @AccountYear And tfacm.Branch =  @Branch
	--Select * from 
	--dbo.tHavaleM where [No] = @No And AccountYear = @AccountYear And DesBranch =  @Branch and State = 0



GO

