

--Script_V26_16_CashReceivedSummary.sql
-- For Versions V26_16_Fix8 الی V26_16_Fix12


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Get_All_tblPub_Pos] AS
select * from [tblPub_Pos] inner join dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblPub_Pos.intBank
 ORDER BY PosId

GO


