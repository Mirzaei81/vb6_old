

--فلگ برای رسیدهای موقت
--رسید موقت فقط یکبار به رسید دائم تبدیل شود
--دسترسی برای دائمی کردن رسید موقت
--93/09/17


IF COL_LENGTH('tFacM','BitTempReceived') IS NULL
BEGIN
	ALTER TABLE tFacM
	ADD BitTempReceived BIT NULL 
END

GO


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 327 , -- intObjectCode - int
          N'frmSaveTempReceived' , -- ObjectId - nvarchar(50)
          N' دائمی کردن رسید موقت' , -- ObjectName - nvarchar(50)
          N'frmSaveTempReceived' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          126  -- ObjectParent - int
        )
GO
        
        
INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          327  -- intObjectCode - int
          )
GO
        


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Get_All_Factors]
    (
      @Status INT ,
      @User INT ,
      @AccountYear SMALLINT ,
      @Branch INT,
	  @DateAfter Nvarchar(8) , 
      @DateBefore Nvarchar(8)
    )
AS 
    DECLARE @AccessLevel INT
    DECLARE @LastfacmNo INT
    DECLARE @Date NVARCHAR(50)
    SET @Date = dbo.[Get_ShamsiDate_For_Current_Shift](GETDATE())
    DECLARE @ShiftNo INT
    SET @ShiftNo = dbo.Get_Shift(dbo.SetTimeFormat(GETDATE()))
--    PRINT @ShiftNo

    SET @AccessLevel = ISNULL(( SELECT MIN(AccessLevel)
                                FROM    ( SELECT TOP 100 PERCENT
                                                    CASE WHEN [ObjectId] LIKE N'viewallstationsfactors'
                                                         THEN 1
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                         THEN 2
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                         THEN 3
                                                         ELSE 4
                                                    END AS AccessLevel
                                          FROM      dbo.tUser
                                                    INNER JOIN dbo.tAccess_Object ON dbo.tUser.intAccessLevel = dbo.tAccess_Object.intAccessLevel
                                                    INNER JOIN dbo.tObjects ON dbo.tAccess_Object.intObjectCode = dbo.tObjects.intObjectCode
                                          WHERE     --tObjects.ObjectId LIKE 'viewallstationsfactors' AND
                                                    UID = @User
                                                    --AND dbo.tUser.Branch = @Branch
                                                    AND ( [dbo].[tObjects].[ObjectId] LIKE N'viewallstationsfactors'
                                                    OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                    OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                        )
                                          ORDER BY  [dbo].[tObjects].[intObjectCode] DESC
                                        ) T1
                              ), 4)

    DECLARE @intAccessLevel INT
    SELECT  @intAccessLevel = intAccessLevel
    FROM    [dbo].[tUser]
    WHERE   uid = @User
            --AND [Branch] = @Branch
    IF @intAccessLevel = 1 
        SET @AccessLevel = @intAccessLevel


    SET @LastfacmNo = ( SELECT  ISNULL(MAX([NO]), 0) + 1
                        FROM    tFacM
                        WHERE   Status = @Status
                                AND Branch = @Branch
                                AND dbo.tFacM.AccountYear = @AccountYear
                      )   
--    IF @LastfacmNo < 1000 
--        SET @LastfacmNo = 0 
--    ELSE 
--        IF @LastfacmNo > 1000 
--            SET @LastfacmNo = @LastfacmNo - 1000

    SELECT  dbo.tFacM.intSerialNo, [No],tfacm.[Date],tfacm.[Time], SumPrice, Balance, Recursive, ServiceTotal, CarryFeeTotal, DiscountTotal,isnull( NvcDescription ,N'') as NvcDescription ,
            dbo.tPer.nvcFirstName ,
            dbo.tPer.nvcSurName ,
            ISNULL(tcust.WorkName + dbo.tCust.Family , '') + ISNULL(tSupplier.WorkName + dbo.tSupplier.Family , '') AS CustomerName
            , ISNULL(tfacm.GuestNo ,'') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.[No]) AS TempNo
            , tshift.Description AS ShiftDescription , ISNULL(tfacm.BitTempReceived , 0) AS BitTempReceived
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno
            INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code
            LEFT OUTER JOIN dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code 
            LEFT OUTER JOIN dbo.tSupplier ON dbo.tFacM.Owner = dbo.tSupplier.Code 
    WHERE   ( @AccessLevel = 1
              OR ( @AccessLevel = 2
                   AND [dbo].[tFacM].[Date] = @Date
                 )
              OR ( @AccessLevel = 3
                   AND [ShiftNo] = @ShiftNo
                   AND [dbo].[tFacM].[Date] = @Date
                 )
	      OR ( @AccessLevel = 4
	           AND dbo.tFacM.[ShiftNo] = @ShiftNo
	           AND dbo.tFacM.[Date] = @Date
	           AND dbo.tFacM.[User] = @User
	         )
            )
            AND dbo.tFacM.Status = @Status
           -- AND dbo.tFacM.[No] > @LastfacmNo
            AND dbo.tFacM.AccountYear = @AccountYear
			AND  dbo.tFacm.[Date] >= @DateAfter 
			And dbo.tFacm.[Date] <= @DateBefore
			AND dbo.tFacM.Branch = @Branch
    ORDER BY No DESC



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   PROCEDURE [dbo].[Get_Define_Factors]
    (
      @Status INT ,
      @User INT ,
      @No BIGINT ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS 
--    DECLARE @Branch INT
--    SELECT  @Branch = [dbo].[Get_Current_Branch]()
    DECLARE @AccessLevel INT
    DECLARE @ShiftNo INT
    DECLARE @Date AS NVARCHAR(10)
    SET @ShiftNo = dbo.Get_Shift(dbo.SetTimeFormat(GETDATE()))
    SET @Date = [dbo].[Get_ShamsiDate_For_Current_Shift](GETDATE())

    SET @AccessLevel = ISNULL(( SELECT MIN(AccessLevel)
                                FROM    ( SELECT TOP 100 PERCENT
                                                    CASE WHEN [ObjectId] LIKE N'viewallstationsfactors'
                                                         THEN 1
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                         THEN 2
                                                         WHEN [ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                         THEN 3
                                                         ELSE 4
                                                    END AS AccessLevel
                                          FROM      dbo.tUser
                                                    INNER JOIN dbo.tAccess_Object ON dbo.tUser.intAccessLevel = dbo.tAccess_Object.intAccessLevel
                                                    INNER JOIN dbo.tObjects ON dbo.tAccess_Object.intObjectCode = dbo.tObjects.intObjectCode
                                          WHERE     --tObjects.ObjectId LIKE 'viewallstationsfactors' AND
                                                    UID = @User
                                                    --AND dbo.tUser.Branch = @Branch
                                                    AND ( [dbo].[tObjects].[ObjectId] LIKE N'viewallstationsfactors'
                                                          OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentDayInvoices'
                                                          OR [dbo].[tObjects].[ObjectId] LIKE N'ViewAllCurrentShiftInvoices'
                                                        )
                                          ORDER BY  [dbo].[tObjects].[intObjectCode] DESC
                                        ) T1
                              ), 4)
    
    DECLARE @intAccessLevel INT
    SELECT  @intAccessLevel = intAccessLevel
    FROM    [dbo].[tUser]
    WHERE   uid = @User
           -- AND [Branch] = @Branch
    IF @intAccessLevel = 1 
        SET @AccessLevel = @intAccessLevel

    SELECT  dbo.tFacM.* ,
            dbo.tPer.nvcFirstName ,
            dbo.tPer.nvcSurName ,
            ISNULL(tcust.WorkName + dbo.tCust.Family , '') + ISNULL(tSupplier.WorkName + dbo.tSupplier.Family , '') AS CustomerName
            ,  ISNULL(tfacm.GuestNo ,'') AS GuestNo , ISNULL(tfacm.TempNo , tfacm.[No]) AS TempNo
            , tshift.Description AS ShiftDescription , ISNULL(tfacm.BitTempReceived ,0) AS BitTempReceived
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno
            INNER JOIN dbo.tShift ON dbo.tfacm.ShiftNo = dbo.tShift.Code
            LEFT OUTER JOIN dbo.tCust ON dbo.tFacM.Customer = dbo.tCust.Code 
            LEFT OUTER JOIN dbo.tSupplier ON dbo.tFacM.Owner = dbo.tSupplier.Code 
    WHERE   ( @AccessLevel = 1
              OR ( @AccessLevel = 2
                   AND [dbo].[tFacM].[Date] = @Date
                 )
              OR ( @AccessLevel = 3
                   AND [ShiftNo] = @ShiftNo
                   AND [dbo].[tFacM].[Date] = @Date
                 )
	      OR ( @AccessLevel = 4
	           AND dbo.tFacM.[ShiftNo] = @ShiftNo
	           AND dbo.tFacM.[Date] = @Date
	           AND dbo.tFacM.[User] = @User
	         )
            )
            AND dbo.tFacM.Status = @Status
            AND dbo.tFacM.[No] = @No
            AND AccountYear = @AccountYear
            AND dbo.tFacM.Branch = @Branch
--===============================================




GO



ALTER    VIEW dbo.vw_FacM_Per
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
		, tfacm.BitLock , tfacm.GuestNo , tfacm.TempNo , Refrence_Acc , ISNULL(BitTempReceived ,0) AS BitTempReceived
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




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_BitTempReceived]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_BitTempReceived
GO


CREATE PROCEDURE [dbo].Update_BitTempReceived (
	@intSerialNo BIGINT  ,
	@Branch INT 
	)

AS

	UPDATE dbo.tFacM
		SET BitTempReceived = 1 WHERE intSerialNo = @intSerialNo AND Branch = @Branch
		
		
GO

--SELECT * FROM dbo.tFacM ORDER BY intSerialNo DESC

