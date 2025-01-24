

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO

ALTER  PROCEDURE [dbo].[Get_EditedFactors_Print] (
@SystemDate  	NVARCHAR(20),
@SystemDay   	NVARCHAR(20),
@SystemTime  	NVARCHAR(20),
@DateAfter Nvarchar(20) , 
@DateBefore Nvarchar(20)

)
 AS

SELECT    
		      @DateBefore  AS DateBefore, @DateAfter AS DateAfter ,
	   	      @SystemDay + ' ' + @SystemDate +' '+N' ÓÇÚÊ : ' + @SystemTime AS Sysdate  ,
		      dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tFacM.Status, 
                      dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.OrderType, 
                      dbo.tFacM.ServePlace, dbo.tFacM.StationID,  
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName + ' ' +  dbo.tPer.nvcSurName As FullName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description ,
		      dbo.tRepFacEditM.[Time] As Time1 , dbo.tRepFacEditM.SumPrice As Price1
FROM         dbo.tFacM INNER JOIN
                      dbo.tRepFacEditM ON dbo.tFacM.intSerialNo = dbo.tRepFacEditM.intSerialNo and dbo.tFacM.Branch = dbo.tRepFacEditM.Branch INNER JOIN
                      dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
                      dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  --and  dbo.tFacM.Branch = dbo.tShift.Branch
WHERE     (dbo.tFacM.intSerialNo IN
                          (SELECT     dbo.tRepFacEditM.intSerialNo
                             FROM         dbo.tRepFacEditM Where Branch = dbo.Get_Current_Branch()))  
			And dbo.tFacm.[Date] >= @DateAfter And dbo.tFacm.[Date] <= @DateBefore 
			And dbo.tFacm.Status =2


order By dbo.tFacM.intSerialNo desc
GO

--exec dbo.Get_EditedFactors_Print;1 N'93/07/26',N'ÔäÈå',N'01:11',N'93/07/20',N'93/07/26'

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS OFF
GO
ALTER   PROCEDURE [dbo].[Get_RefferedFactors_Print] (
@SystemDate  	NVARCHAR(20),
@SystemDay   	NVARCHAR(20),
@SystemTime  	NVARCHAR(20),
@DateAfter Nvarchar(20) , 
@DateBefore Nvarchar(20)

)
 AS

Select T.* , IsNull(dbo.tHistory.RegTime ,' ') As CreateTime From tHistory  inner Join
(SELECT    
		      @DateBefore  AS DateBefore, @DateAfter AS DateAfter ,
	   	      @SystemDay + ' ' + @SystemDate +' '+N' ÓÇÚÊ : ' + @SystemTime AS Sysdate  ,
                      dbo.tPer.nvcFirstName + ' ' +  dbo.tPer.nvcSurName As FullName, dbo.tFacM.intSerialNo, dbo.tFacM.[No],
                      dbo.tFacM.SumPrice, dbo.tFacM.Recursive, 
                      dbo.tFacM.OrderType, dbo.tFacM.ServePlace, dbo.tFacM.StationId, 
                      dbo.tFacM.ShiftNo, dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User],
                      dbo.tFacM.RegDate, dbo.tShift.Description , dbo.tHistory.RegTime 
FROM         dbo.tFacM INNER JOIN
                      dbo.tHistory ON dbo.tFacM.intSerialNo = dbo.tHistory.intSerialNo INNER JOIN
					  dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID AND dbo.tFacM.[Branch] = dbo.tUser.Branch INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno AND dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
                      dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code --AND  dbo.tFacM.Branch = dbo.tShift.Branch
Where (dbo.tFacM.Recursive = 1) And dbo.tFacm.[Date] >= @DateAfter And dbo.tFacm.[Date] <= @DateBefore
			And dbo.tFacm.Status =2  And dbo.tHistory.ActionCode = 3

 
) T  On T.intserialNo = tHistory.intserialNo 
Where tHistory.actionCode = 1
order By T.intSerialNo desc


GO

