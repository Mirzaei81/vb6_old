
--کسر مرجوعی ها از وجه نقد و کارت 
--93/10/20

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
--New

ALTER  VIEW [dbo].[VwStationSaleSummery]
AS 
    SELECT  dbo.tFacM.[No] ,
            dbo.tFacM.[Date] ,
            dbo.tFacM.[Time] ,
            dbo.tFacM.[User] ,
            CASE WHEN tfacm.Recursive = 0 THEN SumPrice ELSE 0 END AS SumPrice ,
            CASE WHEN tfacm.Recursive = 0 THEN CarryFeeTotal ELSE 0 END AS CarryFeeTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN DiscountTotal ELSE 0 END AS DiscountTotal ,
            StationID ,
            CASE WHEN tfacm.Recursive = 0 THEN ServiceTotal ELSE 0 END AS ServiceTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN PackingTotal ELSE 0 END AS PackingTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN TaxTotal ELSE 0 END AS TaxTotal ,
            CASE WHEN tfacm.Recursive = 0 THEN DutyTotal ELSE 0 END AS DutyTotal ,
	        CASE WHEN tfacm.Recursive = 0 THEN dbo.tfacm.[RoundDiscount] ELSE 0 END AS RoundDiscount,
            FacPayment ,
            Balance ,
            dbo.tfacm.Customer ,
            CASE dbo.tCust.[Name] + SPACE(3) + dbo.tCust.family
              WHEN NULL THEN dbo.tCust.WorkName
              WHEN '' THEN dbo.tCust.WorkName
              ELSE dbo.tCust.[Name] + '  ' + dbo.tCust.family
            END AS CustomerName ,
            dbo.tCust.tafsili AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            CASE WHEN Recursive = 1 THEN 0
				WHEN ISNULL(Incharge, 0) = 0 THEN 0
                ELSE CASE ISNULL(TableNo, 0)
                     WHEN 0 THEN SumPrice
                     ELSE 0
                   END
            END AS CarrierSumPrice ,
            CASE WHEN Recursive = 1 THEN 0
				WHEN ISNULL(Incharge, 0) = 0 THEN 0
				ELSE CASE ISNULL(TableNo, 0)
                     WHEN 0 THEN 0
                     ELSE SumPrice
                   END
				END AS GarsonSumPrice ,
            CASE  WHEN Recursive = 1 THEN 0
				  WHEN FacPayment = 0 
                   THEN CASE Balance
                            WHEN 0 THEN CASE ISNULL(Incharge, 0)
                                          WHEN 0 THEN 0
                                          ELSE CASE ISNULL(TableNo, 0)
                                                 WHEN 0 THEN SumPrice
                                                 ELSE 0
                                               END
                                        END
                            ELSE 0
                          END
              ELSE 0
            END AS CarrierDebit ,
            CASE WHEN Recursive = 1 THEN 0
				 WHEN FacPayment = 0 
					THEN CASE ISNULL(Incharge, 0)
								WHEN 0 THEN 0
								ELSE CASE ISNULL(TableNo, 0)
									   WHEN 0 THEN 0
									   ELSE SumPrice
									 END
							  END
				  ELSE 0
            END AS GarsonDebit ,
--            CASE Balance
--              WHEN 0 THEN CASE FacPayment
--                            WHEN 0 THEN 0
--                            ELSE SumPrice
--                          END
--              ELSE 0
--            END AS CustomerDebit ,
			CASE WHEN Recursive = 1 THEN 0
				WHEN Balance = 0 THEN 
				CASE WHEN (Facpayment = 1 or (Incharge is NULL AND serveplace <> 2 AND serveplace <> 16))THEN 
					SumPrice - (ISNULL(Resived.Received , 0) +ISNULL(CardReceived.CardReceived , 0)+ISNULL(PreReceived2.PreReceived2 , 0))
					ELSE 0
					END 
				ELSE 0
			END AS CustomerDebit ,
            CASE WHEN Recursive = 1 THEN 0 
              WHEN Balance = 0 
					THEN CASE FacPayment
                            WHEN 0 THEN CASE ISNULL(Incharge, 0)
                                          WHEN 0 THEN SumPrice
                                          ELSE 0
                                        END
                            ELSE 0
                          END
              ELSE   0
            END AS UnBalanceFich ,
            dbo.tFacM.Branch ,
            0 AS Payment ,
            ISNULL(Resived.Received , 0) AS Recieved ,
            tper.ppno ,
            tfacm.status ,
            ISNULL(CardReceived.CardReceived , 0) AS CardReceived ,
            ISnull(tFactorAdditionalServices.amount , 0)  AS TipAmount ,
			0 AS ManualRecieved , 0 AS OrderPrice , 0 AS OrderReceived ,
			CASE WHEN tfacm.Recursive = 1 THEN SumPrice ELSE 0 END AS SumRecursive 
			
    FROM    dbo.tFacM
        INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
        INNER JOIN tCUst ON tfacM.Customer = tcust.code
       -- LEFT OUTER JOIN [tblAcc_Recieved] ON tfacm.[OrderRefrence] = [tblAcc_Recieved].[intSerialNo] 
        LEFT OUTER JOIN [tFactorAdditionalServices] ON [tFactorAdditionalServices].[Branch] = [tFacM].[Branch] AND [tFactorAdditionalServices].[intSerialNo] = [tFacM].[intSerialNo]
			AND tFactorAdditionalServices.intServiceNo = 3 AND dbo.tFacM.Recursive = 0
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS Received,intSerialNo , Branch FROM  [dbo].[tFacCash] 
			GROUP BY intSerialNo , Branch) AS Resived ON Resived.Branch = tfacm.Branch AND  Resived.intSerialNo = dbo.tFacM.intSerialNo  AND dbo.tFacM.Recursive = 0
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS CardReceived,intSerialNo , Branch FROM  [dbo].[tFacCard] 
			GROUP BY intSerialNo , Branch) AS CardReceived ON CardReceived.Branch = tfacm.Branch AND  CardReceived.intSerialNo = dbo.tFacM.intSerialNo  AND dbo.tFacM.Recursive = 0
	    LEFT outer JOIN (SELECT ISNULL(SUM(Bestankar),0) AS PreReceived2 ,intSerialNo , Branch FROM  [dbo].[tblAcc_Recieved] 
			GROUP BY intSerialNo , Branch) AS PreReceived2 ON PreReceived2.Branch = tfacm.Branch AND  PreReceived2.intSerialNo = dbo.tFacM.intSerialNo  AND dbo.tFacM.Recursive = 0

    WHERE  Status = 2

    UNION
    SELECT  dbo.tFacM.[No] ,
            dbo.tFacM.[Date] ,
            dbo.tFacM.[Time] ,
            dbo.tFacM.[User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
	        0 AS [RoundDiscount],
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tFacM.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    SumPrice AS OrderPrice ,
	    0 AS OrderReceived ,
	    0 AS SumRecursive
    FROM    dbo.tFacM 
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
            INNER JOIN tCust ON tfacM.Customer = tcust.code
    WHERE Recursive = 0 AND Status = 10
    UNION
    SELECT  tblAcc_Cash.Code AS [No] ,
            tblAcc_Cash.[Date] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Cash.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Cash.Branch AS Branch ,
            Bestankar AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived ,
	    0 AS SumRecursive
    FROM    tblAcc_Cash
            INNER JOIN dbo.tUser ON dbo.tblAcc_Cash.[UID] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,--list
            tFacM.[Date] AS [Date] ,  --            tblAcc_Recieved.[Date] AS [Date] , --Date From tFacm
            RegTime AS [Time] ,
            tblAcc_Recieved.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved.Branch AS Branch ,
            0 AS Payment ,
            Bestankar AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS OrderPrice ,
  	    0 AS OrderReceived ,
	    0 AS SumRecursive

    FROM    tblAcc_Recieved 
		INNER JOIN dbo.tFacM ON dbo.tblAcc_Recieved.Branch = dbo.tFacM.Branch AND dbo.tblAcc_Recieved.intSerialNo = dbo.tFacM.intSerialNo
		INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
		INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
	WHERE tblAcc_Recieved.intSerialNo IS NOT NULL AND Status = 2
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,
            tblAcc_Recieved.[Date] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Recieved.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS OrderPrice ,
  	    Bestankar AS OrderReceived ,
	    0 AS SumRecursive

    FROM    tblAcc_Recieved 
	    INNER JOIN dbo.tFacM ON dbo.tblAcc_Recieved.Branch = dbo.tFacM.Branch AND dbo.tblAcc_Recieved.intSerialNo = dbo.tFacM.intSerialNo
        INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    WHERE tblAcc_Recieved.intSerialNo IS NOT NULL AND Status = 10
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,
            tblAcc_Recieved.[Date] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Recieved.UID AS [User] ,
            0 AS SumPrice ,
            0 AS CarryFeeTotal ,
            0 AS DiscountTotal ,
            0 AS StationID ,
            0 AS ServiceTotal ,
            0 AS PackingTotal ,
            0 AS TaxTotal ,
            0 AS DutyTotal ,
            0 AS RoundDiscount ,
            0 AS FacPayment ,
            0 AS Balance ,
            NULL AS Customer ,
            NULL AS CustomerName ,
            NULL AS CustomerTafsili ,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS UserFullName ,
            dbo.tPer.tafsili AS PersonTafsili ,
            CASE dbo.tPer.Gender
              WHEN 1 THEN N'آقاي'
              WHEN 0 THEN N'خانم'
            END AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    Bestankar AS ManualRecieved ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived ,
	    0 AS SumRecursive
    FROM    tblAcc_Recieved 
        INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    WHERE intSerialNo IS NULL 


GO



SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


-------------------------------------------------
-------------------------گزارش خلاصه فروش صندوق
ALTER     PROCEDURE [dbo].[GetStationSaleSummeryInfo]
    (
      --@intLanguage1 INT = Null ,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @Date1 NVARCHAR(50),
      @Date2 NVARCHAR(50),
      @Time1 NVARCHAR(50),
      @Time2 NVARCHAR(50),
      @user1 INT,
      @user2 INT,
      @Station1 INT,
      @Station2 INT,
      @Branch1 INT,
      @Branch2 INT       
    )
AS 
    DECLARE @intLanguage1 INT
--    SET @intLanguage1 = 0   
    IF @intLanguage1 IS NULL 
        SET @intLanguage1 = 0

    DECLARE @strTmp NVARCHAR(50)
    DECLARE @intTmp INT
    DECLARE @Time3 NVARCHAR(50)
    DECLARE @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @Time2

    IF @Date2 < @Date1 
        BEGIN
            SET @strTmp = @Date2
            SET @Date2 = @Date1
            SET @Date1 = @strTmp
        END

    IF @Time2 < @Time1 
        BEGIN 
		/*SET @strTmp   = @Time2
		SET @Time2   = @Time1
		SET @Time1 = @strTmp*/
            SET @Time3 = '00:00'
            SET @Time4 = '24:00'
        END

    IF @user2 < @user1 
        BEGIN
            SET @intTmp = @user2
            SET @user2 = @user1
            SET @user1 = @intTmp
        END
	
    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage1 = 0 
        SET @TimeTitle = N' ساعت : '
    ELSE 
        SET @TimeTitle = N'Time: '

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
    SELECT  SUM(VwStationSaleSummery.SumPrice) AS SumPriceTotal,
            SUM(VwStationSaleSummery.CarryFeeTotal) AS SumCarryFee,
            SUM(VwStationSaleSummery.PackingTotal) AS SumPacking,
            SUM(VwStationSaleSummery.DiscountTotal) AS SumDiscount,
            SUM(VwStationSaleSummery.ServiceTotal) AS SumService,
            SUM(VwStationSaleSummery.GarsonSumPrice) AS SumGarsonSale,
            SUM(VwStationSaleSummery.CarrierSumPrice) AS SumCarrierSale,
            SUM(VwStationSaleSummery.CarrierDebit) AS SumCarrierDebit,
            SUM(VwStationSaleSummery.GarsonDebit) AS SumGarsonDebit,
            SUM(VwStationSaleSummery.CustomerDebit) AS SumCustomerDebit,
            SUM(VwStationSaleSummery.UnBalanceFich) AS SumUnBalanceFich,
            SUM(VwStationSaleSummery.Payment) AS SumPayment,
            SUM(VwStationSaleSummery.Recieved) AS SumRecieved,
            VwStationSaleSummery.[Date] ,
	    SUM(VwStationSaleSummery.TaxTotal) AS SumTax ,
	    SUM(VwStationSaleSummery.DutyTotal) AS SumDuty,
	    SUM(VwStationSaleSummery.TipAmount) AS SumTip,
	    SUM(VwStationSaleSummery.RoundDiscount) AS sumRoundDiscount,
	    SUM(VwStationSaleSummery.CardReceived) AS SumCardReceived,
	    SUM(VwStationSaleSummery.ManualRecieved) AS SumManualRecieved,
	    SUM(VwStationSaleSummery.OrderPrice) AS OrderPrice ,
	    SUM(VwStationSaleSummery.OrderReceived) AS SumOrderReceived ,
	    SUM(VwStationSaleSummery.SumRecursive) AS SumRecursive ,
            
            @Date1 AS DateBefore,
            @Date2 AS DateAfter,
            @Time1 AS FromTime,
            @Time2 AS ToTime,
            @user1 AS FormUser,
            @user2 AS ToUser,
            @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS Sysdate --,
			,VwStationSaleSummery.Branch , @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2
			, dbo.tBranch.nvcBranchName , VwStationSaleSummery.[User] AS Uid , VwStationSaleSummery.UserFullName
    FROM    VwStationSaleSummery 
    	INNER JOIN dbo.tBranch ON dbo.VwStationSaleSummery.Branch = dbo.tBranch.Branch 
    	LEFT OUTER JOIN dbo.tStations ON dbo.VwStationSaleSummery.StationID = dbo.tStations.StationID --AND dbo.VwStationSaleSummery.Branch = dbo.tStations.Branch
    WHERE   VwStationSaleSummery.[Date] >= @Date1
            AND VwStationSaleSummery.[Date] <= @Date2
            AND ( ( VwStationSaleSummery.[Time] >= @Time1
                    AND VwStationSaleSummery.[Time] <= @Time4
                  )
                  OR ( VwStationSaleSummery.[Time] <= @Time2
                       AND VwStationSaleSummery.[Time] >= @Time3
                     )
                )
            AND ( ( VwStationSaleSummery.StationID >= @Station1
                    AND VwStationSaleSummery.StationID <= @Station2
                  )
                  OR VwStationSaleSummery.StationID = 0
                )
            AND ((VwStationSaleSummery.[User] >= @user1 AND VwStationSaleSummery.[User] <= @user2) ) -- OR (dbo.tStations.StationType = 8 AND VwStationSaleSummery.Balance = 0))
            AND VwStationSaleSummery.Branch >= @Branch1
            AND VwStationSaleSummery.Branch <= @Branch2
    GROUP BY VwStationSaleSummery.[Date] , VwStationSaleSummery.Branch , dbo.tBranch.nvcBranchName , VwStationSaleSummery.[User] , VwStationSaleSummery.UserFullName
 --,VwStationSaleSummery.[User],VwStationSaleSummery.StationID,
	--VwStationSaleSummery.UserGender,VwStationSaleSummery.UserFullName



GO
