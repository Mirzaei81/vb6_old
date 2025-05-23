
--Script_V26_16_Fix8_10_RepStationSaleSummary.sql
--دریافت روی فیش از طریق صندوق و پیک و فرم گارسون فقط در همان روز
--خلاصه تر شدن گزارش خلاصه فروش صندوق
-- و افزودن دریافت کارت و نقدی و جمع آنها
--سورت کردن تاریخچه فیش
--  افزودن مبلغ مرجوعی هر کاربر به تفکیک در گزارش 
--93/08/12


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
--New

ALTER VIEW [dbo].[VwStationSaleSummery]
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
	        dbo.tfacm.[RoundDiscount],
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
			AND tFactorAdditionalServices.intServiceNo = 3
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS Received,intSerialNo , Branch FROM  [dbo].[tFacCash] 
			GROUP BY intSerialNo , Branch) AS Resived ON Resived.Branch = tfacm.Branch AND  Resived.intSerialNo = dbo.tFacM.intSerialNo
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS CardReceived,intSerialNo , Branch FROM  [dbo].[tFacCard] 
			GROUP BY intSerialNo , Branch) AS CardReceived ON CardReceived.Branch = tfacm.Branch AND  CardReceived.intSerialNo = dbo.tFacM.intSerialNo
	    LEFT outer JOIN (SELECT ISNULL(SUM(Bestankar),0) AS PreReceived2 ,intSerialNo , Branch FROM  [dbo].[tblAcc_Recieved] 
			GROUP BY intSerialNo , Branch) AS PreReceived2 ON PreReceived2.Branch = tfacm.Branch AND  PreReceived2.intSerialNo = dbo.tFacM.intSerialNo

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
ALTER    PROCEDURE [dbo].[GetStationSaleSummeryInfo]
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


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER   Proc Get_History_By_intSerialNo (@intSerialNo bigint  )

as

Declare @No bigint

Set  @No = (select [No] from dbo.tFacM Where intSerialNo = @intSerialNo AND Branch =  dbo.Get_Current_Branch() )
Select @No as [No] , dbo.tHistory.* , tAction.ActionDescription from dbo.tHistory 
	INNER JOIN dbo.tAction ON dbo.tHistory.ActionCode = dbo.tAction.ActionCode
Where intSerialNo = @intSerialNo  AND Branch =  dbo.Get_Current_Branch()
ORDER BY Code 

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER   PROCEDURE [dbo].[Get_PayFactors]
(@intSerialNo BigInt , @Branch INT ) 

AS


 Select Bestankar AS intAmount, Date , RegTime  , N' --- ' AS Type  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @Branch 
 UNION ALL 
 Select intAmount , dbo.tFacM.Date , dbo.tFacM.[Time] AS RegTime , N'نقد' AS Type  From   dbo.tFacCash 
	INNER JOIN dbo.tFacM ON dbo.tFacM.intSerialNo = tFacCash.intSerialNo AND dbo.tFacM.Branch = tFacCash.Branch
 Where tFacCash.intSerialNo = @intSerialNo  and tFacCash.Branch = @Branch 
 UNION ALL 
 Select intAmount , dbo.tFacM.Date , dbo.tFacM.[Time] AS RegTime , N'كارت' AS Type   From   dbo.tFacCard
	INNER JOIN dbo.tFacM ON dbo.tFacM.intSerialNo = tFacCard.intSerialNo AND dbo.tFacM.Branch = tFacCard.Branch
 Where tFacCard.intSerialNo = @intSerialNo  and tFacCard.Branch = @Branch 
 UNION ALL 
 SELECT intChequeAmount AS  intAmount , dbo.tFacM.Date , dbo.tFacM.[Time] AS RegTime , N'چك' AS Type   From   dbo.tFacCheque
	INNER JOIN dbo.tFacM ON dbo.tFacM.intSerialNo = tFacCheque.intSerialNo AND dbo.tFacM.Branch = tFacCheque.Branch
 Where tFacCheque.intSerialNo = @intSerialNo  and tFacCheque.Branch = @Branch 
 UNION ALL 
 SELECT  intAmount , dbo.tFacM.Date , dbo.tFacM.[Time] AS RegTime , N'بن' AS Type   From   dbo.tFacCredit
	INNER JOIN dbo.tFacM ON dbo.tFacM.intSerialNo = tFacCredit.intSerialNo AND dbo.tFacM.Branch = tFacCredit.Branch
 Where tFacCredit.intSerialNo = @intSerialNo  and tFacCredit.Branch = @Branch 

	 
ORDER BY date , RegTime


GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE Update_tfacm_Balance  
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
   Update tfacm
     set Balance = 1 , FacPayment = 1 , [User] = @Uid --, BitLock = 1
         Where [No] = @No And Status = @Status  And Branch = @Branch and AccountYear = @AccountYear

    DECLARE @Date AS NVARCHAR(10)
    SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())
	DECLARE @FichDate AS NVARCHAR(10)
	SET @FichDate = (SELECT [Date] FROM tfacm Where [No] = @No And Status = @Status  And Branch = @Branch and AccountYear = @AccountYear)
	
	IF (@Status =  1 OR @Status = 2 )  
		BEGIN 
		--IF @Date = @FichDate    
			exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds ,  @Branch  , 0       
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




GO




SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[PayFactors_Table]
    (
      @strSelectedFactors NVARCHAR(4000),
      @strSelectedTables NVARCHAR(4000),
      @Uid INT 
    )
AS 
    DECLARE @newtime NVARCHAR(5)
    SELECT  @newtime = dbo.setTimeFormat(GETDATE())
    DECLARE @RegDate NVARCHAR(20)
    SELECT  @RegDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())

	DECLARE @Branch INT 
    SET @Branch = (SELECT  TOP 1 Branch 
    FROM    tFacM
    WHERE   intSerialNo IN (
            SELECT  CAST (word AS BIGINT)
            FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors, ',') ))


    IF RTRIM(LTRIM(@strSelectedFactors)) <> N'' 
        BEGIN
            UPDATE  tFacM
            SET     FacPayment = 1, Balance = 1 , [User] = @Uid
            WHERE   intSerialNo IN (
                    SELECT  CAST(word AS BIGINT)
                    FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,
                                                           N',') )
                    AND dbo.tFacM.Branch = @Branch

            UPDATE  dbo.tblAcc_Recieved
            SET     Bestankar = 0
            WHERE   intSerialNo IN (
                    SELECT  CAST(word AS BIGINT)
                    FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,
                                                           N',') )
                    AND dbo.tblAcc_Recieved.Branch = @Branch

			DELETE FROM tfaccash WHERE intSerialNo IN 
			(SELECT  CAST(word AS BIGINT) FROM dbo.SplitWithDelimiterNVarChar(@strSelectedFactors, N',') )

            INSERT  INTO tfaccash
                    (
                      Branch,
                      intserialno,
                      intAmount 
                    )
                    SELECT  Branch,
                            intserialno,
                            Sumprice
                    FROM    tfacm
                    WHERE   intSerialNo IN (
                            SELECT  CAST(word AS BIGINT)
                            FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors, N',') )
                            AND dbo.tFacM.Branch = @Branch

            IF @strSelectedTables <> N'' 
				Begin
                UPDATE  tTable
                SET     Empty = 1
                WHERE   [No] IN (
                        SELECT  CAST(word AS BIGINT)
                        FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedTables, N',') )
                        AND dbo.tTable.Branch = @Branch               

				If dbo.Get_TableMonitoring() = 1 		---Table Monitoring
				UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.nvcEndTime=  dbo.SetTimeFormat(getdate())      
				WHERE  tblSamar_TableUsage.intTableNo IN (
                        SELECT  CAST(word AS BIGINT)
                        FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedTables, N',') )
                        AND dbo.tblSamar_TableUsage.intBranch = @Branch               
						AND tblSamar_TableUsage.nvcEndTime IS NULL 
				
				END 
        END
--===============================================


GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER PROCEDURE [dbo].[PayFactors_Payk]
    (
      @strSelectedFactors NVARCHAR(4000) ,
      @Uid INT 
    )
AS 
    DECLARE @NewTime NVARCHAR(5)  
    SELECT  @NewTime = dbo.[SetTimeFormat](GETDATE())  
    DECLARE @RegDate NVARCHAR(20)  
    SELECT  @RegDate =   [dbo].[shamsi](GETDATE())

    DECLARE @Date AS NVARCHAR(10)
--    SET @Date = (
--                  SELECT    GETDATE()
--                )
    SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())

    DECLARE @NoRec AS INT 
    SET @NoRec = (
                   SELECT   MAX(DISTINCT ( [No] )) + 1
                   FROM     [tblAcc_Recieved]
                 )
	DECLARE @Branch INT 
    SET @Branch = (SELECT  TOP 1 Branch 
    FROM    tFacM
    WHERE   intSerialNo IN (
            SELECT  CAST (word AS BIGINT)
            FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors, ',') ))

    IF RTRIM(LTRIM(@strSelectedFactors)) <> '' 
        BEGIN  
            UPDATE  tFacM
            SET     FacPayment = 1 ,
                    Balance = 1 , [User] = @Uid
            WHERE   intSerialNo IN (
                    SELECT  CAST (word AS BIGINT)
                    FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,
                                                           ',') )
                    AND dbo.tFacM.Branch = @Branch  
 
            INSERT  INTO dbo.tblAcc_History
                    ( [Date] ,
                      [Time] ,
                      [No] ,
                      Status ,
                      UID ,
                      ActionCode ,
                      Bedehkar ,
                      Bestankar 
                    )
                    SELECT  [dbo].[tFacM].[Date] ,
                            @NewTime ,
                            [dbo].[tFacM].[No] ,
                            2 ,
                            @Uid ,
                            6 ,
                            [dbo].[tFacM].[SumPrice] ,
                            0
                    FROM    [dbo].[tFacM]
                    WHERE   [intSerialNo] IN (
                            SELECT  CAST (word AS BIGINT)
                            FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,
                                                              ',') )  
		                    AND dbo.tFacM.Branch = @Branch  
		
		--Save Receive in same tfacm date
			DELETE FROM tfaccash WHERE intSerialNo IN 
			(SELECT  CAST(word AS BIGINT) FROM dbo.SplitWithDelimiterNVarChar(@strSelectedFactors, N',') )

            INSERT  INTO tfaccash
                    (
                      Branch,
                      intserialno,
                      intAmount 
                    )
                    SELECT  Branch,
                            intserialno,
                            Sumprice
                    FROM    tfacm
                    WHERE   intSerialNo IN (
                            SELECT  CAST(word AS BIGINT)
                            FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors, N',') )
                            AND dbo.tFacM.Branch = @Branch


	 --   DECLARE @intSerialNo INT 
	 --   DECLARE Serials CURSOR FOR
		--SELECT  CAST (word AS BIGINT) AS intserialNo
  --                          FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,',')  
     
  --      OPEN Serials
  --      FETCH NEXT FROM Serials INTO @intSerialNo
  --      WHILE @@FETCH_STATUS = 0 
  --          BEGIN

  --         INSERT  INTO dbo.[tblAcc_Recieved]
  --                  ( Code ,
  --                    [No] ,
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
  --                          ISNULL(MAX([tblAcc_Recieved].[No]), 0) + 1 ,
  --                          1 ,
  --                          @Date ,
  --                          @RegDate ,
  --                          @NewTime ,
  --                          @Uid ,
  --                          N'دريافت از پيك بابت فاكتور ' + CAST( [tFacM].[No] AS NVARCHAR(7)) ,
  --                          [dbo].[tFacM].[SumPrice] ,
  --                          @Branch ,
  --                          3 , --5
  --                          [dbo].[tFacM].[Customer] ,
  --                          [dbo].[tFacM].[intSerialNo] ,
  --                          [dbo].[Get_AccountYear]()
  --                  FROM    [dbo].[tFacM]
		--					LEFT OUTER JOIN dbo.tblAcc_Recieved ON dbo.tFacM.Branch = dbo.tblAcc_Recieved.Branch
  --                  WHERE   [dbo].[tFacM].intSerialNo = @intSerialNo   --- IN (
  --                  --        SELECT  CAST (word AS BIGINT)
  --                  --        FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,
  --                  --                                          ',') )
  --                          --AND [tFacM].[Date] <> @Date
  --                  GROUP BY [dbo].[tFacM].[Date] ,
  --                          [dbo].[tFacM].[SumPrice] ,
  --                          [dbo].[tFacM].[intSerialNo] ,
		--					[dbo].[tFacM].[Customer] ,
		--					[dbo].[tFacM].[No]


  --			    FETCH NEXT FROM Serials INTO @intSerialNo
  --            END
  --      CLOSE Serials
  --      DEALLOCATE Serials

  			    
        END
--===============================================



GO





