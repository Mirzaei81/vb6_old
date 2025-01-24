--New

ALTER   VIEW [dbo].[VwStationSaleSummery]
AS 
    SELECT  dbo.tFacM.[No] ,
            dbo.tFacM.[Date] ,
            dbo.tFacM.[Time] ,
            dbo.tFacM.[User] ,
            SumPrice ,
            CarryFeeTotal ,
            DiscountTotal ,
            StationID ,
            ServiceTotal ,
            PackingTotal ,
            TaxTotal ,
            DutyTotal ,
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
              WHEN 1 THEN N'ÂÞÇí'
              WHEN 0 THEN N'ÎÇäã'
            END AS UserGender ,
            CASE ISNULL(Incharge, 0)
              WHEN 0 THEN 0
              ELSE CASE ISNULL(TableNo, 0)
                     WHEN 0 THEN SumPrice
                     ELSE 0
                   END
            END AS CarrierSumPrice ,
            CASE ISNULL(Incharge, 0)
              WHEN 0 THEN 0
              ELSE CASE ISNULL(TableNo, 0)
                     WHEN 0 THEN 0
                     ELSE SumPrice
                   END
            END AS GarsonSumPrice ,
            CASE FacPayment
              WHEN 0 THEN CASE Balance
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
            CASE FacPayment
              WHEN 0 THEN CASE ISNULL(Incharge, 0)
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
			CASE Balance
              WHEN 0 THEN 
				CASE WHEN (Facpayment = 1 or (Incharge is NULL AND serveplace <> 2 AND serveplace <> 16)) THEN 
					SumPrice - (ISNULL(Resived.Received , 0) +ISNULL(ChequeReceived.ChequeReceived , 0)+ISNULL(CardReceived.CardReceived , 0)+ISNULL(BonReceived.BonReceived , 0)+ISNULL(PreReceived2.PreReceived2 , 0))
					ELSE 0
					END 
				ELSE 0
			END AS CustomerDebit ,
            CASE Balance
              WHEN 0 THEN CASE FacPayment
                            WHEN 0 THEN CASE ISNULL(Incharge, 0)
                                          WHEN 0 THEN SumPrice
                                          ELSE 0
                                        END
                            ELSE 0
                          END
              ELSE 0
            END AS UnBalanceFich ,
            dbo.tFacM.Branch ,
            0 AS Payment ,
            ISNULL(Resived.Received , 0) AS Recieved ,
            tper.ppno ,
            tfacm.status ,
            ISNULL(tblAcc_Recieved.bestankar , 0) AS preRecieved ,
            ISNULL(ChequeReceived.ChequeReceived ,0) AS ChequeRecieved ,
            ISNULL(CardReceived.CardReceived , 0) AS CardReceived ,
            ISNULL(BonReceived.BonReceived , 0) AS BonReceived ,
            ISnull(tFactorAdditionalServices.amount , 0)  AS TipAmount ,
			0 AS ManualRecieved ,
			ISNULL(PreReceived.PreReceived , 0) AS TablePreReceived
			, 0 AS OrderPrice 
			, 0 AS OrderReceived 
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
            INNER JOIN tCUst ON tfacM.Customer = tcust.code
            LEFT OUTER JOIN [tblAcc_Recieved] ON tfacm.[OrderRefrence] = [tblAcc_Recieved].[intSerialNo] 
            LEFT OUTER JOIN [tFactorAdditionalServices] ON [tFactorAdditionalServices].[Branch] = [tFacM].[Branch] AND [tFactorAdditionalServices].[intSerialNo] = [tFacM].[intSerialNo]
		AND tFactorAdditionalServices.intServiceNo = 3
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS Received,intSerialNo , Branch FROM  [dbo].[tFacCash] 
			GROUP BY intSerialNo , Branch) AS Resived ON Resived.Branch = tfacm.Branch AND  Resived.intSerialNo = dbo.tFacM.intSerialNo
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS CardReceived,intSerialNo , Branch FROM  [dbo].[tFacCard] 
			GROUP BY intSerialNo , Branch) AS CardReceived ON CardReceived.Branch = tfacm.Branch AND  CardReceived.intSerialNo = dbo.tFacM.intSerialNo
	    LEFT outer JOIN (SELECT ISNULL(SUM(intAmount),0) AS BonReceived,intSerialNo , Branch FROM  [dbo].[tFacCredit] 
			GROUP BY intSerialNo , Branch) AS BonReceived ON BonReceived.Branch = tfacm.Branch AND  BonReceived.intSerialNo = dbo.tFacM.intSerialNo
	    LEFT outer JOIN (SELECT ISNULL(SUM(intChequeAmount),0) AS ChequeReceived,intSerialNo , Branch FROM  [dbo].[tFacCheque] 
			GROUP BY intSerialNo , Branch) AS ChequeReceived ON ChequeReceived.Branch = tfacm.Branch AND  ChequeReceived.intSerialNo = dbo.tFacM.intSerialNo
	    LEFT outer JOIN (SELECT ISNULL(SUM(Bestankar),0) AS PreReceived ,intSerialNo , Branch FROM  [dbo].[tblAcc_Recieved] 
			GROUP BY intSerialNo , Branch) AS PreReceived ON PreReceived.Branch = tfacm.Branch AND  PreReceived.intSerialNo = dbo.tFacM.intSerialNo AND dbo.tFacM.ServePlace = 16
	    LEFT outer JOIN (SELECT ISNULL(SUM(Bestankar),0) AS PreReceived2 ,intSerialNo , Branch FROM  [dbo].[tblAcc_Recieved] 
			GROUP BY intSerialNo , Branch) AS PreReceived2 ON PreReceived2.Branch = tfacm.Branch AND  PreReceived2.intSerialNo = dbo.tFacM.intSerialNo

    WHERE   ( Recursive = 0  AND Status = 2 )

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
              WHEN 1 THEN N'ÂÞÇí'
              WHEN 0 THEN N'ÎÇäã'
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
            0 AS preRecieved ,
            0 AS ChequeRecieved ,
	    0 AS CardReceived ,
	    0 AS BonReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS TablePreReceived ,
	    SumPrice AS OrderPrice ,
	    0 AS OrderReceived
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
              WHEN 1 THEN N'ÂÞÇí'
              WHEN 0 THEN N'ÎÇäã'
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
            0 AS preRecieved ,
            0 AS ChequeRecieved ,
	    0 AS CardReceived ,
	    0 AS BonReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS TablePreReceived ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived
    FROM    tblAcc_Cash
            INNER JOIN dbo.tUser ON dbo.tblAcc_Cash.[UID] = dbo.tUser.UID
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    UNION
    SELECT  tblAcc_Recieved.Code AS [No] ,--list
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
              WHEN 1 THEN N'ÂÞÇí'
              WHEN 0 THEN N'ÎÇäã'
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
            0 AS preRecieved ,
            0 AS ChequeRecieved ,
	    0 AS CardReceived ,
	    0 AS BonReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS TablePreReceived ,
	    0 AS OrderPrice ,
  	    0 AS OrderReceived

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
              WHEN 1 THEN N'ÂÞÇí'
              WHEN 0 THEN N'ÎÇäã'
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
            0 AS preRecieved ,
            0 AS ChequeRecieved ,
	    0 AS BonReceived ,
	    0 AS CardReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS TablePreReceived ,
	    0 AS OrderPrice ,
  	    Bestankar AS OrderReceived

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
              WHEN 1 THEN N'ÂÞÇí'
              WHEN 0 THEN N'ÎÇäã'
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
            0 AS preRecieved ,
            0 AS ChequeRecieved ,
	    0 AS CardReceived ,
	    0 AS BonReceived ,
	    0 AS TipAmount ,
	    Bestankar AS ManualRecieved ,
	    0 AS TablePreReceived ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived
    FROM    tblAcc_Recieved 
        INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.[UID] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
    WHERE intSerialNo IS NULL 
    UNION
    SELECT  [No] ,
            tblAcc_Recieved_Cheque.[RegDate] AS [Date] ,
            RegTime AS [Time] ,
            tblAcc_Recieved_Cheque.UID AS [User] ,
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
            NULL AS PersonTafsili ,
            NULL AS UserGender ,
            0 AS CarrierSumPrice ,
            0 AS GarsonSumPrice ,
            0 AS CarrierDebit ,
            0 AS GarsonDebit ,
            0 AS CustomerDebit ,
            0 AS UnbalaceFich ,
            tblAcc_Recieved_Cheque.Branch AS Branch ,
            0 AS Payment ,
            0 AS Recieved ,
            NULL AS ppno ,
            2 AS status ,
            0 AS preRecieved ,
            intChequeAmount AS ChequeRecieved ,
	    0 AS CardReceived ,
	    0 AS BonReceived ,
	    0 AS TipAmount ,
	    0 AS ManualRecieved ,
	    0 AS TablePreReceived ,
	    0 AS OrderPrice ,
	    0 AS OrderReceived
    FROM    tblAcc_Recieved_Cheque
        INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved_Cheque.[UID] = dbo.tUser.UID
        INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
--===============================================


GO


