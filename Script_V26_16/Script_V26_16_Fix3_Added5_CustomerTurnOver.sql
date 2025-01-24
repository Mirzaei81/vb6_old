


ALTER    VIEW [dbo].[vw_CustomerBillPayment]
AS 
    SELECT  tfacm.AccountYear ,
            tcust.Code ,
            tfacm.date ,
--             CASE WHEN ( dbo.tfacm.balance = 0
--                         AND dbo.tfacm.facpayment = 1
--                       ) THEN SUM(sumprice)
--                  ELSE 0
--             END AS notPaidFactor ,
--             CASE WHEN ( dbo.tfacm.balance = 1
--                         AND dbo.tfacm.facpayment = 1
--                       ) THEN SUM(sumprice)
--                  ELSE 0
--             END AS PaidFactor ,
            CAST(SUM(dbo.tFacM.sumprice)AS MONEY) AS SumSale ,
            CAST(SUM(Isnull(tFacCash_1.intAmount , 0))AS MONEY) AS paidfactor ,
			membershipid ,
            0 AS paid ,
            CASE WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN CASE dbo.tcust.Sex
                        WHEN 0
                        THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                        ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                      END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name = ''
                      ) THEN CASE dbo.tcust.Sex
                               WHEN 0 THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                               ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName <> ''
                      ) THEN dbo.tcust.WorkName
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                           WHEN 0
                           THEN N' ÎÇäã '
                                + dbo.tcust.Family + ' '
                                + dbo.tcust.Name
                           ELSE N' ÂÞÇí '
                                + dbo.tcust.Family + ' '
                                + dbo.tcust.Name
                         END
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name = ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                           WHEN 0
                           THEN N' ÎÇäã '
                                + dbo.tcust.Family + ' '
                           ELSE N' ÂÞÇí '
                                + dbo.tcust.Family + ' '
                         END
            END AS family
	FROM    tfacm
    	INNER JOIN tcust ON tfacm.customer = tcust.code
                               -- AND ( tfacm.Branch = tcust.Branch
                               --       OR tCust.Branch IS NULL
                               --     )
	LEFT OUTER JOIN (SELECT SUM(intAmount) AS intAmount , Branch , intSerialNo FROM tFacCash GROUP BY Branch , intSerialNo ) AS tFacCash_1 ON dbo.tFacM.Branch = tFacCash_1.Branch AND tFacCash_1.intSerialNo = tFacM.intSerialNo
    WHERE  FacPayment = 1 AND  Status = 2
    GROUP BY code ,
            Balance ,
            FacPayment ,
            MembershipId ,
            Family ,
            NAME ,
            sex ,
            mastercode ,
            workname ,
            tfacm.date ,
            tfacm.AccountYear
    UNION ALL
    SELECT  tblAcc_Recieved.AccountYear ,
            tcust.code ,
            [tblAcc_Recieved].date ,
            0 AS SumSale ,
            0 AS paidfactor ,
            membershipid ,
            ISNULL(SUM(Bestankar), 0) AS paid ,
            CASE WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN CASE dbo.tcust.Sex
                        WHEN 0
                        THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                        ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                      END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name = ''
                      ) THEN CASE dbo.tcust.Sex
                               WHEN 0 THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                               ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName <> ''
                      ) THEN dbo.tcust.WorkName
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' ÎÇäã '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                               ELSE N' ÂÞÇí '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                             END
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name = ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' ÎÇäã '
                                                    + dbo.tcust.Family + ' '
                                               ELSE N' ÂÞÇí '
                                                    + dbo.tcust.Family + ' '
                                             END
            END AS family
    FROM    tblAcc_Recieved
            INNER JOIN tcust ON tcust.code = tblAcc_Recieved.Code_Bes
                                --AND ( tblAcc_Recieved.Branch = tcust.Branch
                                --      OR tCust.Branch IS NULL
                                --    )
    WHERE   RecieveType = 3
    GROUP BY tcust.code ,
            [membershipid] ,
            family ,
            NAME ,
            sex ,
            mastercode ,
            workname ,
            tblAcc_Recieved.Date ,
            tblAcc_Recieved.AccountYear
    UNION ALL
    SELECT  tblAcc_Cash.AccountYear ,
            tcust.code ,
            [tblAcc_Cash].date ,
            0 AS SumSale ,
            0 AS paidfactor ,
            membershipid ,
            ISNULL(SUM(Bestankar), 0)* -1 AS paid ,
            CASE WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN CASE dbo.tcust.Sex
                        WHEN 0
                        THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                        ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                      END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name = ''
                      ) THEN CASE dbo.tcust.Sex
                               WHEN 0 THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                               ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName <> ''
                      ) THEN dbo.tcust.WorkName
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' ÎÇäã '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                               ELSE N' ÂÞÇí '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                             END
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name = ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' ÎÇäã '
                                                    + dbo.tcust.Family + ' '
                                               ELSE N' ÂÞÇí '
                                                    + dbo.tcust.Family + ' '
                                             END
            END AS family
    FROM    tblAcc_Cash
            INNER JOIN tcust ON tcust.code = tblAcc_Cash.Uid_Bede
                               -- AND ( tblAcc_Cash.Branch = tcust.Branch
                               --       OR tCust.Branch IS NULL
                               --     )
    WHERE   PaymentType = 6
    GROUP BY tcust.code ,
            [membershipid] ,
            family ,
            NAME ,
            sex ,
            mastercode ,
            workname ,
            tblAcc_Cash.Date ,
            tblAcc_Cash.AccountYear
    UNION ALL
    SELECT  tblAcc_Recieved_Cheque.AccountYear ,
            tcust.code ,
            [tblAcc_Recieved_Cheque].Regdate AS [date] ,
            0 AS SumSale ,
            0 AS paidfactor ,
            membershipid ,
            ISNULL(SUM(intChequeAmount), 0) AS paid ,
            CASE WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN CASE dbo.tcust.Sex
                        WHEN 0
                        THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                        ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                      END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name = ''
                      ) THEN CASE dbo.tcust.Sex
                               WHEN 0 THEN N' ÎÇäã ' + dbo.tcust.Family + ' '
                               ELSE N' ÂÞÇí ' + dbo.tcust.Family + ' '
                             END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName <> ''
                      ) THEN dbo.tcust.WorkName
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' ÎÇäã '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                               ELSE N' ÂÞÇí '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                             END
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name = ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' ÎÇäã '
                                                    + dbo.tcust.Family + ' '
                                               ELSE N' ÂÞÇí '
                                                    + dbo.tcust.Family + ' '
                                             END
            END AS family
    FROM    tblAcc_Recieved_Cheque
            INNER JOIN tcust ON tcust.code = tblAcc_Recieved_Cheque.Code_Bes
                             --   AND ( tblAcc_Recieved_Cheque.Branch = tcust.Branch
                             --         OR tCust.Branch IS NULL
                             --       )
    GROUP BY tcust.code ,
            [membershipid] ,
            family ,
            NAME ,
            sex ,
            mastercode ,
            workname ,
            tblAcc_Recieved_Cheque.regDate ,
            tblAcc_Recieved_Cheque.AccountYear
--==============================================



GO



