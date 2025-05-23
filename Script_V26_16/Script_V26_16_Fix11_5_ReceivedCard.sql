
--محاسبه دریافت کارت در صورتحساب مشتریان
--93/12/05

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER    PROCEDURE [dbo].[Get_Recieved_tFaccash]
(
@Code Bigint ,
@DateBefore Nvarchar(8) ,
@DateAfter  Nvarchar(8) ,
@AccountYear  SMALLINT ,
@Branch INT 
)
AS
DECLARE @Membershipd BIGINT
SET @Membershipd=(SELECT DISTINCT membershipid FROM [tCust] WHERE [Code]=@Code AND Branch = dbo.Get_Current_Branch())

	SELECT        dbo.tFacM.[No] , dbo.tFacM.[Date] ,N' دريافت بابت تسويه فيش  ' + CAST([No] AS NVARCHAR(10) ) AS Description
		           , Isnull(tFacCash_1.intAmount , 0) AS Bestankar 
			   , dbo.tPer.nvcSurName AS  [USER_NAME] , tfacm.Time AS RegTime ,
  				
            CASE WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN CASE dbo.tcust.Sex
                        WHEN 0
                        THEN N' خانم ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                        ELSE N' آقاي ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                      END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name = ''
                      ) THEN CASE dbo.tcust.Sex
                               WHEN 0 THEN N' خانم ' + dbo.tcust.Family + ' '
                               ELSE N' آقاي ' + dbo.tcust.Family + ' '
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
                                               THEN N' خانم '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                               ELSE N' آقاي '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                             END
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name = ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' خانم '
                                                    + dbo.tcust.Family + ' '
                                               ELSE N' آقاي '
                                                    + dbo.tcust.Family + ' '
                                             END
            END AS family
            , tfacm.RegDate , dbo.tFacM.Branch
	FROM    tfacm
	INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
	INNER JOIN dbo.tPer ON  dbo.tUser.pPno = dbo.tPer.pPno
    	INNER JOIN tcust ON tfacm.customer = tcust.code
                               -- AND ( tfacm.Branch = tcust.Branch
                               --       OR tCust.Branch IS NULL
                               --     )
	INNER JOIN  (SELECT SUM(intAmount) AS intAmount , Branch , intSerialNo FROM tFacCash GROUP BY Branch , intSerialNo ) AS tFacCash_1 ON dbo.tFacM.Branch = tFacCash_1.Branch AND tFacCash_1.intSerialNo = tFacM.intSerialNo

        WHERE	dbo.tFacM.[Date] >= @DateBefore  And  dbo.tFacM.[Date] <= @DateAfter  
                And dbo.tFacM.Customer = @Code
				And dbo.tFacM.AccountYear = @AccountYear
				AND (dbo.tFacM.Branch = @Branch OR @Branch = 0)

UNION ALL 

	SELECT        dbo.tFacM.[No] , dbo.tFacM.[Date] ,N' دريافت کارت بابت تسويه فيش  ' + CAST([No] AS NVARCHAR(10) ) AS Description
		           , Isnull(tFacCrad_1.intAmount , 0) AS Bestankar 
			   , dbo.tPer.nvcSurName AS  [USER_NAME] , tfacm.Time AS RegTime ,
  				
            CASE WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name <> ''
                      )
                 THEN CASE dbo.tcust.Sex
                        WHEN 0
                        THEN N' خانم ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                        ELSE N' آقاي ' + dbo.tcust.Family + ' '
                             + dbo.tcust.Name
                      END
                 WHEN ( dbo.tcust.MasterCode IS NULL
                        AND dbo.tcust.WorkName = ''
                        AND dbo.tcust.Name = ''
                      ) THEN CASE dbo.tcust.Sex
                               WHEN 0 THEN N' خانم ' + dbo.tcust.Family + ' '
                               ELSE N' آقاي ' + dbo.tcust.Family + ' '
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
                                               THEN N' خانم '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                               ELSE N' آقاي '
                                                    + dbo.tcust.Family + ' '
                                                    + dbo.tcust.Name
                                             END
                 WHEN ( dbo.tcust.MasterCode IS NOT NULL
                        AND tcust.WorkName <> ''
                        AND dbo.tcust.Name = ''
                      )
                 THEN tcust.WorkName + '_' + CASE dbo.tcust.Sex
                                               WHEN 0
                                               THEN N' خانم '
                                                    + dbo.tcust.Family + ' '
                                               ELSE N' آقاي '
                                                    + dbo.tcust.Family + ' '
                                             END
            END AS family
            , tfacm.RegDate , dbo.tFacM.Branch
	FROM    tfacm
	INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
	INNER JOIN dbo.tPer ON  dbo.tUser.pPno = dbo.tPer.pPno
    	INNER JOIN tcust ON tfacm.customer = tcust.code
                               -- AND ( tfacm.Branch = tcust.Branch
                               --       OR tCust.Branch IS NULL
                               --     )
	INNER JOIN  (SELECT SUM(intAmount) AS intAmount , Branch , intSerialNo FROM dbo.tFacCard GROUP BY Branch , intSerialNo ) AS tFacCrad_1 ON dbo.tFacM.Branch = tFacCrad_1.Branch AND tFacCrad_1.intSerialNo = tFacM.intSerialNo

        WHERE	dbo.tFacM.[Date] >= @DateBefore  And  dbo.tFacM.[Date] <= @DateAfter  
                And dbo.tFacM.Customer = @Code
				And dbo.tFacM.AccountYear = @AccountYear
				AND (dbo.tFacM.Branch = @Branch OR @Branch = 0)

ORDER BY dbo.tFacM.[No] 




GO
