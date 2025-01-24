SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER     PROCEDURE [dbo].[PayFactors_Payk]
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


	    DECLARE @intSerialNo INT 
	    DECLARE Serials CURSOR FOR
		SELECT  CAST (word AS BIGINT) AS intserialNo
                            FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,',')  
     
        OPEN Serials
        FETCH NEXT FROM Serials INTO @intSerialNo
        WHILE @@FETCH_STATUS = 0 
            BEGIN

           INSERT  INTO dbo.[tblAcc_Recieved]
                    ( Code ,
                      [No] ,
                      [List] ,
                      [Date] ,
                      [RegDate] ,
                      [RegTime] ,
                      [UID] ,
                      [Description] ,
                      [Bestankar] ,
                      [Branch] ,
                      [RecieveType] ,
                      [Code_Bes] ,
                      [intSerialNo] ,
                      [AccountYear]
                    )
                    SELECT  ISNULL(MAX([tblAcc_Recieved].Code), 0) + 1 ,
                            ISNULL(MAX([tblAcc_Recieved].[No]), 0) + 1 ,
                            1 ,
                            @Date ,
                            @RegDate ,
                            @NewTime ,
                            @Uid ,
                            N'œ—Ì«›  «“ ÅÌﬂ »«»  ›«ﬂ Ê— ' + CAST( [tFacM].[No] AS NVARCHAR(7)) ,
                            [dbo].[tFacM].[SumPrice] ,
                            @Branch ,
                            3 , --5
                            [dbo].[tFacM].[Customer] ,
                            [dbo].[tFacM].[intSerialNo] ,
                            [dbo].[Get_AccountYear]()
                    FROM    [dbo].[tFacM]
							LEFT OUTER JOIN dbo.tblAcc_Recieved ON dbo.tFacM.Branch = dbo.tblAcc_Recieved.Branch
                    WHERE   [dbo].[tFacM].intSerialNo = @intSerialNo   --- IN (
                    --        SELECT  CAST (word AS BIGINT)
                    --        FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedFactors,
                    --                                          ',') )
                            --AND [tFacM].[Date] <> @Date
                    GROUP BY [dbo].[tFacM].[Date] ,
                            [dbo].[tFacM].[SumPrice] ,
                            [dbo].[tFacM].[intSerialNo] ,
							[dbo].[tFacM].[Customer] ,
							[dbo].[tFacM].[No]


  			    FETCH NEXT FROM Serials INTO @intSerialNo
              END
        CLOSE Serials
        DEALLOCATE Serials

  			    
        END
--===============================================



GO
