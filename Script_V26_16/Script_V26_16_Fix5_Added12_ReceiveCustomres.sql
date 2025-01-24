
SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER PROCEDURE [dbo].[PayFactors_CustCredit_Account2]
    (
      @strSelectedFactors NVARCHAR(4000),
      @strSelectedIntSerialNos NVARCHAR(4000),
      @Uid INT,
      @Customer BIGINT,
      @SumPrice BIGINT,
      @AccountYear SMALLINT,
      @intSerialNo BIGINT
    )
AS 
    BEGIN TRAN 
    IF @intSerialNo = 0 
        SET @intSerialNo = NULL

    DECLARE @newtime NVARCHAR(5)
    SELECT  @newtime = dbo.setTimeFormat(GETDATE())
    DECLARE @RegDate NVARCHAR(20)
    SELECT  @RegDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())

--     DECLARE @NoRecieved BIGINT
--     SET @NoRecieved = ( SELECT  ISNULL(MAX([No]), 0) + 1 AS [No]
--                         FROM    tblAcc_Recieved
--                         WHERE   AccountYear = @AccountYear
--                                 AND Branch = dbo.Get_Current_Branch()
--                       )
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
                            @RegDate ,
                            @RegDate ,
                            @NewTime ,
                            @Uid ,
                            N'œ—Ì«›  »«»   ”ÊÌÂ ›Ì‘ '  + CAST([tFacM].[No] AS NVARCHAR(7)) ,  --@strSelectedFactors    +
                            [dbo].[tFacM].[SumPrice] ,
                            [dbo].[tFacM].Branch ,
                            3 , --5
                            [dbo].[tFacM].[Customer] ,
                            [dbo].[tFacM].[intSerialNo] ,
                            @AccountYear
                    FROM    [dbo].[tFacM]
					LEFT OUTER JOIN dbo.tblAcc_Recieved ON dbo.tFacM.Branch = dbo.tblAcc_Recieved.Branch
                    WHERE   [dbo].[tFacM].intSerialNo = @intSerialNo 
                    GROUP BY [dbo].[tFacM].[Date] ,
                            [dbo].[tFacM].[SumPrice] ,
                            [dbo].[tFacM].[intSerialNo] ,dbo.tFacM.Branch ,
							[dbo].[tFacM].[Customer] ,
							[dbo].[tFacM].[No]

  			    FETCH NEXT FROM Serials INTO @intSerialNo
              END
        CLOSE Serials
        DEALLOCATE Serials

--     EXEC Insert_tblAcc_Recieved @NoRecieved, 1, @RegDate, @Uid, @Description,
--         @SumPrice, 3, @Customer, @Uid, @AccountYear, @intSerialNo
	 
    IF RTRIM(LTRIM(@strSelectedFactors)) <> '' 
        BEGIN
            UPDATE  [dbo].[tFacM]
            SET     [FacPayment] = 1,
                    [Balance] = 1
            WHERE   [intSerialNo] IN (
                    SELECT  CAST(word AS BIGINT)
                    FROM    dbo.SplitWithDelimiterNVarChar(@strSelectedIntSerialNos, ',') )
                    --AND dbo.tFacM.Branch = dbo.Get_Current_Branch()
        END

    COMMIT TRAN


GO

