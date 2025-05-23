


--Script_V26_16_Fix0
--اضافه شدن حسابداری به سیستم فروش رستورانی
--کنترل تاریخ میلادی در گزارشات و ریپورت صندوق 
--افزایش سرعت  گزارشات گروهی کالا و ریز کالا  
-- 92/10/15

INSERT  INTO dbo.tblPub_Script2
        ( Version ,
          Script ,
          LastScriptNo ,
          [Date] ,
          FixNumber
        )
VALUES  ( 26 ,
          16 ,
          0 ,
          dbo.shamsi(GETDATE()) ,
          0
        )
GO

UPDATE dbo.tCust SET Tafsili = NULL 
UPDATE dbo.tSupplier SET Tafsili = NULL 
UPDATE dbo.tper  SET Tafsili = NULL 
go


ALTER PROC Update_transferAccounting
(
  @Branch INT ,
  @DateBefore NVARCHAR(8) ,
  @DateAfter NVARCHAR(8),
  @SanadNo INT 
)


AS
	UPDATE dbo.tFacM
	SET dbo.tFacM.transferAccounting=1	,
		dbo.tFacM.BitLock = 1 ,
		dbo.tFacM.Refrence_Acc = @SanadNo
	WHERE tfacm.Branch = @Branch
		AND tfacm.[Date] >= @DateBefore
		AND tfacm.[Date] <= @DateAfter
		AND [Recursive] = 0
		AND transferAccounting = 0
		AND (Status = 2 OR Status = 5) 
		AND (Customer = -1 OR Customer IS NULL )


GO




ALTER PROCEDURE [dbo].[Get_AccountDocument]
    (
      @Branch INT ,
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @Status INT ,
      @Code INT
    )
AS 
    IF ( @Code = 1 ) 
            BEGIN
                SELECT  tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
                        tPer.Tafsili AS PersonTafsili 
                       , ISNULL(SUM(ISNULL(tFacCash.intAmount , 0)), 0) + ISNULL(SUM(t1.Bestankar1) ,0) AS sp
                       , SUM(t1.Bestankar1) AS aaa
                FROM    tFacM
                        LEFT OUTER JOIN tFacCash ON tFacM.Branch = tFacCash.Branch
                                               AND tFacM.intSerialNo = tFacCash.intSerialNo
                        INNER JOIN tUser ON tUser.UID = tFacM.[User]
                        INNER JOIN tPer ON tUser.pPno = tPer.pPno
						LEFT OUTER JOIN 
							(Select SUM(IsNull(Bestankar,0)) AS Bestankar1 , intSerialNo , Branch From   tblAcc_Recieved GROUP BY intSerialNo , Branch )t1
								ON  t1.intSerialNo = dbo.tFacM.intSerialNo  and t1.Branch = dbo.tFacM.Branch 	
                WHERE    tFacM.Branch = @Branch 
                        AND tFacM.Recursive = 0 
                        AND tFacM.Status =@Status
                        AND dbo.tFacM.transferAccounting=0  
                GROUP BY tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
                        tPer.Tafsili 
                HAVING  tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,
                        tPer.Tafsili
            END

    IF ( @Code = 5 ) 
            BEGIN
                SELECT  tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
                        tPer.Tafsili AS PersonTafsili ,
                        ISNULL(SUM(tFacCard.intAmount), 0) AS sp
                FROM    tUser
                        INNER JOIN tPer ON tUser.pPno = tPer.pPno
                        INNER JOIN tFacM ON tUser.UID = tFacM.[User]
                                           -- AND tUser.Branch = tFacM.Branch
                        INNER JOIN tFacCard ON tFacM.Branch = tFacCard.Branch
                                               AND tFacM.intSerialNo = tFacCard.intSerialNo
                WHERE    tFacM.Branch = @Branch 
                        AND  tFacM.Recursive = 0 
                        AND ( tFacM.Status = @Status
                              OR tFacM.Status = 8
                            )AND dbo.tFacM.transferAccounting=0
                GROUP BY tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
                        tPer.Tafsili
                HAVING   tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,
                        tPer.Tafsili
            END



GO



ALTER   view [dbo].[vw_Get_Cust]  

as  

SELECT  
tcust.[Code], 
cast(tcust.MembershipId as bigint) as MembershipId,   
  CASE WHEN (tcust.workname IS NOT NULL AND tcust.workname <> '')   
  THEN (tcust.workname)   
 WHEN ((tcust.workname IS  NULL OR tcust.workname = '')  and MasterWorkName IS NOT NULL )  
  THEN (tcust.MasterWorkName  + '_' + tcust.family + ' ' + tcust.Name)  
         ELSE (tcust.family + ' ' + tcust.Name)   
 END AS [Name],
  tcust.[Tel1] +' '+ tcust.[Tel2] + ' ' + tcust.[Tel3] +' '  + tcust.[Tel4] + ' ' + tcust.[Mobile] AS Telephone, 
  ISNULL(tcust.[Address] , N'') AS Address,
  	tcust.[MasterCode],
	 tcust.[Prefix],
	 tcust.[Owner],   
 --tcust.[Sex], tcust.[City],  
 tcust.[ActKind],tcust.[ActDeAct],   
 tcust.[Assansor],   
 tcust.[PostalCode], tcust.[Tel1] ,  
 tcust.[Tel2] , tcust.[Tel3] , tcust.[Tel4], tcust.[Mobile],  Tafsili ,
 --tcust.[Fax], tcust.[Email], tcust.[CarryFee],   
 --tcust.[PaykFee],tcust.[Distance], tcust.[Credit],   
 --tcust.[Discount], tcust.[BuyState], tcust.[Description],   
 --tcust.[Date],tcust.[Time], tcust.[User], tcust.[Unit],   
 --tcust.[InternalNo], tcust.[Flour] ,  
         isnull(SUM(T .sumPrice), 0) AS Price , tcust.Central , tcust.FamilyNo, [tCust].[Branch]  
	, Credit
FROM   ( SELECT     dbo.tCust.*, tCust_1.WorkName AS MasterWorkName  
FROM         dbo.tCust LEFT OUTER JOIN  
                      dbo.tCust tCust_1 ON dbo.tCust.Branch = tCust_1.Branch AND dbo.tCust.MasterCode = tCust_1.Code  ) tcust LEFT OUTER JOIN   --Where dbo. tcust.Branch = dbo.Get_Current_Branch()  
                       (SELECT     sumPrice , Customer  
                        FROM         dbo.tFacM  
                        WHERE     dbo.tFacm.Facpayment = 0 and dbo.tFacm.Branch = dbo.Get_Current_Branch()) T ON tcust.code = T.Customer  
--WHERE tcust.[Branch] = dbo.[Get_Current_Branch]()

GROUP BY   
 tcust.[Code], tcust.[MasterCode], tcust.[Owner],   
 --tcust.[City],tcust.[Sex],  
 tcust.[ActKind],tcust.[ActDeAct],   
 tcust.[Prefix], tcust.[Assansor],   
 tcust.[Address], tcust.[PostalCode], tcust.[Tel1],  
 tcust.[Tel2], tcust.[Tel3], tcust.[Tel4], tcust.[Mobile], Tafsili , 
 --tcust.[Fax], tcust.[Email], tcust.[CarryFee],   
 --tcust.[PaykFee],tcust.[Distance], tcust.[Credit],   
 --tcust.[Discount], tcust.[BuyState], tcust.[Description],   
 --tcust.[Date],tcust.[Time], tcust.[User], tcust.[Unit],   
 --tcust.[InternalNo], tcust.[Flour] ,  
 tcust.[workname],  
 tcust.[Name],tcust.[family],tcust.[MembershipId] ,  
 tcust.MasterWorkName , tcust.Central , tcust.FamilyNo  , [tCust].[Branch] , tcust.Credit





GO



ALTER Proc Get_All_Customers
@ActDeact INT , @Branch INT = NULL 
as

IF @Branch IS NULL SET @Branch = dbo.Get_Current_Branch()

Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust where code > 0 and actdeact <> @ActDeact --AND Branch = @Branch



GO


ALTER   PROCEDURE [dbo].[Get_SaleSummary]
(
@Branch int,
@DateBefore nvarchar(8),
@DateAfter nvarchar(8) ,
@Uid INT 
)
 AS
BEGIN
 SELECT SUM(tvw.SumPrice) AS SumPriceTotal ,
        tvw.[Date] ,
        tvw.inventoryName ,
        ISNULL(tvw.Tafsili, 0) AS Tafsili ,
        Branch ,
        inVentoryNo --, nvcDescription
 FROM   ( SELECT    dbo.tFacM.Branch ,
                    dbo.tFacD.[IntInventoryNo] AS inVentoryNo ,
                    dbo.tFacM.[No] ,
                    dbo.tFacM.[Date] ,
                    dbo.tFacM.[Time] ,
                    dbo.tFacM.[User] ,
                    CarryFeeTotal ,
                    DiscountTotal ,
                    StationID ,
                    ServiceTotal ,
                    PackingTotal ,
                    TaxTotal ,
                    DutyTotal ,
                    FacPayment ,
                    Balance ,
                    tInventory.[Description] AS inventoryName ,
                    tInventory.[Tafsili] --, dbo.tFacM.[nvcDescription]
                    ,
                    ( tfacd.Amount * tfacd.Feeunit ) AS SumPrice
          FROM      dbo.tFacM
                    INNER JOIN tfacD ON tfacm.intserialno = tfacD.intserialno
                                        AND tfacm.Branch = tfacD.Branch
                    INNER JOIN tInventory ON tInventory.inVentoryNo = tfacD.IntInventoryNo
                                             AND tInventory.Branch = tfacD.Branch
          WHERE     tfacm.Branch = @Branch
                    AND tfacm.[Date] >= @DateBefore
                    AND tfacm.[Date] <= @DateAfter
                    AND Recursive = 0
                    AND Status = 2
                    AND transferAccounting = 0 
                    AND (tfacm.[User] = @Uid OR @Uid = 0)
        ) tvw
 GROUP BY tvw.[Date] ,
        inventoryName ,
        Tafsili ,
        Branch ,
        inVentoryNo--, nvcDescription
 ORDER BY tvw.[Date] ,
        Tafsili
END

GO




ALTER PROCEDURE [dbo].[Get_SaleReturnSummary]
    (
	@Branch int,
	@DateBefore NVARCHAR(8) ,
	@DateAfter NVARCHAR(8) ,
	@Uid INT 
    )
AS 
    SELECT  dbo.tfacm.[Date] ,
            dbo.tfacm.[No] ,
            dbo.tfacm.[User] ,
            dbo.tfacm.Balance ,
            dbo.tfacm.Branch ,
            dbo.tfacm.Status ,
            dbo.tfacm.Recursive ,
            dbo.tfacm.SumPrice ,
            dbo.tfacm.Customer ,
            CASE dbo.tCust.[Name] + SPACE(3) + dbo.tCust.family
              WHEN NULL THEN dbo.tCust.WorkName
              WHEN '' THEN dbo.tCust.WorkName
              ELSE dbo.tCust.[Name] + '  ' + dbo.tCust.family
            END AS CustomerName ,
            dbo.tCust.tafsili AS CustomerTafsili
    FROM    dbo.tFacM
            INNER JOIN dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID
                                    AND dbo.tFacM.[Branch] = dbo.tUser.Branch
            INNER JOIN dbo.tPer ON tUser.ppno = tPer.ppno
                                   AND tUser.Branch = tPer.Branch
            INNER JOIN tCUst ON tfacM.Customer = tcust.code
    WHERE   ( Recursive = 0 )
            AND ( Status = 5 )
			AND dbo.tFacM.transferAccounting=0
            AND dbo.tFacM.Branch = @Branch
            AND dbo.tfacm.[Date] >= @DateBefore
            AND dbo.tfacm.[Date] <= @DateAfter
             AND (tfacm.[User] = @Uid OR @Uid = 0)
--and tfacm.Customer > 0 And Balance = 0 And Facpayment = 1


GO



ALTER   PROCEDURE [dbo].[Get_AccountDocument]
    (
      @Branch INT ,
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @Status INT ,
      @Code INT
    )
AS 
    IF ( @Code = 1 ) 
            BEGIN
                SELECT  tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
                        tPer.Tafsili AS PersonTafsili 
                       , ISNULL(SUM(ISNULL(tFacCash.intAmount , 0)), 0)  AS sp
                FROM    tFacM
                        LEFT OUTER JOIN tFacCash ON tFacM.Branch = tFacCash.Branch
                                               AND tFacM.intSerialNo = tFacCash.intSerialNo
                        INNER JOIN tUser ON tUser.UID = tFacM.[User]
                        INNER JOIN tPer ON tUser.pPno = tPer.pPno
                WHERE    tFacM.Branch = @Branch 
                        AND tFacM.Recursive = 0 
                        AND tFacM.Status =@Status
                        AND dbo.tFacM.transferAccounting=0  
                GROUP BY tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
                        tPer.Tafsili 
                HAVING  tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,
                        tPer.Tafsili
            END

    IF ( @Code = 5 ) 
            BEGIN
                SELECT  tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
                        tPer.Tafsili AS PersonTafsili ,
                        ISNULL(SUM(tFacCard.intAmount), 0) AS sp
                FROM    tUser
                        INNER JOIN tPer ON tUser.pPno = tPer.pPno
                        INNER JOIN tFacM ON tUser.UID = tFacM.[User]
                                           -- AND tUser.Branch = tFacM.Branch
                        INNER JOIN tFacCard ON tFacM.Branch = tFacCard.Branch
                                               AND tFacM.intSerialNo = tFacCard.intSerialNo
                WHERE    tFacM.Branch = @Branch 
                        AND  tFacM.Recursive = 0 
                        AND ( tFacM.Status = @Status
                              OR tFacM.Status = 8
                            )AND dbo.tFacM.transferAccounting=0
                GROUP BY tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
                        tPer.Tafsili
                HAVING   tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,
                        tPer.Tafsili
            END



GO




ALTER  Proc Get_New_FacM_No ( @Status int, @AccountYear smallint, @Branch INT )


as
DECLARE @No INT 
DECLARE @TempNo INT 
DECLARE @ShiftNo INT 
--DECLARE @Date NVARCHAR(8)  ---problem with miladi
	SET @ShiftNo= dbo.Get_Shift(GETDATE())     
	--SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      
 
	set @No = (Select isnull(max([No]),0)+ 1 as No From tFacM  Where  Status = @Status and Branch =  @Branch AND AccountYear = @AccountYear)
	set @TempNo = (Select isnull(max([TempNo]),0)+ 1 as No From tFacM  Where  Status = @Status and Branch =  @Branch AND Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())  AND shiftNo = @ShiftNo)

SELECT @No AS No , @TempNo AS TempNo


GO


--افزایش سرعت  گزارشات گروهی کالا و ریز کالا  

-------------------------------------------- 'Recipt گزارش فروش ريزكالا وتخفيفات'----------------------------------------------------

ALTER  PROCEDURE [dbo].[GetSellKindInfo]
    (
      --@intLanguage INT = 0 ,
      @SystemDate NVARCHAR(50) ,
      @SystemDay NVARCHAR(50) ,
      @SystemTime NVARCHAR(50) ,
      @Date1 VARCHAR(50) ,
      @Date2 VARCHAR(50) ,
      @User1 INT ,
      @User2 INT ,
      @station1 INT ,
      @station2 INT ,
      @Time1 NVARCHAR(50) ,
      @Time2 NVARCHAR(50) ,
      @level11 INT ,
      @level12 INT ,
      @level21 INT ,
      @level22 INT ,
      @Branch1 INT ,
      @Branch2 INT ,     
      @Status1 INT 

    )
AS 

	DECLARE @intLanguage INT
		SET @intLanguage = 0   
    DECLARE @tmp1 INT
    DECLARE @tmp2 NVARCHAR(50)
    DECLARE @ServiceTotal INT
    DECLARE @DiscountTotal INT
    DECLARE @CarryFeeTotal INT
    DECLARE @PackingTotal INT
    DECLARE @TaxTotal INT
    DECLARE @DutyTotal INT
    DECLARE @Time3 NVARCHAR(50)
    DECLARE @Time4 NVARCHAR(50)
    SET @Time3 = @Time1
    SET @Time4 = @tmp2

    DECLARE @SumUnbalanceFich BIGINT
    DECLARE @SumPayment BIGINT
    DECLARE @SumManualReceived BIGINT
    DECLARE @SumCashReceived BIGINT
    DECLARE @SumCardReceived BIGINT
    DECLARE @SumBonReceived BIGINT
    DECLARE @SumChequeReceived BIGINT
    DECLARE @SumCustomerDebit BIGINT
    DECLARE @SumGarsonDebit BIGINT
    DECLARE @SumRoundDiscount BIGINT
    DECLARE @SumOrderPrice BIGINT
    DECLARE @SumOrderReceived BIGINT
	
    IF @User2 < @User1 
        BEGIN 
            SET @tmp1 = @User2
            SET @User2 = @User1
            SET @User1 = @tmp1	
        END	

    IF @Time2 < @Time1 
        BEGIN
		/*SET @tmp2 = @Time2
		SET @Time2 = @Time1
		SET @Time1 = @tmp2*/
            SET @Time3 = '00:00'
            SET @Time4 = '24:00'
        END
	
     SELECT   @TaxTotal = SUM(TaxTotal) ,
			  @DutyTotal = SUM(DutyTotal) ,
			  @PackingTotal = SUM(PackingTotal) ,
			  @CarryFeeTotal = SUM(CarryFeeTotal) ,
			  @DiscountTotal = SUM(DiscountTotal) ,
			  @ServiceTotal =  SUM(ServiceTotal)
                          FROM      tfacm
                          WHERE     [date] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = @Status1
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                         
   SET @SumUnbalanceFich = ( SELECT    SUM(SumPrice)
                          FROM      tfacm
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                                    AND Balance = 0 AND FacPayment = 0 AND InCharge IS NULL 
                        ) 
    SET @SumPayment = ( SELECT    SUM(Bestankar)
                          FROM      dbo.tblAcc_Cash
                          WHERE     [date] >= @Date1
                                    AND [date] <= @Date2
                                    --AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND UID >= @User1 AND UID <= @User2
                        ) 
    SET @SumManualReceived = ( SELECT    SUM(Bestankar)
                          FROM      dbo.tblAcc_Recieved
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND UID >= @User1 AND UID <= @User2
                                    AND intSerialNo IS NULL 
                        ) 
    SET @SumCashReceived = ( SELECT    SUM(intAmount)
                          FROM      tfacm
                          INNER JOIN tfaccash ON dbo.tFacM.Branch = dbo.tFacCash.Branch AND dbo.tFacM.intSerialNo = dbo.tFacCash.intSerialNo
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                        ) 
    SET @SumCardReceived = ( SELECT    SUM(intAmount)
                          FROM      tfacm
                          INNER JOIN dbo.tFacCard ON dbo.tFacM.Branch = dbo.tFacCard.Branch AND dbo.tFacM.intSerialNo = dbo.tFacCard.intSerialNo
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                        ) 
    SET @SumBonReceived = ( SELECT    SUM(intAmount)
                          FROM      tfacm
                          INNER JOIN dbo.tFacCredit ON dbo.tFacM.Branch = dbo.tFacCredit.Branch AND dbo.tFacM.intSerialNo = dbo.tFacCredit.intSerialNo
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                        ) 
    SET @SumChequeReceived = ( SELECT    SUM(intChequeAmount)
                          FROM      tfacm
                          INNER JOIN dbo.tFacCheque ON dbo.tFacM.Branch = dbo.tFacCheque.Branch AND dbo.tFacM.intSerialNo = dbo.tFacCheque.intSerialNo
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                        ) 
    SET @SumCustomerDebit = ( SELECT    SUM(SumPrice)
                          FROM      tfacm
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                                    AND Customer > 0 AND Balance =0 AND FacPayment = 1
                        ) 
   SET @SumGarsonDebit = ( SELECT    SUM(SumPrice)
                          FROM      tfacm
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                                    AND incharge IS NOT NULL  AND Balance =0 AND FacPayment = 0
                        ) 
    SET @SumRoundDiscount = ( SELECT    SUM(RoundDiscount)
                          FROM      tfacm
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 2
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                        ) 
    SET @SumOrderPrice = ( SELECT    SUM(SumPrice)
                          FROM      tfacm
                          WHERE     [date ] >= @Date1
                                    AND [date] <= @Date2
                                    AND Status = 10
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND Branch >= @Branch1
                                    AND Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                        ) 
    SET @SumOrderReceived = ( SELECT    SUM(DutyTotal)
                          FROM      tfacm
			  			  INNER JOIN dbo.tblAcc_Recieved ON dbo.tblAcc_Recieved.Branch = dbo.tFacM.Branch AND dbo.tblAcc_Recieved.intSerialNo = dbo.tFacM.intSerialNo
                          WHERE     tfacm.[date ] >= @Date1
                                    AND tfacm.[date] <= @Date2
                                    AND Status = 10
                                    AND ( ( [Time] >= @Time1
                                            AND [Time] <= @Time4
                                          )
                                          OR ( [Time] <= @Time2
                                               AND [Time] >= @Time3
                                             )
                                        )
                                    AND StationID >= @station1
                                    AND StationID <= @station2
                                    AND Recursive = 0
                                    AND tfacm.Branch >= @Branch1
                                    AND tfacm.Branch <= @Branch2
                                    AND [User] >= @User1 AND [User] <= @User2
                                    AND dbo.tblAcc_Recieved.intSerialNo IS NOT NULL 
                        ) 

    DECLARE @TimeTitle NVARCHAR(10)
    IF @intLanguage = 0 
        SET @TimeTitle = N' ساعت : '
    ELSE 
        SET @TimeTitle = N'Time: '

DECLARE @BranchName1 NVARCHAR(20)
DECLARE @BranchName2 NVARCHAR(20)
SELECT @BranchName1 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch1
SELECT @BranchName2 = nvcBranchName FROM dbo.tBranch WHERE dbo.tBranch.Branch = @Branch2
    SELECT  SUM(dbo.ViewRepSellKind.Amount) AS SumAmount ,
            dbo.ViewRepSellKind.FeeUnit ,
            dbo.ViewRepSellKind.FeeUnit * SUM(dbo.ViewRepSellKind.Amount) AS PriceTotal ,
            dbo.ViewRepSellKind.Level1Code ,
            dbo.ViewRepSellKind.Level2Code ,
            dbo.ViewRepSellKind.UnitGoodDescription ,
            @Date1 AS DateBefore ,
            @Date2 AS DateAfter ,
            dbo.ViewRepSellKind.GoodCode ,
            dbo.ViewRepSellKind.Barcode ,
            @User1 AS FromUser ,
            @User2 AS ToUser ,
            @Time1 AS FromTime ,
            @Time2 AS ToTime ,
            @SystemDay + ' ' + @SystemDate + @TimeTitle + @SystemTime AS Sysdate ,
            @level11 AS FromGoodCodeLvele1 ,
            @level21 AS FromGoodCodeLvele2 ,
            @level12 AS ToGoodCodeLevel1 ,
            @level22 AS ToGoodCodeLevel2 ,
            CASE @intLanguage
              WHEN 0 THEN dbo.ViewRepSellKind.NamePrn
              WHEN 1 THEN dbo.ViewRepSellKind.LatinNamePrn
            END AS GoodName ,
            CASE @intLanguage
              WHEN 0 THEN dbo.ViewRepSellKind.Level1Desc
              WHEN 1 THEN dbo.ViewRepSellKind.Level1LatinDesc
            END AS Level1Description ,
            CASE @intLanguage
              WHEN 0 THEN dbo.ViewRepSellKind.Level2Desc
              WHEN 1 THEN dbo.ViewRepSellKind.Level2LatinDesc
            END AS Level2Description ,
            dbo.ViewRepSellKind.DiscountRate ,
            SUM(dbo.ViewRepSellKind.Discount) AS SumDiscount ,
            dbo.ViewRepSellKind.Rate ,
            dbo.ViewRepSellKind.nvcBranchName ,
            @ServiceTotal AS ServiceTotal ,
            @DiscountTotal AS DiscountTotal ,
            @CarryFeeTotal AS CarryFeeTotal ,
            @PackingTotal AS PackingTotal ,
            @Status1 AS Status ,
            ViewRepSellKind.AccountYear ,
            ViewRepSellKind.Branch 
          --,dbo.ViewRepSellKind.ServePlace
	    , @taxTotal AS TaxTotal , @DutyTotal AS DutyTotal ,'' AS [serveplacedesc]
    -- @SumUnbalanceFich AS SumUnbalanceFich,
--      @SumPayment AS SumPayment,
--      @SumManualReceived AS SumManualReceived,
--      @SumCashReceived AS SumCashReceived,
--      @SumCardReceived AS SumCardReceived,
--      @SumBonReceived AS SumBonReceived ,
--      @SumChequeReceived AS SumChequeReceived,
--      @SumCustomerDebit AS SumCustomerDebit,
--      @SumGarsonDebit AS SumGarsonDebit,
--      @SumRoundDiscount AS SumRoundDiscount,
--      @SumOrderPrice AS SumOrderPrice,
--      @SumOrderReceived AS SumOrderReceived
		, @BranchName1 AS BranchName1 , @BranchName2 AS BranchName2 , Branch 
     FROM    dbo.ViewRepSellKind
    WHERE   dbo.ViewRepSellKind.[date] >= @Date1
            AND dbo.ViewRepSellKind.[date] <= @Date2
	--AND dbo.ViewRepSellKind.[Supplier] IN (SELECT code FROM tSupplier WHERE code  >= @Sup1 AND code < =@Sup2)
            AND dbo.ViewRepSellKind.[User] >= @User1
                  AND dbo.ViewRepSellKind.[User] <= @User2

            AND ( ( dbo.ViewRepSellKind.[Time] >= @Time1
                    AND dbo.ViewRepSellKind.[Time] <= @Time4
                  )
                  OR ( dbo.ViewRepSellKind.[Time] <= @Time2
                       AND dbo.ViewRepSellKind.[Time] >= @Time3
                     )
                )
            AND dbo.ViewRepSellKind.Status = @Status1
            AND dbo.ViewRepSellKind.StationID >= @station1
            AND dbo.ViewRepSellKind.StationID <= @station2
            AND dbo.ViewRepSellKind.Level1Code >= @level11
            AND dbo.ViewRepSellKind.Level1Code <= @level12
            AND dbo.ViewRepSellKind.Level2Code >= @level21
            AND dbo.ViewRepSellKind.Level2Code <= @level22
            AND dbo.ViewRepSellKind.Branch >= @Branch1
            AND dbo.ViewRepSellKind.Branch <= @Branch2
		--	AND dbo.ViewRepSellKind.Balance = 1
    GROUP BY --dbo.ViewRepSellKind.ServePlace ,
            dbo.ViewRepSellKind.GoodCode ,
            dbo.ViewRepSellKind.UnitGoodDescription ,
            dbo.ViewRepSellKind.FeeUnit ,
            dbo.ViewRepSellKind.Level2LatinDesc ,
            dbo.ViewRepSellKind.Level2Desc ,
            dbo.ViewRepSellKind.Level1LatinDesc ,
            dbo.ViewRepSellKind.Level1Desc ,
            dbo.ViewRepSellKind.NamePrn ,
            dbo.ViewRepSellKind.LatinNamePrn ,
            dbo.ViewRepSellKind.Level1Code ,
            dbo.ViewRepSellKind.Level2Code ,
            dbo.ViewRepSellKind.Barcode ,
            dbo.ViewRepSellKind.DiscountRate ,
            dbo.ViewRepSellKind.Rate ,
            ViewRepSellKind.AccountYear ,
            ViewRepSellKind.Branch ,
            ViewRepSellKind.nvcBranchName
	        --,ViewRepSellKind.[serveplacedesc]
    ORDER BY dbo.ViewRepSellKind.GoodCode





GO



