

--Script_V26_16_Fix13_OutDeliveryService
--برای V3.33  تبلت سمر 
-- اگر محل سرو بیرون بر یا دلیوری باشد سرویس را صفر در نظر می گیرد

--94/061/13

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER   PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

            @Status INT ,      
            @Owner INT ,      
            @Customer INT ,      
            @DiscountTotal FLOAT ,      
            @CarryFeeTotal FLOAT ,      
            @Recursive INT ,      
            @InCharge INT ,      
            @FacPayment BIT ,      
            @OrderType INT ,      
            @StationId INT ,      
            @ServiceTotal FLOAT ,      
            @PackingTotal FLOAT ,      
            @TableNo INT ,      
            @User INT ,      
            @Date NVARCHAR(50) ,      
            @DetailsString nText,      
            @ds nText = '',      
            @Balance BIT ,      
            @AccountYear smallint = null  ,       
            @NvcDescription Nvarchar(150) = Null ,      
            @HavaleNo int = Null  ,      
            @TempAddress Nvarchar(255) = '',  
			@GuestNo INT,    
            @lastFacMNo INT OUT  ,
		    @Person INT = NULL     
             )      

AS      

Declare @intserialNo int      
Declare @intserialNo2 int      
--Declare @intserialNo3 Bigint    

SET @intserialNo = 0        
SET @intserialNo2   = 0      
--SET @intserialNo3   = 0      

DECLARE @No1  INT     
DECLARE @No2  INT     
--DECLARE @No3  INT     

DECLARE @SumPrice  BIGINT       
Set @SumPrice = 0      

DECLARE @proper_time nvarchar(5)      

IF @AccountYear IS  Null
    Set @AccountYear = dbo.get_AccountYear() 
    
IF  @Owner = 0      
    SET @Owner = NULL      

IF  @TableNo < 1      
    SET @TableNo = NULL      

IF  @Incharge < 1      
    SET @Incharge = NULL      

IF  @Customer=0      
    SET @Customer = NULL      

BEGIN TRAN      

    DECLARE @MasterServePlace INT      
    DECLARE @newtime nvarchar(5)      
    select @newtime=dbo.setTimeFormat(getdate())      
    SELECT @MasterServePlace = SUM(tmpTable.SServePlace)      
    FROM (  SELECT DISTINCT ServePlace As SServePlace      
         FROM Split(@DetailsString)      
           ) tmpTable      

----------------------------------------Date From Server-----------------------------------------------------------------      
If @Status = 2 And dbo.Get_DateFromServer() = 1      
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())      
ELSE
	IF LEN(@Date) < 8
		SET @Date = dbo.fnFixDateString(@Date) ------For Check Date String in Valid Format YY/MM/DD-----


------Start New Line For Avoid Repeat in tFacm------
DECLARE @RepeatNo INT

SELECT @RepeatNo = COUNT(*) FROM dbo.tFacM WHERE Status = 2 AND [Date] = @Date 
    AND TableNo = @TableNo AND Recursive = 0 AND FacPayment = 0 --AND Balance = 0

IF @RepeatNo > 0 
    GOTO EventHandler

----End New Line -----------------------------------------------------------------------------------------------      

 Declare @intBranch  int      
 Declare @ShiftNo int      
 DECLARE @TempNo INT 

 SELECT @intBranch = dbo.Get_Current_Branch()
 
 --select @intBranch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@DetailsString))      
 --IF @intBranch = 0 OR @intBranch IS NULL     SET @intBranch = dbo.Get_Current_Branch()

    DECLARE @IdentityNo INT
    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
    FROM    tFacm
    WHERE   Branch = @intBranch 

    IF @IdentityNo < ( @intBranch * 10000000 ) 
        SET @IdentityNo = ( @intBranch * 10000000 )

 SET @NO1 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND AccountYear = @AccountYear)      

 SET @ShiftNo= dbo.Get_Shift(GETDATE())      
 SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      

IF COL_LENGTH('[tFacM]','ServePlaceTempNo') IS NULL
	ALTER TABLE dbo.tFacM  ADD ServePlaceTempNo INT NULL 

DECLARE @ServePlaceTempNo INT 
 SET @ServePlaceTempNo = (SELECT ISNULL(MAX(ServePlaceTempNo),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ServePlace = @MasterServePlace)      


     INSERT INTO tFacM (   
		intSerialNo ,   
		[No] ,      
		[Date] ,      
		RegDate ,      
		Status ,      
		Customer ,      
		SumPrice ,      
		OrderType ,      
		ServePlace ,      
		StationId ,      
		ServiceTotal ,      
		Recursive ,      
		CarryFeeTotal ,      
		PackingTotal ,      
		DiscountTotal ,      
		[Time] ,      
		[User] ,      
		TableNo ,      
		shiftNo ,      
		incharge,      
		owner ,      
		FacPayment ,       
		Balance ,       
		Branch,      
		AccountYear ,      
		NvcDescription,      
		TempAddress ,
		GuestNo ,
		TempNo ,
		ServePlaceTempNo    
		
 )      
     Values       

(	    @IdentityNo ,  
        @NO1 ,      
        @Date ,      
        dbo.Shamsi(GETDATE()) ,      
        @Status,      
        @Customer ,      
        @SumPrice ,      
        @OrderType ,      
        @MasterServePlace ,      
        @StationId ,      
        @ServiceTotal ,      
        @Recursive ,      
        @CarryFeeTotal ,      
        @PackingTotal ,      
        @DiscountTotal ,      
        @newtime,      
        @User ,      
        @TableNo,      
        @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
        @Incharge ,      
        @owner ,      
        @FacPayment ,      
        @Balance ,      
		@intBranch , --dbo.Get_Current_Branch(),      
		@AccountYear ,      
		@NvcDescription,      
		@TempAddress,
		@GuestNo,
		@TempNo ,
		@ServePlaceTempNo  
 )      
     IF @@ERROR <>0      
        GoTo EventHandler       

    SET @intserialNo = @IdentityNo

declare @destbranch  INT 
SET @destbranch = 0
DECLARE @TempNo2 INT 
DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      
 
If @Status = 6 AND @DestinventoryNo > 0  -- And (@destbranch= @intBranch Or dbo.AutoResid() = 1)    
	Begin      
	select @destbranch=  @intBranch --   branch from tInventory where inventoryNo=(SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

	  SET @NO2 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @destbranch AND AccountYear = @AccountYear)      
	  --SET @TempNo2 = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      

     INSERT INTO tFacM ( 
				intSerialNo ,     
                [No] ,      
                [Date] ,      
                RegDate ,      
                Status ,      
                Customer ,      
                SumPrice ,      
                OrderType ,      
                ServePlace ,      
                StationId ,      
                ServiceTotal ,      
                Recursive ,      
                CarryFeeTotal ,      
                PackingTotal ,      
                DiscountTotal ,      
                TaxTotal ,
                DutyTotal ,     
                [Time] ,      
                [User] ,      
                TableNo ,      
                shiftNo ,      
                incharge,      
                owner ,      
                FacPayment ,       
                Balance ,       
                Branch,      
			  AccountYear ,      
			  NvcDescription,      
			  TempAddress,
			  GuestNo ,
			  TempNO     

 )      
     Values      
(				@IdentityNo + 1 ,     
                @NO2 ,      
                @Date ,      
                dbo.Shamsi(GETDATE()) ,      
                7,      
                @Customer ,      
                @SumPrice ,      
                @OrderType ,      
                @MasterServePlace ,      
                @StationId ,      
                @ServiceTotal ,      
                @Recursive ,      
                @CarryFeeTotal ,      
                @PackingTotal ,      
                @DiscountTotal ,      
                0 ,
                0 ,      
                @newtime,      
                @User ,      
                @TableNo,      
                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
                @Incharge ,      
                @owner ,      
                @FacPayment ,      
                @Balance ,      
				@DestBranch ,     
				@AccountYear ,      
				@NvcDescription,      
				@TempAddress,
				@GuestNo ,
				NULL --@TempNo2    
		
 )      
		 IF @@ERROR <>0      
			GoTo EventHandler      
		SET @intserialNo2 = @IdentityNo + 1      

            UPDATE  tfacm
            SET     NvcDescription = @NvcDescription + N' رسيد -   '
                    + CAST(@No2 AS NVARCHAR(8))
            WHERE   intSerialNo = @intserialNo
            UPDATE  tfacm
            SET     RefrenceHavale = @intserialNo2
            WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch

end      


----------------------------------Fill Details Factor  --------------------------------------------------------------      
If @Status = 6 AND @DestinventoryNo > 0 -- AND (@destbranch= @intBranch  Or dbo.AutoResid() = 1)        
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch      
Else       
 exec InsertFactorDetail @DetailsString , @intserialNo , 0, @Customer , @intBranch      

     IF @@ERROR <>0      
        GoTo EventHandler      
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------      

----------------------------------Total SumPrice Calculate  --------------------------------------------------------------      
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(SUM(ROUND( (Amount * FeeUnit * discount/100 ) ,0) ) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(SUM( ROUND((Amount * FeeUnit) * (1 - discount/100),0) )  AS BIGINT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

     IF @@ERROR <>0      
        GoTo EventHandler      
DECLARE @TaxTotal FLOAT  
SET @TaxTotal = 0
DECLARE @ValueGoodsTax FLOAT
SET @ValueGoodsTax = 0
IF @status = 2 OR @status = 5  OR @status = 10
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxBuy = 1
	            )  
SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
IF @@ERROR <> 0 
GOTO EventHandler

DECLARE @DutyTotal INT 
SET @DutyTotal = 0
DECLARE @ValueGoodsDuty FLOAT
SET @ValueGoodsDuty = 0
IF @status = 2 OR @status = 5  OR @status = 10
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutyBuy = 1
	            )  
SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
IF @@ERROR <> 0 
GOTO EventHandler

Declare @SumPrice2 Bigint      
Set @SumPrice2 = (Select Cast(Sum(Amount * FeeUnit) as Bigint) From tFacd Where intSerialNo = @intserialNo2 And Branch = @DestBranch )        
     IF @@ERROR <>0      
        GoTo EventHandler      
----------------------------------ServiceRate Calculate  --------------------------------------------------------------      
Declare @ReserveServiceRate Int      
Set @ReserveServiceRate = 0      

If  @TableNo >0      
Begin      
	Declare @Reserve Bit      
	Set @Reserve = (Select Reserve From tTable Where tTable.[No] = @TableNo)      
	If @Reserve = 1      
		Set @ReserveServiceRate = (Select t.ReserveServiceRate From(Select ReserveServiceRate From tPartitions  inner join tTable        
		On tPartitions.PartitionId = tTable.PartitionId and  tTable.[No] = @TableNo)t )      

        Update dbo.tTable      
           Set   dbo.tTable.Empty  = 0      
                Where dbo.tTable.[No] = @TableNo AND  @Balance = 0    
	If dbo.Get_TableMonitoring() = 1   ---Table Monitoring      
	 	Begin      
		DECLARE @intTableUsedNo INT      
--		SELECT @intTableUsedNo=intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
--		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@TableNo and vwSamar_TableUsage_BusyTable.intBranch=@intBranch      
		DECLARE @nvcString NVARCHAR(100)      
		SET @nvcString=','+CAST(@TableNo AS NVARCHAR(5))+'/'      
		--IF @intTableUsedNo is NULL      
		EXEC insert_tblSamar_TableUsage @nvcString,1      
--		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.nvcStartTime=  @newtime      
--		FROM    ( SELECT     dbo.vwSamar_TableUsage_BusyTable.intTableUsedNo, dbo.vwSamar_TableUsage_BusyTable.nvcStartTime,       
--		                 dbo.vwSamar_TableUsage_BusyTable.intBranch, dbo.tTable.[No]      
--				FROM         dbo.tTable LEFT OUTER JOIN      
--		                 dbo.vwSamar_TableUsage_BusyTable ON dbo.vwSamar_TableUsage_BusyTable.intTableNo = dbo.tTable.[No] AND       
--		                 dbo.vwSamar_TableUsage_BusyTable.intBranch = dbo.tTable.Branch)t      
--		WHERE  tblSamar_TableUsage.intTableNo=t.[No] and tblSamar_TableUsage.intBranch=t.intBranch      
--		and tblSamar_TableUsage.intTableNo=@TableNo and tblSamar_TableUsage.intBranch= @intBranch     
		END        
End      
     IF @@ERROR <>0      
        GoTo EventHandler      


If @ReserveServiceRate > 0       
 Set @ServiceTotal = @ReserveServiceRate      

-- ===================For Calculate Service In Delivery Or Out ==================
	IF @MasterServePlace = 2 OR @MasterServePlace = 4 
		SET @ServiceTotal = 0

-------------------------------------

 If @ServiceTotal <> 0      
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)      
     IF @@ERROR <>0      
        GoTo EventHandler       
----------------------------------Round Sumprice  --------------------------------------------------------------      
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5  OR @status = 10
 BEGIN 
  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
 END 
    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal     

    Declare @Remain INT
    SET @Remain = 0  
    IF @Status = 2 OR @Status = 10
    BEGIN   
    Set @Remain = dbo.RoundSumPrice(@SumPrice )         
    Set @SumPrice = @SumPrice - @Remain      
    Set @DiscountTotal = @DiscountTotal + @Remain    
    END  
---select @Remain as remain      
----------------------------------Calculate Packing---------------------------------------------------------------      
If dbo.Get_AutoPacking() = 1      
Begin      
    Declare @UserPacking INT      
    SET @UserPacking = ISNULL((Select ISNULL(Sum(Amount),0)  From tGood inner join tFacD On tFacD.GoodCode = tGood.Code       
        where intSerialNo = @intSerialNo and MainType =1 and Branch = @intBranch Group By MainType) ,0)      
    Set @UserPacking = @UserPacking * [dbo].[Get_User_Packing]()      
    Set @SumPrice = @SumPrice + @UserPacking      
    Update tFacm Set PackingTotal = @UserPacking  Where intSerialNo = @intserialNo And Branch = @intBranch       
End      
----------------------------------Net Price Update  --------------------------------------------------------------      

Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch      
Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DiscountTotal = @DiscountTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

If @Status = 6 AND @DestinventoryNo > 0-- AND (@destbranch= @intBranch )  -- Or dbo.AutoResid() = 1   
	Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch       
      IF @@ERROR <>0       

        GoTo EventHandler           
-------------------------------------Fill Detail Cash , ..........--------------------------------------------------      
DECLARE @Result INT 
IF (@Status =  1 OR @Status = 2 )      
	 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds ,  @intBranch  , @Remain ,1  , @Result OUT   

     IF @@ERROR <>0      
   GoTo EventHandler      
IF @Result = -1
   GoTo EventHandler      

-------------------------------------Monitoring---------------------------------------------------------------------      
--Declare  @Monitor1 int      
--Declare  @Monitor2 int       

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  @intBranch)      
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  @intBranch)      


--IF @Monitor1 > 0       
--   exec Notify_to_Clients      

--Else If @Monitor2 > 0       
--   exec Notify_to_Clients      

----------------------------History---------------------------      

Exec InsertHistory  @No1, @Status , @User , 1 , @AccountYear , @intBranch      
     IF @@ERROR <>0      
   GoTo EventHandler      
IF @STATUs = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 7 , @User , 1 , @AccountYear , @destbranch      
     IF @@ERROR <>0      
   GoTo EventHandler      


----------------------------Cash ---------------------------      

------------------------Mojodi Control Online--------------------------------------------      

Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @intBranch      
IF @@ERROR <>0      
 GoTo EventHandler      
IF @STATUS = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1      
 BEGIN      
 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch      
 IF @@ERROR <>0      
 GoTo EventHandler      
 Exec InsertMojodiCalculate  7,  @intserialNo2 , @AccountYear , @destbranch      
 IF @@ERROR <>0      
 GoTo EventHandler      
 END       

------------------------------------------Update Balance After Recived----------------------------
Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch


COMMIT TRAN

--DECLARE @TemporaryNo BIT 
--SELECT @TemporaryNo = TemporaryNo FROM dbo.tStations WHERE StationID = @StationId AND Branch = @intBranch
--IF @TemporaryNo = 0 set @lastFacMNo = @No1
--ELSE set @lastFacMNo = @TempNo

set @lastFacMNo = @intserialNo


---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @lastFacMNo , 1

--------------------------------------------------------------------------------------------------------------------------------------


Return @lastFacMNo      

EventHandler:      

    ROLLBACK TRAN      
    SET @LastFacMNo = -1      

    RETURN @lastFacMNo
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--تغییر تاریخ فاکتور خرید و عدم تغییر تاریخ فاکتور فروش
ALTER     PROCEDURE [dbo].[EditFactorMasterDetails]  (  


	@No       INT,  
	@Status  INT ,  
	@Owner  INT ,  
	@Customer  INT ,  
	@DiscountTotal Float ,  
	@CarryFeeTotal Float ,  
	@Recursive  INT ,  
	@InCharge  INT ,  
	@FacPayment  BIT ,  
	@OrderType  INT ,  
	@StationId  INT ,  
	@ServiceTotal  Float ,  
	@PackingTotal  Float ,  
	@TableNo  INT ,  
	@User INT ,  
	@Date   Nvarchar(50) =NULL,  
	@DetailsString  nText,  
	@ds nText = '',  
	@Balance Bit,  
	@AccountYear Smallint = Null ,  
	@NvcDescription Nvarchar(150) = Null ,  
	@TempAddress Nvarchar(255) = '', 
	@GuestNo INT,     
	@LastFacMNo  INT OUT  ,
	@Person INT = NULL 
  )  


AS  
DECLARE @SumPrice BIGINT  
DECLARE @SumPrice2 BIGINT  
DECLARE @intSerialNo BIGINT  
DECLARE @intSerialNo2 BIGINT  
--DECLARE @intSerialNo3 BIGINT  
DECLARE @OldRegDate Nvarchar(50)  
DECLARE  @FactorSerial BIGINT  

SET @Sumprice = 0  
SET @Sumprice2= 0  
SET @intSerialNo = 0  
SET @intSerialNo2 = 0  
--SET @intserialNo3 = 0  


 Declare @intBranch  int  
 Declare @ShiftNo int  

 Declare @DestBranch INT  
 SET @DestBranch = 0

IF @AccountYear IS  Null
    Set @AccountYear = dbo.get_AccountYear() 

 SELECT @intBranch = dbo.Get_Current_Branch()

-- select @intBranch = branch from tInventory where inventoryNo=(SELECT TOP 1  IntInventoryNo FROM Split(@DetailsString))  
 SET @ShiftNo= dbo.Get_Shift(GETDATE())  

SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  

--Control is difficult
--If No received then Bypass received
--DECLARE @DestinventoryNo INT 
--select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

IF @Status = 6  
	SET @intSerialNo2 = (SELECT ISNULL(tFacM.RefrenceHavale ,0) FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  


if @status=10   
set @OldRegDate = (SELECT tFacM.regdate FROM tFacM WHERE [No] = @No AND Status = @Status and Branch =  @intBranch AND AccountYear = @AccountYear)  
else set @OldRegDate=dbo.Shamsi(GETDATE())  
-------------No Change StationId , If this Fich Is For Pocket Pc---------------------------------------  
DECLARE @OldStationId INT  
 SET @OldStationId = (Select StationId From tFacm Where intserialNo = @intSerialNo and Branch =  dbo.Get_Current_Branch())  

DECLARE @StationType INT  
 SET @StationType = (Select StationType From tStations Where StationId = @OldStationId and Branch =  dbo.Get_Current_Branch())  
If  @StationType = 8  
 SET @StationId = @OldStationId  
----------------------------------------------------------------------------------------------------------  
IF  @Owner = 0  
    SET @Owner = NULL  

IF  @TableNo < 1  
    SET @TableNo = NULL  

Declare @OldTableNo   int  

SET  @OldTableNo =  IsNull((SELECT tFacM.TableNo FROM tFacM WHERE intSerialNo = @intSerialNo and Branch = dbo.Get_Current_Branch()) , 0)  

IF  @Incharge < 1  
    SET @Incharge = NULL  

IF  @Customer=0  
    SET @Customer = NULL  
IF @Date IS NULL  
 SET @Date=Rtrim(LTRIM(dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())))  

BEGIN TRANSACTION  

If IsNull(@TableNo , 0) <> @OldTableNo  
BEGIN  
 IF @OldTableNo > 0   
	-- Add For Tablet & Ppc
	DECLARE @TableNotEmpty INT 

	SELECT @TableNotEmpty = COUNT(*) FROM dbo.tFacM WHERE Status = 2 AND [Date] = @Date 
	  --AND [Time] <= @NewTime AND [Time] >= CONVERT(VARCHAR(5),@d1,108) 
	  AND TableNo = @TableNo AND Recursive = 0 AND FacPayment = 0 --AND Balance = 0

		IF @TableNotEmpty > 0 
			GOTO EventHandler

	 Update ttable SET Empty = 1 where No = @OldTableNo  
END  

    DECLARE @MasterServePlace INT  

 SELECT @MasterServePlace = SUM(tmpTable.SServePlace)  
 FROM   
 (  SELECT DISTINCT ServePlace As SServePlace  FROM Split(@DetailsString)) tmpTable  


 if @Status = 2  
 begin  
       INSERT INTO tRepFacEditM (Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance , OrderType,  
                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate, AccountYear , TaxTotal , DutyTotal  )  
          SELECT Branch , intSerialNo, [No], Status, Owner, Customer, DiscountTotal, CarryFeeTotal, SumPrice, Recursive, InCharge, FacPayment, Balance, OrderType,  
                        ServePlace, StationId, ServiceTotal, PackingTotal, ShiftNo, TableNo, [Date], [Time], [User], RegDate , AccountYear , TaxTotal , DutyTotal    
   FROM tFacM WHERE tFacM.intSerialNo = @intSerialNo and Branch = @intBranch  

      IF @@ERROR <>0  
          GoTo EventHandler  

      INSERT INTO tFacD2(Code , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate], intInventoryNo )   
    SELECT @@identity , Branch ,intRow, Amount, GoodCode, FeeUnit, intSerialNo, ServePlace,ChairName , DifferencesCodes , DifferencesDescription ,Discount ,[ExpireDate],intInventoryNo  
                 From tFacD  
                 WHERE intSerialNo = @intSerialNo  And Branch = @intBranch

      IF @@ERROR <>0  
          GoTo EventHandler  

 end  



If @status = 6 AND @intSerialNo2 > 0 --And (@destbranch = dbo.Get_Current_Branch()  )--or dbo.AutoResid() = 1   
 select @destbranch=branch from tInventory where inventoryNo=(SELECT TOP 1 DestInventoryNo FROM Split(@DetailsString))  

---------------------------------------Mojodi Control Online---------------------------------------------------------  
Exec DeleteMojodiCalculate @Status , @intserialNo  ,  1 , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = @intBranch )  --or dbo.AutoResid() = 1 
 Exec DeleteMojodiCalculate 7 , @intserialNo2  , 1 , @AccountYear , @DestBranch  
----------------------------------------Delete Old Details -----------------------------------------------------------  
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = @intBranch or dbo.AutoResid() = 1 )   
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo2 AND Branch =  @DestBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
------------------------------------------------------------    
  Exec DeleteFactorChildren @intSerialNo , @intBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
 If @status = 6 AND @intSerialNo2 > 0--And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
  Exec DeleteFactorChildren @intSerialNo2 , @DestBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
----------------------------------------Date From Server-----------------------------------------------------------------  
If @Status = 2 And dbo.Get_DateFromServer() = 1  
 SET @Date = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())  
----------------------------------------Update Master-----------------------------------------------------------------  

    Update tFacM  
        SET Owner       = @Owner,  
        Customer        = @Customer,  
        DiscountTotal   = @DiscountTotal,  
        CarryFeeTotal   = @CarryFeeTotal,  
        SumPrice        = @SumPrice,  
        Recursive       = @Recursive,   
        InCharge        = @InCharge,  
        FacPayment      = @FacPayment,  
        Balance         = @Balance,  
        OrderType       = @OrderType,  
        ServePlace      = @MasterServePlace,  
        StationId       = @StationId,  
        ServiceTotal    = @ServiceTotal,  
        PackingTotal    = @PackingTotal,  
        ShiftNo         = @ShiftNo , --dbo.Get_Shift(GETDATE()),  
        TableNo         = @TableNo,  
        [Date]          =  CASE WHEN @Status = 2 THEN  [Date] WHEN @Status = 5 THEN [Date] ELSE @Date END,  
        [Time]          = dbo.SetTimeFormat(GETDATE()),  
        [User]          = @User,  
        RegDate 	= @OldRegDate,---dbo.Shamsi(GETDATE()),  
        NvcDescription  = @NvcDescription ,  
 		TempAddress     = @TempAddress,
		GuestNo		= @GuestNo ,
		TempNo = CASE WHEN @Status = 2 THEN  TempNo WHEN @Status = 5 THEN TempNo ELSE NULL END      
    WHERE tFacM.intSerialNo = @intSerialNo  AND Branch =  @intBranch  

    IF @@ERROR <>0  
        GoTo EventHandler  

DECLARE @No2 INT
SET @No2 = 0

If @Status = 6 AND @intSerialNo2 > 0 --And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
Begin  
    SET @NO2 = (SELECT [NO] FROM tFacM WHERE intserialNo = @intSerialNo2  And Branch =  @DestBranch )      

    Update tFacM  
        SET Owner       = @Owner,  
        Customer        = @Customer,  
        DiscountTotal   = @DiscountTotal,  
        CarryFeeTotal   = @CarryFeeTotal,  
        SumPrice        = @SumPrice,  
        Recursive       = @Recursive,   
        InCharge        = @InCharge,  
        FacPayment      = @FacPayment,  
        Balance         = @Balance,  
        OrderType       = @OrderType,  
        ServePlace      = @MasterServePlace,  
        StationId       = @StationId,  
        ServiceTotal    = @ServiceTotal,  
        PackingTotal    = @PackingTotal,  
        ShiftNo         = @ShiftNo ,--dbo.Get_Shift(GETDATE()),  
        TableNo         = @TableNo,  
        [Date]          = @Date,  
        [Time]          =dbo.SetTimeFormat(GETDATE()),  
        [User]          = @User,  
        RegDate 	= dbo.Shamsi(GETDATE()),  
        NvcDescription  = @NvcDescription,  
 		TempAddress     = @TempAddress ,
		GuestNo		= @GuestNo  
    WHERE tFacM.intSerialNo = @intSerialNo2  AND Branch =  @DestBranch  

END  

----------------------------------Fill Details Factor ----------------------------------------------------------------------  
If @Status = 6  AND @intSerialNo2 > 0--AND (@destbranch= @intBranch Or dbo.AutoResid() = 1 )    
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch  
Else         
 exec InsertFactorDetail @DetailsString , @intserialNo , 0 , @Customer , @intBranch        

     IF @@ERROR <>0  
        GoTo EventHandler  
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------  


----------------------------------Total SumPrice Calculate  --------------------------------------------------------------  
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(SUM(ROUND( (Amount * FeeUnit * discount/100 ) ,0) ) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(SUM( ROUND((Amount * FeeUnit) * (1 - discount/100)  ,0)) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

     IF @@ERROR <>0      
        GoTo EventHandler      
DECLARE @TaxTotal FLOAT  
SET @TaxTotal = 0
DECLARE @ValueGoodsTax FLOAT
SET @ValueGoodsTax = 0
IF @status = 2 OR @status = 5 OR @status = 10
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxSale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsTax = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.TaxBuy = 1
	            )  
SET @ValueGoodsTax = ISNULL(@ValueGoodsTax , 0)
IF @@ERROR <> 0 
GOTO EventHandler

DECLARE @DutyTotal INT 
SET @DutyTotal = 0
DECLARE @ValueGoodsDuty FLOAT
SET @ValueGoodsDuty = 0
IF @status = 2 OR @status = 5 OR @status = 10
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutySale = 1
	            )  
ELSE IF @status = 1 OR @status = 4
	SET @ValueGoodsDuty = ( SELECT    CAST(ROUND(SUM( Amount * FeeUnit) ,0) AS INT)
	              FROM      tFacd INNER JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
	              WHERE     intSerialNo = @intSerialNo
	                        AND Branch = @intBranch AND dbo.tGood.DutyBuy = 1
	            )  
SET @ValueGoodsDuty = ISNULL(@ValueGoodsDuty , 0)
IF @@ERROR <> 0 
GOTO EventHandler

If @Status = 6 AND @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1 )  
 Set @SumPrice2 = (Select Cast (Sum(Amount * FeeUnit) as Bigint)   From tFacd Where intSerialNo = @intSerialNo2 And Branch = @DestBranch )    
IF @@ERROR <>0  
        GoTo EventHandler  
----------------------------------ServiceRate Calculate  --------------------------------------------------------------  
Declare @ReserveServiceRate Int  
Set @ReserveServiceRate = 0  
If  @TableNo >0  
Begin  
	Declare @Reserve Bit  
	Set @Reserve = (Select Reserve From tTable Where tTable.[No] = @TableNo)  
	If @Reserve = 1  
		Set @ReserveServiceRate = (Select t.ReserveServiceRate From(Select ReserveServiceRate From tPartitions  inner join tTable    
		On tPartitions.PartitionId = tTable.PartitionId and  tTable.[No] = @TableNo)t )  


	If   @Recursive = 0  
	 Update dbo.tTable  
	    Set   dbo.tTable.Empty  = 0  
	        Where dbo.tTable.[No] = @TableNo  AND @Balance = 0
	
	if  @Recursive = 1  
         Update dbo.tTable  
            Set   dbo.tTable.Empty  = 1  
                Where dbo.tTable.[No] = @TableNo  

	If dbo.Get_TableMonitoring() = 1 AND IsNull(@TableNo , 0) <> @OldTableNo   ---Table Monitoring      
	 	Begin      
		DECLARE @intTableUsedNo INT      
		SET @intTableUsedNo = (SELECT TOP 1 intTableUsedNo FROM vwSamar_TableUsage_BusyTable      
		WHERE vwSamar_TableUsage_BusyTable.intTableNo=@OldTableNo and vwSamar_TableUsage_BusyTable.intBranch=@intBranch ORDER BY intTableUsedNo DESC   )   
		SET @intTableUsedNo = ISNULL(@intTableUsedNo , 0) 
		UPDATE tblSamar_TableUsage SET tblSamar_TableUsage.intTableNo = @TableNo      
		WHERE  tblSamar_TableUsage.intTableUsedNo = @intTableUsedNo      
		END        

End  

If @ReserveServiceRate > 0   
 Set @ServiceTotal = @ReserveServiceRate  

-- ===================For Calculate Service In Delivery Or Out ==================
	IF @MasterServePlace = 2 OR @MasterServePlace = 4 
		SET @ServiceTotal = 0

-------------------------------------

 If @ServiceTotal <> 0  
       Set @ServiceTotal = CAST((@SumPrice + @DiscountD) * @ServiceTotal /100 AS INT)  

     IF @@ERROR <>0  
        GoTo EventHandler   
----------------------------------Round Sumprice  --------------------------------------------------------------  
--IF @StationType = 8  --Because ppc program not calculate discountD
    --Set @DiscountTotal = @DiscountTotal + @DiscountD      
 IF @status = 1 OR @status = 2  Or @status = 4 OR @status = 5 OR @status = 10
 BEGIN 
  SELECT @DutyTotal = dbo.Get_Duty(@ValueGoodsDuty ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
  SELECT @TaxTotal = dbo.Get_Tax(@ValueGoodsTax ,@DiscountTotal ,@ServiceTotal ,@CarryFeeTotal, @PackingTotal) 
 END 
    Set @SumPrice = @SumPrice + @ServiceTotal + @CarryFeeTotal + @PackingTotal - @DiscountTotal + @DiscountD + @TaxTotal + @DutyTotal   

    Declare @Remain INT  
    SET @Remain = 0
    IF @Status = 2 OR @status = 10
    BEGIN
    Set @Remain = dbo.RoundSumPrice(@SumPrice )     
    Set @SumPrice = @SumPrice - @Remain  
    Set @DiscountTotal = @DiscountTotal + @Remain  
    END
----------------------------------Calculate Packing---------------------------------------------------------------  
IF dbo.Get_AutoPacking() = 1  
Begin  
    Declare @UserPacking INT  
    SET @UserPacking = ISNULL((Select ISNULL(Sum(Amount),0)  From tGood inner join tFacD On tFacD.GoodCode = tGood.Code   
 where intSerialNo = @intSerialNo and MainType =1 and Branch = @intBranch Group By MainType) ,0)  
    Set @UserPacking = @UserPacking * [dbo].[Get_User_Packing]()  
   Set @SumPrice = @SumPrice + @UserPacking  
   Update tFacm Set PackingTotal = @UserPacking  Where intSerialNo = @intserialNo And Branch = @intBranch   
End  
----------------------------------Net Price Update  --------------------------------------------------------------  

    Update tFacm Set SumPrice = @SumPrice Where intSerialNo = @intserialNo And Branch = @intBranch   
 IF @@ERROR <>0  
         GoTo EventHandler  
If @Status = 6 AND @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1)   

    Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch  
 IF @@ERROR <>0  
         GoTo EventHandler  

Update tFacm Set DiscountTotal = @DiscountTotal Where intSerialNo = @intserialNo  And Branch = @intBranch   
Update tFacm Set ServiceTotal = @ServiceTotal Where intSerialNo = @intserialNo And Branch = @intBranch   
Update tFacm Set RoundDiscount = @Remain  Where intSerialNo = @intserialNo And Branch = @intBranch   
Update tFacm Set TaxTotal = @TaxTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       
Update tFacm Set DutyTotal = @DutyTotal  Where intSerialNo = @intserialNo And Branch = @intBranch       

-----------------------------------------Fill Detail Cash ,....---------------------------------------------------  
DECLARE @Result INT 
If (@Status = 2 OR @Status = 1)  
 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds  , @intBranch  , @Remain  , 2 , @Result OUT 
 IF @@ERROR <>0  
        GoTo EventHandler  
IF @Result = -1
   GoTo EventHandler      
-----------------------------------------Monitoring  --------------------------------------------------------------  

--Declare  @Monitor1 int  
--Declare  @Monitor2 int  

--Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  
--Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  dbo.Get_Current_Branch())  


--If @Monitor1 > 0   
--  exec Notify_to_Clients  
--Else If @Monitor2 > 0   
--  exec Notify_to_Clients  

-- IF @@ERROR <>0  
--        GoTo EventHandler  

-----------------------------------------History  --------------------------------------------------------------  

Exec InsertHistory  @No, @Status , @User , 2 ,@AccountYear  , @intBranch
 IF @@ERROR <>0  
        GoTo EventHandler  

-----------------------------------------Cash  --------------------------------------------------------------  

------------------------------------------Mojodi Control Online-----------------------------------------------------  

Exec InsertMojodiCalculate @Status , @intserialNo , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
IF @STATUs = 6 AND @intSerialNo2 > 0--AND (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
 BEGIN  
 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch  
 IF @@ERROR <>0  
 GoTo EventHandler  
 Exec InsertMojodiCalculate  7,  @intserialNo2 , @AccountYear , @destbranch  
 IF @@ERROR <>0  
 GoTo EventHandler  
 END   
------------------------------------------Update Balance After Recived----------------------------
Declare @SumRecieved INT
SET @SumRecieved =0

Set @SumRecieved =(Select IsNull(SUM(Bestankar),0)  From   tblAcc_Recieved 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCash 
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCard
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intChequeAmount),0)  From   dbo.tFacCheque
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )
Set @SumRecieved = @SumRecieved + (Select IsNull(SUM(intAmount),0)  From   dbo.tFacCredit
 Where intSerialNo = @intSerialNo  and Branch = @intBranch )

    IF @SumRecieved >= @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 1 , FacPayment = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch
    ELSE IF @SumRecieved < @sumPrice AND @Status = 2 
        UPDATE  tfacm
        SET     [Balance] = 0
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch


COMMIT TRANSACTION  

---------------------------------------Sync Kitchen Monitoring -----------------------------------------------------------------------

IF EXISTS (select * from syscomments where id = object_id ('dbo.sp_KM_SyncKitchenFacGoods'))
	exec dbo.sp_KM_SyncKitchenFacGoods @intSerialNo , 2

--------------------------------------------------------------------------------------------------------------------------------------
Set @LastFacMNo = @No  
Return @LastFacMNo  


EventHandler:  
    ROLLBACK TRAN  
    SET @LastFacMNo = -1   

    RETURN @LastFacMNo
GO
