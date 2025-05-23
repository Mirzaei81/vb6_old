

--Script_V26_16_Fix5_Added2
--DestBranch is CurrentBranch For Replication
--اصلاح گردش کالا در انبار برای شعبات
--اصلاح بروزرسانی در انبار شعبات
--93/04/14

ALTER  PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

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
            @lastFacMNo INT OUT      
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

DECLARE @SumPrice  float      
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

---------------------------------------------------------      

 Declare @intBranch  int      
 Declare @ShiftNo int      
 DECLARE @TempNo INT 

 select @intBranch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@DetailsString))      

    DECLARE @IdentityNo INT
    SELECT  @IdentityNo = ISNULL(MAX(intserialno), 0) + 1
    FROM    tFacm
    WHERE   Branch = @intBranch 

    IF @IdentityNo < ( @intBranch * 10000000 ) 
        SET @IdentityNo = ( @intBranch * 10000000 )

 SET @NO1 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND AccountYear = @AccountYear)      

 SET @ShiftNo= dbo.Get_Shift(GETDATE())      
 SET @TempNo = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=@Status  And Branch =  @intBranch AND Date = @Date AND ShiftNo = @ShiftNo)      


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
		TempNo    
		
 )      
     Values       

(			    @IdentityNo ,  
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
		@TempNo  
 )      
    SET @intserialNo = @IdentityNo
     IF @@ERROR <>0      
        GoTo EventHandler       



declare @destbranch  INT 
SET @destbranch = 0
DECLARE @TempNo2 INT 
 
If @Status = 6  -- And (@destbranch= @intBranch Or dbo.AutoResid() = 1)    
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
		SET @intserialNo2 = @IdentityNo + 1      
		 IF @@ERROR <>0      
			GoTo EventHandler      

            UPDATE  tfacm
            SET     NvcDescription = @NvcDescription + N' رسيد -   '
                    + CAST(@No2 AS NVARCHAR(8))
            WHERE   intSerialNo = @intserialNo
            UPDATE  tfacm
            SET     RefrenceHavale = @intserialNo2
            WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch

end      


----------------------------------Fill Details Factor  --------------------------------------------------------------      
If @Status = 6 -- AND (@destbranch= @intBranch  Or dbo.AutoResid() = 1)        
 exec InsertFactorDetail @DetailsString , @intserialNo , @intserialNo2, @Customer , @intBranch      
Else       
 exec InsertFactorDetail @DetailsString , @intserialNo , 0, @Customer , @intBranch      

     IF @@ERROR <>0      
        GoTo EventHandler      
--------------------------------- Fill Havalem & HavaleD   ------------------------------------------------------------      

----------------------------------Total SumPrice Calculate  --------------------------------------------------------------      
DECLARE @DiscountD INT 
Set @DiscountD = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * discount/100 ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        
Set @SumPrice = (Select CAST(ROUND(SUM( (Amount * FeeUnit) * (1 - discount/100) ) ,0) AS INT)  From tFacd Where intSerialNo = @intserialNo And Branch = @intBranch )        

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


If @ReserveServiceRate > 0       
 Set @ServiceTotal = @ReserveServiceRate      


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

If @Status = 6 -- AND (@destbranch= @intBranch )  -- Or dbo.AutoResid() = 1   
	Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @DestBranch       
      IF @@ERROR <>0       

        GoTo EventHandler           
-------------------------------------Fill Detail Cash , ..........--------------------------------------------------      
IF (@Status =  1 OR @Status = 2 )      
 exec Do_SaveInDetailsFactorReceived @intSerialNo, @ds ,  @intBranch  , @Remain    

     IF @@ERROR <>0      
   GoTo EventHandler      
-------------------------------------Monitoring---------------------------------------------------------------------      
Declare  @Monitor1 int      
Declare  @Monitor2 int       

Set @Monitor1 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  16  = 16) and  IsActive =1 And Branch =  @intBranch)      
Set @Monitor2 = (Select Count(Stationid) from  dbo.tStations Where (StationType  &  32  = 32) and  IsActive =1 And Branch =  @intBranch)      


IF @Monitor1 > 0       
   exec Notify_to_Clients      

Else If @Monitor2 > 0       
   exec Notify_to_Clients      

----------------------------History---------------------------      

Exec InsertHistory  @No1, @Status , @User , 1 , @AccountYear , @intBranch      
IF @STATUs = 6 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 7 , @User , 1 , @AccountYear , @destbranch      


----------------------------Cash ---------------------------      

------------------------Mojodi Control Online--------------------------------------------      

Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @intBranch      
IF @@ERROR <>0      
 GoTo EventHandler      
IF @STATUS = 6 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1      
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
Return @lastFacMNo      

EventHandler:      

    ROLLBACK TRAN      
    SET @LastFacMNo = -1      

    RETURN @lastFacMNo



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[InsertFactorDetail]  (
	 @DetailsString Nvarchar(4000) ,
	 @intSerialNo bigint ,
	 @intserialNo2 bigint ,
	 @Customer Bigint ,
	 @Branch int = Null
	
) 
As


if @Branch is null
    select @Branch = branch from tInventory where inventoryNo=(SELECT Top 1  intInventoryNo FROM Split(@DetailsString))



Declare @Status Int 

Set @Status = (Select Status from tfacm Where intserialno = @intSerialNo and Branch = @Branch)


     INSERT INTO tFacD
	(
	    
		intRow,
		Amount ,
		GoodCode  ,
		FeeUnit ,
		Discount ,
		Rate ,
		ChairName ,
		[ExpireDate] ,
		intInventoryNo ,
		DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
		ServePlace ,
		DifferencesCodes , 
		DifferencesDescription ,
		intSerialNo , 
		Branch 
	)
	     SELECT
		
		tmpTable.Row ,
		tmpTable.Amount ,
		tmpTable.GoodCode ,
		tmpTable.FeeUnit ,
		tmpTable.Discount ,
		tmpTable.Rate ,
		tmpTable.ChairName ,
		tmpTable.[ExpireDate],
		tmpTable.intInventoryNo ,
		tmpTable.DestInventoryNo ,--Because Has a Relation and Can not insert for  Another Branch
		tmpTable.ServePlace ,
		tmpTable.DifferencesCode ,
		tmpTable.DifferencesDescription ,
		@intSerialNo , 
		@Branch 	
	
	FROM (SELECT * FROM Split(@DetailsString )) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode

	If @Status = 6
	Begin
	
	declare @destbranch INT
	select @destbranch=@Branch --branch from tInventory where inventoryNo=(SELECT Top 1  DestInventoryNo FROM Split(@DetailsString))
	  	   begin
			 INSERT INTO tFacD
			(
			    
				intRow,
				Amount ,
				GoodCode  ,
				FeeUnit ,
				Discount ,
				Rate ,
				ChairName ,
				[ExpireDate] ,
				intInventoryNo ,
				DestInventoryNo , --Because Has a Relation and Can not insert for  Another Branch
				ServePlace ,
				DifferencesCodes , 
				DifferencesDescription ,
				intSerialNo , 
				Branch
			)
				 SELECT
				
				tmpTable.Row ,
				tmpTable.Amount ,
				tmpTable.GoodCode ,
				tmpTable.FeeUnit ,
				tmpTable.Discount ,
				tmpTable.Rate ,
				tmpTable.ChairName ,
				tmpTable.[ExpireDate],
				tmpTable.DestInventoryNo ,
				tmpTable.intInventoryNo ,--Because Has a Relation and Can not insert for  Another Branch
				tmpTable.ServePlace ,
				tmpTable.DifferencesCode ,
				tmpTable.DifferencesDescription ,
				@intSerialNo2 , 
				@DestBranch --dbo.Get_Current_Branch()
		
		
			FROM (SELECT * FROM Split(@DetailsString )) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode
	
		   end
	end
	
			

Update tFacD Set Amount = 1 where amount = 0 and intserialno = @intSerialNo and Branch = @Branch
--Update tFacD Set DestInventoryNo = Null Where intserialno = @intSerialNo and Branch = dbo.Get_Current_Branch()

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

--exec CheckPreSave_Edit N'2;11010002;33000;10;1;; ;1;;1/', 2, 1174
--GO


--ControlMojodi Online
--V26_11 & V26_12 



ALTER   PROCEDURE [dbo].[InsertMojodiCalculate]
    (
      @Status INT,
      @intserialNo BIGINT,
      @AccountYear SMALLINT,
      @Branch INT = NULL
    )
AS 
    IF @Branch IS NULL 
        SET @Branch = dbo.Get_Current_Branch()

---------------------------------------Mojodi Control Online---------------------------------------------------------

    IF @Status = 2 
        BEGIN
	--IF dbo.AutoHavale() = 0
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    SaleAmount = SaleAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - X.Amount ,
                    SaleAmount = SaleAmount + X.Amount
            FROM    ( SELECT    SUM(( Amount * fltUsedValue ) + ( [Amount] * [Pert] )) AS Amount,
                                GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType = 4 
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) X
            WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = X.intInventoryNo
                    --AND tInventory_Good.Branch = X.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

 	    UPDATE  tInventory_Good	--Mojodi not less zero because in edit mode not show message
            SET     Mojodi = 0
            FROM    ( SELECT    GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType = 4
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) X
            WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = X.intInventoryNo
                    --AND tInventory_Good.Branch = X.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
		    AND tInventory_Good.Mojodi < 0

	IF dbo.AutoHavale() = 1
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - Y.Amount ,
                    SaleAmount = SaleAmount + Y.Amount
            FROM    ( SELECT    SUM(( Amount * fltUsedValue ) + ( [Amount] * [Pert] )) AS Amount,
                                GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType <> 4 
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) Y
            WHERE   Y.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = Y.intInventoryNo
                    --AND tInventory_Good.Branch = Y.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

        END
    IF @Status = 1 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + t.Amount,
                    BuyAmount = BuyAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 3 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    LossAmount = LossAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

	IF dbo.AutoHavale() = 1
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - Y.Amount ,
                    LossAmount = LossAmount + Y.Amount
            FROM    ( SELECT    SUM(( Amount * fltUsedValue ) + ( [Amount] * [Pert] )) AS Amount,
                                GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType <> 4 
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) Y
            WHERE   Y.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = Y.intInventoryNo
                    --AND tInventory_Good.Branch = Y.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

        END
    IF @Status = 4 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    BuyReturnAmount = BuyReturnAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 5 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + t.Amount,
                    SaleReturnAmount = SaleReturnAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 6 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - t.Amount,
                    FromStoreAmount = FromStoreAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

	IF dbo.AutoHavale() = 1
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi - Y.Amount ,
                    FromStoreAmount = FromStoreAmount + Y.Amount
            FROM    ( SELECT    SUM(( Amount * fltUsedValue ) + ( [Amount] * [Pert] )) AS Amount,
                                GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType <> 4 
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) Y
            WHERE   Y.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = Y.intInventoryNo
                    --AND tInventory_Good.Branch = Y.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
    IF @Status = 7 
        BEGIN
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + t.Amount,
                    toStoreAmount = toStoreAmount + t.Amount
            FROM    ( SELECT    SUM(Amount) AS Amount,
                                GoodCode,
                                intInventoryNo,
                                Branch
                      FROM      tfacd
                      WHERE     intserialno = @intserialNo
                                AND Branch = @Branch
                      GROUP BY  GoodCode,
                                intInventoryNo,
                                Branch
                    ) t
            WHERE   t.GoodCode = tInventory_Good.GoodCode
                    AND tInventory_Good.InventoryNo = t.intInventoryNo
                    --AND tInventory_Good.Branch = t.Branch
                    AND tInventory_Good.AccountYear = @AccountYear

	IF dbo.AutoHavale() = 1
            UPDATE  tInventory_Good
            SET     Mojodi = Mojodi + Y.Amount ,
                    toStoreAmount = toStoreAmount + Y.Amount
            FROM    ( SELECT    SUM(( Amount * fltUsedValue ) + ( [Amount] * [Pert] )) AS Amount,
                                GoodFirstCode,
                                intInventoryNo,
                                Branch
                      FROM      ( SELECT    *
                                  FROM      tFacd
                                            INNER JOIN usepercent ON tFacd.GoodCode = usepercent.code
                                                                     AND tFacd.serveplace = usepercent.intserveplace
                                  WHERE     intserialNo = @intserialNo
                                            AND Branch = @Branch
                                            
                                ) FirstGoods
                                INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType <> 4 
                      GROUP BY  FirstGoods.GoodFirstCode,
                                FirstGoods.intInventoryNo,
                                FirstGoods.Branch
                    ) Y
            WHERE   Y.GoodFirstCode = tInventory_Good.Goodcode
                    AND tInventory_Good.InventoryNo = Y.intInventoryNo
                    --AND tInventory_Good.Branch = Y.Branch
                    AND tInventory_Good.AccountYear = @AccountYear
        END
--===============================================

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER  PROCEDURE [dbo].[GetInventoryGood_Mojodi_New]
    (
      @intLanguage INT,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @DateBefore NVARCHAR(8),
      @DateAfter NVARCHAR(8),
      @GoodCode INT,
      @InventoryNo INT,
      @Branch INT,
      @AccountYear SMALLINT 
    )
AS 
    SELECT  @DateBefore AS DateBefore,
            @DateAfter AS DateAfter,
            @SystemDay + ' ' + @SystemDate + ' ' + N' ساعت : ' + @SystemTime AS Sysdate,
            [dbo].[tFacD].[intInventoryNo],
            ( SELECT    [Description]
              FROM      [dbo].[tInventory]
              WHERE     [InventoryNo] = [dbo].[tFacD].[intInventoryNo]
            ) AS [FromStore],
            ISNULL(( SELECT [Description]
                     FROM   [dbo].[tInventory]
                     WHERE  [InventoryNo] = [dbo].[tFacD].[DestInventoryNo]
                   ), '') AS [DestDescription],
            [dbo].[tFacD].[GoodCode],
            SUM([dbo].[tFacD].[Amount] * [dbo].[tStatusType].[Flag]) AS [Amount],
            [dbo].[tFacD].[FeeUnit] * ( 1 - ( [dbo].[tFacD].[Discount] / 100 ) ) AS [FeeUnit],
            [dbo].[tFacM].[Status],
            [dbo].[tFacM].[Branch],
            [dbo].[tFacM].[Date],
            [dbo].[tFacM].[No],
            [dbo].[tGood].[Name],
            [dbo].[tGood].[Name] AS [GoodName],
            [dbo].[tGood].[BuyPrice],
            [dbo].[tInventory_Good].[Mojodi],
            [dbo].[tInventory_Good].[FirstMojodi],
            [dbo].[tInventory_Good].[FirstPrice],
            [dbo].[tStatusType].[NvcDescription]
    FROM    [dbo].[tInventory_Good]
            INNER JOIN dbo.tFacD ON tInventory_Good.GoodCode = tFacd.GoodCode
                                    AND tInventory_Good.InventoryNo = tFacD.intInventoryNo
                                    AND tInventory_Good.AccountYear = @AccountYear
                                    --AND ( tInventory_Good.Branch = @Branch
                                          --OR @Branch = 0
                                       -- )
            INNER JOIN tInventory ON tInventory.InventoryNo = tInventory_Good.InventoryNo
            INNER JOIN dbo.tFacM ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    AND dbo.tFacM.Branch = dbo.tFacD.Branch
            INNER JOIN dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo
            INNER JOIN dbo.tGood ON dbo.tGood.Code = dbo.tFacd.GoodCode
    WHERE   [dbo].[tFacM].[Recursive] = 0
            AND [dbo].[tFacM].[AccountYear] = @AccountYear
            AND [dbo].[tFacM].[Date] >= @DateBefore
            AND [dbo].[tFacM].[Date] <= @DateAfter
            AND [dbo].[tFacD].[GoodCode] = @GoodCode
            --AND ( [dbo].[tFacM].[Branch] = @Branch
            --      OR @Branch = 0
            --    )
            AND ( [dbo].[tFacD].[intInventoryNo] = @InventoryNo
                  OR @InventoryNo = 0
                )
            AND ( [dbo].[tFacM].[Status] <> 2
                  AND [dbo].[tFacM].[Status] <> 5
                  AND [dbo].[tFacM].[Status] <> 10
                )
    GROUP BY [dbo].[tFacM].[Date],
            [dbo].[tInventory_Good].[Mojodi],
            [dbo].[tGood].[Name],
            [dbo].[tFacM].[Branch],
            [dbo].[tFacD].[intInventoryNo],
            [dbo].[tFacD].[GoodCode],
            [dbo].[tFacM].[Status],
            [dbo].[tStatusType].[NvcDescription],
            [dbo].[tInventory].[Description],
            [dbo].[tFacD].[FeeUnit],
            [dbo].[tInventory_Good].[FirstMojodi],
            [dbo].[tFacD].[DestInventoryNo],
            [dbo].[tInventory_Good].[FirstPrice],
            [dbo].[tFacD].[Discount],
            [dbo].[tFacM].[No],
            [dbo].[tFacM].[intSerialNo],
            [dbo].[tGood].[BuyPrice]
    ORDER BY [dbo].[tFacM].[Date] ASC,
            [dbo].[tFacM].[intSerialNo] ASC
--===============================================
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  FUNCTION [dbo].[tblTotal_tInventory_tGood_For_Mojodi]
    (
      @intLanguage INT ,
      @SystemDate NVARCHAR(50) ,
      @SystemDay NVARCHAR(50) ,
      @SystemTime NVARCHAR(50) ,
      @DateBefore NVARCHAR(50) ,
      @DateAfter NVARCHAR(50) ,
      @Type INT ,
      @InventoryNo1 INT ,
      @InventoryNo2 INT ,
      @Branch INT ,
      @UsePercentFlag INT ,
      @AccountYear SMALLINT
    )
RETURNS @ReturnTable TABLE
    (
      DateBefore NVARCHAR(50) ,
      DateAfter NVARCHAR(50) ,
      Sysdate NVARCHAR(50) ,
      [Name] NVARCHAR(50) ,
      goodtype INT ,
      GoodCode INT ,
      Branch INT ,
      intInventoryNo INT ,
      firstMojodi FLOAT ,
      BuyAmount FLOAT ,
      SaleAmount FLOAT ,
      LossAmount FLOAT ,
      BuyReturnAmount FLOAT ,
      SaleReturnAmount FLOAT ,
      FromStoreAmount FLOAT ,
      toStoreAmount FLOAT ,
      Mojodi FLOAT
    )
AS 
    BEGIN
-------------------------------------------------------------------------------------------------------------------
        DECLARE @TimeTitle NVARCHAR(10)
        IF @intLanguage = 0 
            SET @TimeTitle = N' ساعت : '
        ELSE 
            SET @TimeTitle = N'Time: '
        INSERT  INTO @ReturnTable
                ( DateBefore ,
                  DateAfter ,
                  Sysdate ,
                  [Name] ,
                  goodtype ,
                  GoodCode ,
                  Branch ,
                  intInventoryNo ,
                  firstMojodi ,
                  BuyAmount ,
                  SaleAmount ,
                  LossAmount ,
                  BuyReturnAmount ,
                  SaleReturnAmount ,
                  FromStoreAmount ,
                  toStoreAmount ,
                  Mojodi  
                )
                SELECT  DateBefore ,
                        DateAfter ,
                        Sysdate ,
                        GoodFirstName ,
                        @Type AS goodtype ,
                        GoodCode ,
                        Branch ,
                        InventoryNo ,
                        firstMojodi ,
                        BuyAmount ,
                        SaleAmount ,
                        LossAmount ,
                        BuyReturnAmount ,
                        SaleReturnAmount ,
                        FromStoreAmount ,
                        toStoreAmount ,
                        Mojodi
                FROM    ( SELECT    Y.* ,
                                    CASE @intLanguage
                                      WHEN 0 THEN tinventory.Description
                                      WHEN 1 THEN tinventory.LatinDescription
                                    END AS InventoryName
                          FROM      ( SELECT    ISNULL(tGood.Name,
                                                       W.GoodFirstName) AS GoodFirstName ,
                                                tUnitGood.[Description] ,
                                                ISNULL(tInventory_Good.GoodCode,
                                                       W.GoodFirstCode) AS GoodCode ,
                                                ISNULL(tInventory_Good.Branch,
                                                       W.Branch) AS Branch ,
                                                ISNULL(tInventory_Good.InventoryNo,
                                                       W.intInventoryNo) AS InventoryNo ,
                                                ISNULL(FirstMojodi, 0) AS firstMojodi ,
                                                ISNULL(MAX(W.BuyAmount), 0) AS BuyAmount ,
                                                ISNULL(SUM(W.SaleAmount), 0) AS SaleAmount ,
                                                ISNULL(MAX(W.LossAmount), 0) AS LossAmount ,
                                                ISNULL(MAX(W.BuyReturnAmount),
                                                       0) AS BuyReturnAmount ,
                                                ISNULL(MAX(W.SaleReturnAmount),
                                                       0) AS SaleReturnAmount ,
                                                ISNULL(MAX(W.FromStoreAmount),
                                                       0) AS FromStoreAmount ,
                                                ISNULL(MAX(W.toStoreAmount), 0) AS toStoreAmount ,
                                                ISNULL(FirstMojodi, 0)
                                                + ISNULL(MAX(W.BuyAmount), 0)
                                                - ISNULL(SUM(W.SaleAmount), 0)
                                                - ISNULL(MAX(W.LossAmount), 0)
                                                - ISNULL(MAX(W.BuyReturnAmount),
                                                         0)
                                                + ISNULL(MAX(W.SaleReturnAmount),
                                                         0)
                                                - ISNULL(MAX(W.FromStoreAmount),
                                                         0)
                                                + ISNULL(MAX(W.toStoreAmount),
                                                         0) AS Mojodi ,
                                                @DateBefore AS DateBefore ,
                                                @DateAfter AS DateAfter ,
                                                @SystemDay + ' ' + @SystemDate
                                                + ' ' + @TimeTitle
                                                + @SystemTime AS Sysdate
                                      FROM      tInventory_Good
                                                INNER JOIN dbo.tGood ON dbo.tGood.Code = tInventory_Good.GoodCode
                                                              AND tInventory_Good.InventoryNo >= @InventoryNo1
                                                              AND tInventory_Good.InventoryNo <= @InventoryNo2
                                                              AND tInventory_Good.Branch = @Branch
                                                              AND dbo.tInventory_Good.AccountYear = @AccountYear
                                                              AND tGood.GoodType = @Type
                                                FULL OUTER   JOIN ( SELECT
                                                              MAX(ISNULL(X.Branch,
                                                              FirstGoods.branch)) AS branch ,
                                                              MAX(ISNULL(X.intInventoryNo,
                                                              FirstGoods.intInventoryNo)) AS intInventoryNo ,
                                                              ISNULL(FirstGoods.GoodFirstName,
                                                              X.GoodFirstName) AS GoodFirstName ,
                                                              ISNULL(X.GoodfirstCode,
                                                              FirstGoods.GoodFirst1) AS GoodfirstCode ,
                                                              MAX(ISNULL(FirstGoods.SaleAmount,
                                                              0))
                                                              + MAX(ISNULL(X.SaleAmount,
                                                              0)) AS saleamount ,
                                                              MAX(ISNULL(FirstGoods.SaleReturnAmount,
                                                              0))
                                                              + MAX(ISNULL(X.SaleReturnAmount,
                                                              0)) AS SaleReturnAmount ,
                                                              MAX(ISNULL(FirstGoods.Buy,
                                                              0)) AS BuyAmount ,
                                                              MAX(ISNULL(FirstGoods.Losses,
                                                              0)) AS LossAmount ,
                                                              MAX(ISNULL(FirstGoods.BuyReturn,
                                                              0)) AS BuyReturnAmount ,
                                                              MAX(ISNULL(FirstGoods.FromStore,
                                                              0)) AS FromStoreAmount ,
                                                              MAX(ISNULL(FirstGoods.ToStore,
                                                              0)) AS ToStoreAmount
                                                              FROM
                                                              ( SELECT
                                                              Branch ,
                                                              intInventoryNo ,
                                                              GoodfirstCode ,
                                                              SUM(SaleAmount) AS SaleAmount ,
                                                              SUM(SaleReturnAmount) AS SaleReturnAmount ,
                                                              GoodFirstName
                                                              FROM
                                                              ( SELECT
                                                              Branch ,
                                                              intInventoryNo ,
                                                              ISNULL(GoodfirstCode,
                                                              GoodCode) AS GoodfirstCode ,
                                                              ( ( SaleAmount
                                                              * fltUsedvalue )
                                                              + ( [SaleAmount]
                                                              * Pert ) ) AS SaleAmount ,
                                                              SaleReturnAmount ,
                                                              CASE @intLanguage
                                                              WHEN 0
                                                              THEN tGood.[Name]
                                                              WHEN 1
                                                              THEN tGood.LatinName
                                                              END AS GoodFirstName
                                                              FROM
                                                              ( SELECT
                                                              @Branch AS Branch , --tFacd.Branch ,
                                                              intInventoryNo ,
                                                              Goodcode ,
                                                              tGood.[Name] AS MainName ,
                                                              tfacd.Serveplace ,
                                                              CASE Status
                                                              WHEN 2
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleAmount ,
                                                              CASE Status
                                                              WHEN 5
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleReturnAmount
                                                              FROM
                                                              dbo.tFacM
                                                              INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                                              AND dbo.tFacM.Branch = dbo.tFacD.Branch
                                                              INNER JOIN tGood ON tGood.Code = tFacd.GoodCode
                                                              WHERE
                                                              dbo.tFacM.Recursive = 0
                                                              AND dbo.tFacM.AccountYear = @AccountYear
                                                              AND tFacM.[Date] >= @DateBefore
                                                              AND tFacM.[Date] <= @DateAfter
                                                              AND ( dbo.tFacD.intInventoryNo >= @InventoryNo1
                                                              AND dbo.tFacD.intInventoryNo <= @InventoryNo2
                                                              --AND tFacd.Branch = @Branch
                                                              )
					--AND dbo.tFacM.ShiftNo = dbo.Get_Current_Shift(@SystemTime) --Just only for Mashad(Malek Restaurant)
                                                              GROUP BY Goodcode ,
                                                              Name ,
                                                              tfacd.serveplace ,
                                                              tfacm.Status ,
                                                              intInventoryNo --,
                                                              --tFacd.Branch
                                                              ) T
                                                              LEFT OUTER JOIN usepercent ON T.GoodCode = usepercent.code
                                                              AND T.serveplace = usepercent.intserveplace
                                                              INNER JOIN tGood ON tGood.Code = usepercent.GoodFirstCode
                                                              ) AS U
                                                              GROUP BY GoodFirstCode ,
                                                              intInventoryNo ,
                                                              Branch ,
                                                              GoodFirstName --, ServePlace
                                                              ) X
                                                              FULL OUTER  JOIN ( SELECT
                                                              F.Branch ,
                                                              F.intInventoryNo ,
                                                              F.GoodFirst1 ,
                                                              F.GoodFirstName ,
                                                              SUM(Buy) AS Buy ,
                                                              SUM(SaleAmount) AS SaleAmount ,
                                                              SUM(Losses) AS Losses ,
                                                              SUM(BuyReturn) AS BuyReturn ,
                                                              SUM(SaleReturnAmount) AS SaleReturnAmount ,
                                                              SUM(FromStore) AS FromStore ,
                                                              SUM(ToStore) AS ToStore
                                                              FROM
                                                              ( SELECT
                                                              @Branch AS Branch ,--tFacd.Branch ,
                                                              intInventoryNo ,
                                                              GoodCode AS GoodFirst1 ,
                                                              tgood.NAME AS GoodFirstName ,
                                                              CASE Status
                                                              WHEN 1
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS Buy ,
                                                              CASE Status
                                                              WHEN 2
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleAmount ,
                                                              CASE Status
                                                              WHEN 3
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS Losses ,
                                                              CASE Status
                                                              WHEN 4
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS BuyReturn ,
                                                              CASE Status
                                                              WHEN 5
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS SaleReturnAmount ,
                                                              CASE Status
                                                              WHEN 6
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS FromStore ,
                                                              CASE Status
                                                              WHEN 7
                                                              THEN SUM(Amount)
                                                              ELSE 0
                                                              END AS ToStore
                                                              FROM
                                                              dbo.tFacM
                                                              INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                                              AND dbo.tFacM.Branch = dbo.tFacD.Branch
                                                              INNER JOIN tGood ON tGood.Code = tFacd.GoodCode
                                                              WHERE
                                                              dbo.tFacM.Recursive = 0
                                                              AND dbo.tFacM.AccountYear = @AccountYear
                                                              AND tFacM.[Date] >= @DateBefore
                                                              AND tFacM.[Date] <= @DateAfter
                                                              AND ( dbo.tFacD.intInventoryNo >= @InventoryNo1
                                                              AND dbo.tFacD.intInventoryNo <= @InventoryNo2
                                                              --AND tFacd.Branch = @Branch
                                                              )
			--AND dbo.tFacM.ShiftNo = dbo.Get_Current_Shift(@SystemTime) --Just only for Mashad(Malek Restaurant)
                                                              GROUP BY Goodcode ,
                                                              tfacm.Status ,
                                                              intInventoryNo ,
                                                            --  tFacd.Branch ,
                                                              tGood.[name]
                                                              ) F
                                                              GROUP BY F.GoodFirst1 ,
                                                              F.intInventoryNo ,
                                                              F.Branch ,
                                                              GoodFirstName
                                                              ) FirstGoods ON FirstGoods.GoodFirst1 = X.GoodfirstCode
                                                              GROUP BY GoodFirstCode ,
                                                              FirstGoods.intInventoryNo ,
                                                              FirstGoods.Branch ,
                                                              FirstGoods.GoodFirstName ,
                                                              X.GoodFirstName ,
                                                              GoodFirst1
                                                              ) W ON tInventory_Good.GoodCode = W.GoodfirstCode
                                                              AND tInventory_Good.Branch = W.Branch
                                                              AND tInventory_Good.InventoryNo = W.intInventoryNo
                                                INNER JOIN tUnitGood ON tGood.Unit = tUnitGood.Code
                                      WHERE     GoodType = @Type
                                      GROUP BY  tInventory_Good.GoodCode ,
                                                GoodFirstCode ,
                                                tGood.Name ,
                                                W.GoodFirstName ,
                                                FirstMojodi ,
                                                W.intInventoryNo ,
                                                tUnitGood.[Description] ,
                                                W.Branch ,
                                                tInventory_Good.InventoryNo ,
                                                tInventory_Good.Branch
                                    ) y
                                    INNER JOIN tInventory ON tInventory.InventoryNo = Y.InventoryNo
                                                             AND tInventory.Branch = Y.Branch
                        ) AS T

-----------------------------------------------------------------------------------------
--END
        RETURN


    END
--===============================================
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTotal_Inventory_Good_CycleStock_tInventory]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTotal_Inventory_Good_CycleStock] DROP CONSTRAINT FK_tblTotal_Inventory_Good_CycleStock_tInventory
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tFacD_tInventory]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tFacD] DROP CONSTRAINT FK_tFacD_tInventory
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tInventory_Good_tInventory]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tInventory_Good] DROP CONSTRAINT FK_tInventory_Good_tInventory
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tInventory_Level1_tInventory]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tInventory_Level1] DROP CONSTRAINT FK_tInventory_Level1_tInventory
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tInventory]') and OBJECTPROPERTY(id, N'IsPrimaryKey') = 1)
ALTER TABLE [dbo].[tInventory] DROP CONSTRAINT [PK_tInventory]
GO


ALTER TABLE [dbo].[tInventory] WITH NOCHECK ADD 
	CONSTRAINT [PK_tInventory] PRIMARY KEY  CLUSTERED 
	(
		[InventoryNo]
	)  ON [PRIMARY] 
GO


ALTER TABLE [dbo].[tblTotal_Inventory_Good_CycleStock] ADD 
	CONSTRAINT FK_tblTotal_Inventory_Good_CycleStock_tInventory FOREIGN KEY 
	(
		InventoryNo
	) REFERENCES [dbo].[tInventory] (
		InventoryNo
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tFacD] ADD 
	CONSTRAINT FK_tFacD_tInventory FOREIGN KEY 
	(
		intInventoryNo
	) REFERENCES [dbo].[tInventory] (
		InventoryNo
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tInventory_Good] ADD 
	CONSTRAINT FK_tInventory_Good_tInventory FOREIGN KEY 
	(
		InventoryNo
	) REFERENCES [dbo].[tInventory] (
		InventoryNo
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tInventory_Level1] ADD 
	CONSTRAINT FK_tInventory_Level1_tInventory FOREIGN KEY 
	(
		InventoryNo
	) REFERENCES [dbo].[tInventory] (
		InventoryNo
	) ON UPDATE CASCADE 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tStation_Inventory_Good_tInventory_Good]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tStation_Inventory_Good] DROP CONSTRAINT FK_tStation_Inventory_Good_tInventory_Good
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PK_tInventory_Good]') and OBJECTPROPERTY(id, N'IsPrimaryKey') = 1)
ALTER TABLE [dbo].[tInventory_Good] DROP CONSTRAINT [PK_tInventory_Good]
GO

ALTER TABLE [dbo].[tInventory_Good] WITH NOCHECK ADD 
	CONSTRAINT [PK_tInventory_Good] PRIMARY KEY  CLUSTERED 
	(
		[InventoryNo],
		[GoodCode],
		[AccountYear]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tStation_Inventory_Good] ADD 
	CONSTRAINT [FK_tStation_Inventory_Good_tInventory_Good] FOREIGN KEY 
	(
		[InventoryNo],
		[GoodCode],
		[AccountYear]
	) REFERENCES [dbo].[tInventory_Good] (
		[InventoryNo],
		[GoodCode],
		[AccountYear]
	) ON DELETE CASCADE  ON UPDATE CASCADE
	
GO
