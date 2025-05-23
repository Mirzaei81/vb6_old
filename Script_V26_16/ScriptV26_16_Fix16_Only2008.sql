
--Only for Sql 2008
--Script_V26_16_Fix16_only2008.sql
--اضافه شدن  تعداد ردیف های فاکتور برای ثبت
-- با امکان محاسبه ضریب مصرفAuto Havale
-- حواله دستی روزانه با امکان محاسبه ضریب مصرف 
--  در فرم انبارگردانی - اضافه شدن فی آخرین خرید به فی خرید در هنگام انتقال به سال مالی جدید
-- در فرم کاردکس کالا - اضافه شدن فی آخرین خرید به فی اول دوره در سال مالی 
--اضافه شدن گروه اصلی و فرعی به فرم ضریب مصرف
--اصلاح گزارش موجودی جنسی از تاریخ تا تاریخ   RepInventoryGood_Mojodi_All_A4.rpt
--95/03/12

IF NOT EXISTS(SELECT * FROM tblPub_Script2 WHERE [Version] = 26 AND Script = 16 AND FixNumber = 16 )

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
			  16
			)
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   FUNCTION [dbo].[AutoHavale]()
RETURNS int 
AS  
BEGIN 
	Return 1
END


GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS OFF
GO



ALTER   Function Split

(
    @nvcMainString NVARCHAR(Max)
)

RETURNS  @ReturnTable TABLE(
	Row int IDENTITY (1, 1) NOT NULL ,
	Amount FLOAT , 
	GoodCode INT , 
	FeeUnit Float , 
	Discount Float ,
	Rate Int ,
	ChairName nvarchar(50),
	[ExpireDate] Int,
	intInventoryNo Int ,
	DestInventoryNo INT ,
	ServePlace INT , 
	DifferencesCode NVARCHAR(50) , 
	DifferencesDescription NVARCHAR(500))
	
As

BEGIN

IF @nvcMainString IS NOT  NULL
BEGIN
    DECLARE @intDelimiterPosField  INT
    DECLARE @intDelimiterPosRecord INT

    DECLARE @Amount FLOAT
    DECLARE @GoodCode INT
    DECLARE @FeeUnit Float
    Declare @Discount Float	
    Declare @Rate Int	
    DECLARE @ChairName  NVARCHAR(50)
    DECLARE @ExpireDate  INT 
    DECLARE @intInventoryNo INT
    DECLARE @DestInventoryNo INT
    DECLARE @ServePlace INT

    DECLARE @DifferencesCode NVARCHAR(50)
    DECLARE @DifferencesDescription NVARCHAR(500)

    DECLARE @TempDifference int 
    DECLARE @TempTable Table (nvcMainString NVARCHAR(Max))
    

    insert into @TempTable values (@nvcMainString)
   

    SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
    SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)

    WHILE @intDelimiterPosRecord <> 0
    BEGIN
--**********
        	SET @Amount = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS FLOAT)  from @TempTable )

        	SET @Amount =  ROUND(CAST(@Amount AS DECIMAL(15,3)),3)

	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @GoodCode = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )

        	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @FeeUnit = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Float)  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @Discount = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Float)  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @Rate = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS Int)  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @ChairName = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS NVARCHAR(50))  from @TempTable )

        	Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @ExpireDate = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )

        	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @intInventoryNo = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )

        	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )


-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)


        	SET @DestInventoryNo = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
	If @DestInventoryNo = 0 SET @DestInventoryNo = Null
             
	Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )


-----------------------------------------------------------------------------------------------------
--**********
	SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
	

	SET @DifferencesCode = ''
	SET @DifferencesDescription = ' '

	IF @intDelimiterPosField < @intDelimiterPosRecord  and  @intDelimiterPosField > 0
		Begin
			SET @ServePlace = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS INT)  from @TempTable )
		
		        Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField ) from @TempTable )
			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
	
			--Set @DifferencesCode =( select  LTrim(RTrim(SUBSTRING(nvcMainString , 1 , @intDelimiterPosRecord - 1)))  from @TempTable )

			WHILE @intDelimiterPosField < @intDelimiterPosRecord  and  @intDelimiterPosField > 0
				BEGIN
					SET @TempDifference  = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS INT)  from @TempTable )
					SET @DifferencesCode = @DifferencesCode + ';' + CAST (@TempDifference AS nvarchar(50))
					SET @DifferencesDescription = @DifferencesDescription + ' , ' + (SELECT RTRIM(LTRIM([Difference])) FROM tDifferences WHERE Code = @TempDifference)
		        		
					Update @TempTable  SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )
		
					SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
					SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)

				        	
				END
			SET @TempDifference = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1 , @intDelimiterPosRecord - 1))) AS INT)  from @TempTable )
			SET @DifferencesCode = @DifferencesCode + ';' + CAST (@TempDifference AS nvarchar(50))
			SET @DifferencesDescription = @DifferencesDescription + ' , ' + (SELECT RTRIM(LTRIM([Difference])) FROM tDifferences WHERE Code = @TempDifference)
		        
			Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosRecord + 1, DataLength(nvcMainString) - @intDelimiterPosRecord  ) from @TempTable )
			IF @DifferencesCode <> ''
				BEGIN
					Set @DifferencesCode = RIGHT (@DifferencesCode , LEN(@DifferencesCode) - 1)
					Set @DifferencesDescription = RIGHT (@DifferencesDescription , LEN(@DifferencesDescription) - 3)				
				End					
		END        
	ELSE		
		BEGIN
			SET @ServePlace = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString , 1 , @intDelimiterPosRecord - 1))) AS INT)  from @TempTable )
		
		      	Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosRecord + 1, DataLength(nvcMainString) - @intDelimiterPosRecord  ) from @TempTable )
	
		END

        INSERT INTO @ReturnTable(Amount , GoodCode , FeeUnit , Discount, Rate , ServePlace,ChairName ,[ExpireDate] , intInventoryNo ,DestInventoryNo ,  DifferencesCode ,DifferencesDescription) VALUES(@Amount, @GoodCode, @FeeUnit, @Discount , @Rate ,@ServePlace,@ChairName ,@ExpireDate ,@intInventoryNo , @DestInventoryNo , @DifferencesCode , @DifferencesDescription )
                
        SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable )
        SET @intDelimiterPosRecord = ( Select patindex('%/%' , nvcMainString)  from @TempTable )

    End

End

Return


End


GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

ALTER    PROCEDURE CheckPreSave
	@DetailsString NVARCHAR(4000)  ,
	@DetailsString2 NVARCHAR(4000) = NULL ,
	@DetailsString3 NVARCHAR(4000) = NULL ,
	@DetailsString4 NVARCHAR(4000) = NULL 
AS

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''


DECLARE @D1 NVARCHAR(Max) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(Max))  +  CAST(@DetailsString2 AS NVARCHAR(Max))  + CAST(@DetailsString3 AS NVARCHAR(Max))  + CAST(@DetailsString4 AS NVARCHAR(Max)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  

Declare @AccountYear Smallint
Declare @Branch int

Set @AccountYear = dbo.Get_AccountYear()
select @Branch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@D1))  
--Set @Branch = dbo.Get_Current_Branch()
Select * From
     (
select tGood.GoodType , MojodiControl , isnuLL(Tgood.Code,0) As GoodCode ,[Name] As GoodName ,  tUnitGood.[Description] ,Sum(FirstGoods.fltUsedValue * FirstGoods.Amount) As Used , Mojodi , Sum(FirstGoods.fltUsedValue * FirstGoods.Amount)- Mojodi  As Decrease 

     FROM
     tGood  inner join 
          (
	SELECT Amount , Goodcode , Serveplace ,IsNull(Code,GoodCode)As Code , IsNull(GoodfirstCode,Goodcode) as Goodfirstcode,IsNull(intServeplace,Serveplace)as intserveplace , IsNull(fltusedValue ,1) As fltUsedvalue , intInventoryNo 
		FROM dbo.Split(@D1) selectedGood   
	  left outer join (SELECT    Goodcode as code, GoodFirstCode, intServePlace, fltUsedValue FROM         dbo.tUsePercent)usepercent on selectedGood.GoodCode=usepercent.code and selectedGood.serveplace=usepercent.intserveplace
 
        )FirstGoods
 
         on FirstGoods.GoodFirstCode=tGood.code 

        inner join tInventory_Good On tGood.Code = tInventory_Good.GoodCode And tInventory_Good.InventoryNo = FirstGoods.intInventoryNo And tInventory_Good.AccountYear = @AccountYear And tInventory_Good.Branch = @Branch
		inner join tUnitGood On tGood.Unit = tUnitGood.Code 
        Group By  Tgood.Code , Tgood.[Name] , tUnitGood.[Description] , MojodiControl , Mojodi , GoodType
    )X

     where Decrease > 0 And MojodiControl = 1 And (GoodType = 3 OR GoodType = 4)




GO

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

--exec CheckPreSave N'5;11010002;30000;10;1;; ;1;;1/'
--GO


ALTER  Proc CheckPreSave_Edit 
(
@DetailsString NVARCHAR(4000),
@Status int,
@No BIGINT ,
	@DetailsString2 NVARCHAR(4000) = NULL ,
	@DetailsString3 NVARCHAR(4000) = NULL ,
	@DetailsString4 NVARCHAR(4000) = NULL 
	)
AS

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''
 
DECLARE @D1 NVARCHAR(Max) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(Max))  +  CAST(@DetailsString2 AS NVARCHAR(Max))  + CAST(@DetailsString3 AS NVARCHAR(Max))  + CAST(@DetailsString4 AS NVARCHAR(Max)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  

Declare @AccountYear Smallint
Declare @Branch int

Set @AccountYear = dbo.Get_AccountYear()

 select @Branch = branch from tInventory where inventoryNo=(SELECT Top 1 IntInventoryNo FROM Split(@D1))  
--Set @Branch = dbo.Get_Current_Branch()

Declare @intSerialNo BigInt
SET @intSerialNo = (SELECT tFacM.intSerialNo FROM tFacM WHERE [No] = @No AND Status = @Status and AccountYear = @AccountYear And Branch =  @Branch)

/*
--declare @DetailsString nvarchar(400)
--set @DetailsString='1;1122001;60000;0;1;; ;1/1;1114006;64000;0;1;; ;1/1;1110001;44000;0;1;; ;1/'

--SELECT * FROM dbo.Split(@DetailsString)

--
*/

select 
      used.firstcode,Mojodi,MojodiControl ,used.NewFacdAmount,isnull(used.tfacamount ,0) as currentfacdval,tGood.[Name] as GoodName,Mojodi- used.NewFacdAmount + isnull(used.tfacamount ,0) as remain , tUnitGood.[Description] 
    from 
       (  select  firstcode , Sum(fltused) As fltused, Sum(Amount) As Amount,Sum(fltused * Amount) As NewFacdAmount ,Sum(tfacamount) As tfacamount ,Max(intInventoryNo) As intInventoryNo   from 
	      (select tGood.code,ISNULL(GoodFirstCode,tGood.code)firstcode,ISNULL(fltUsedvalue,1) fltused,Amount , intInventoryNo from tGood 
			Inner join 
			  (   SELECT * FROM dbo.Split(@D1) selectedGood  
			    full outer join (SELECT    Goodcode as code, GoodFirstCode, intServePlace, fltUsedValue FROM         dbo.tUsePercent)usepercent on selectedGood.GoodCode=usepercent.code and selectedGood.serveplace=usepercent.intserveplace
	                      	where goodCode is  not null 
		      )selectedgood  on tGood.code=selectedgood.goodcode
	)X
	FULL OUTER JOIN
	    ( SELECT intserialNo,goodcode,ISNULL(GoodFirstCode,GoodCode) GoodFirstCode, ISNULL(usepercent.fltUsedValue,1) * ISNULL(tfacd.Amount,0) as tfacamount FROM tfacd   
		 LEFT OUTER join (SELECT    Goodcode as code, GoodFirstCode, intServePlace, fltUsedValue 
                                      FROM         dbo.tUsePercent)usepercent on tfacd.GoodCode=usepercent.code and tfacd.serveplace=usepercent.intserveplace
		                             where intserialNo = @intSerialNo
	     )Y
	on (Y.GoodCode=X.code And X.FirstCode = Y.GoodfirstCode) where code is not null Group By FirstCode 
	
      ) used inner join tGood on used.firstcode=tgood.code 
	     inner join tInventory_Good On tInventory_Good.GoodCode = tGood.Code And tInventory_Good.inventoryNo = used.intInventoryNo and tInventory_Good.AccountYear = @AccountYear And tInventory_Good.Branch =  @Branch
         inner join tUnitGood On tGood.Unit = tUnitGood.Code And Mojodi- used.NewFacdAmount + isnull(used.tfacamount ,0) < 0 And  MojodiControl = 1 And (tGood.GoodType = 3 OR tGood.GoodType = 4)

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER PROCEDURE [dbo].[InsertFactorMasterDetailsTemp]  (
                    @Status INT ,
                    @Owner INT ,
                    @Customer INT ,
                    @DiscountTotal FLOAT ,
                    @CarryFeeTotal FLOAT ,
                    @SumPrice FLOAT ,
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
                    @DetailsString NVARCHAR(4000)  ,
					@NvcDescription Nvarchar(150) ,
					@GuestNo INT = NULL  ,
					@Branch INT ,
					@TempAddress Nvarchar(255) = '', 
					@DetailsString2 NVARCHAR(4000) = NULL ,
					@DetailsString3 NVARCHAR(4000) = NULL ,
					@DetailsString4 NVARCHAR(4000) = NULL ,
                    @lastFacMNo BIGINT OUT
                     )

AS


IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''
 
DECLARE @D1 NVARCHAR(Max) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(Max))  +  CAST(@DetailsString2 AS NVARCHAR(Max))  + CAST(@DetailsString3 AS NVARCHAR(Max))  + CAST(@DetailsString4 AS NVARCHAR(Max)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  

DECLARE @No  INT
DECLARE @intSerialNo INT 
DECLARE @proper_time nvarchar(5)

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
    declare @newtime nvarchar(5)
    select @newtime=dbo.setTimeFormat(getdate())
    SELECT @MasterServePlace = SUM(tmpTable.SServePlace)
    FROM (  SELECT DISTINCT ServePlace As SServePlace
         FROM Split(@D1)
           ) tmpTable
    

     SET @NO = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacMTemp WHERE Status=@Status AND Branch =  @Branch )


     INSERT INTO tFacMTemp (
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
				NvcDescription  ,
				GuestNo ,
				Branch ,
				TempAddress
				
 )
     Values
(
                @NO ,
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
                dbo.Get_Shift(GETDATE()) ,
                @Incharge ,
                @owner ,
                @FacPayment ,
                @FacPayment ,
				@NvcDescription  ,
				@GuestNo ,
				@Branch ,
				@TempAddress
 )

    SET @intSerialNo=@@IDENTITY
     IF @@ERROR <>0
        GoTo EventHandler
    
     INSERT INTO tFacDTemp
(
    
	intRow,
	Amount ,
	GoodCode  ,
	FeeUnit ,
	Discount ,
	ChairName ,
	ExpireDate ,
	intInventoryNo ,
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
	tmpTable.ChairName ,
	tmpTable.ExpireDate ,
	tmpTable.intInventoryNo ,
	tmpTable.ServePlace ,
	tmpTable.DifferencesCode ,
	tmpTable.DifferencesDescription ,
	@intSerialNo , 
	@Branch 

     FROM (SELECT * FROM Split(@D1)) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode


     IF @@ERROR <>0
        GoTo EventHandler


set @lastFacMNo = @NO

COMMIT TRAN

Return @lastFacMNo

EventHandler:

    ROLLBACK TRAN
    SET @LastFacMNo = -1
    SELECT  -1 AS FACNO

    RETURN -1



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[EditFactorMasterDetailsTemp]  (
                    @No  BIGINT,
                    @Status INT ,
                    @Owner INT ,
                    @Customer INT ,
                    @DiscountTotal FLOAT ,
                    @CarryFeeTotal FLOAT ,
                    @SumPrice FLOAT ,
                    @Recursive INT ,
                    @InCharge INT ,
                    @FacPayment BIT ,
                    @OrderType INT ,
                    @StationId INT ,
                    @ServiceTotal FLOAT ,
                    @PackingTotal FLOAT ,
                    @TableNo INT ,
                    @Date Nvarchar(50) ,
                    @DetailsString NVARCHAR(4000),
					@NvcDescription Nvarchar(150) ,
					@GuestNo INT = NULL ,
					@Branch INT ,
					@TempAddress Nvarchar(255) = '', 
					@DetailsString2 NVARCHAR(4000) = NULL ,
					@DetailsString3 NVARCHAR(4000) = NULL ,
					@DetailsString4 NVARCHAR(4000) = NULL ,
                    @LastFacMNo BIGINT OUT
                     )

AS

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''

DECLARE @D1 NVARCHAR(Max) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(Max))  +  CAST(@DetailsString2 AS NVARCHAR(Max))  + CAST(@DetailsString3 AS NVARCHAR(Max))  + CAST(@DetailsString4 AS NVARCHAR(Max)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  

DECLARE @intSerialNo BIGINT
DECLARE  @FactorSerial BIGINT

SET @intSerialNo = (SELECT tFacMTemp.intSerialNo FROM tFacMTemp WHERE [No] = @No AND Status = @Status and Branch = @Branch )

IF  @Owner = 0
    SET @Owner = NULL

IF  @TableNo < 1
    SET @TableNo = NULL

IF  @Incharge < 1
    SET @Incharge = NULL

IF  @Customer=0
    SET @Customer = NULL

BEGIN TRANSACTION

     DECLARE @MasterServePlace INT

     SELECT @MasterServePlace = SUM(tmpTable.SServePlace)
     FROM (  SELECT DISTINCT ServePlace As SServePlace
         FROM Split(@D1)
           ) tmpTable



    DELETE FROM tFacDTemp
    WHERE tFacDTemp.intSerialNo = @intSerialNo and Branch = @Branch
 
 
     IF @@ERROR <>0
         GoTo EventHandler

    Update tFacMTemp
        SET Owner      	= @Owner,
        Customer       	= @Customer,
        DiscountTotal  	= @DiscountTotal,
        CarryFeeTotal  	= @CarryFeeTotal,
        SumPrice       	= @SumPrice,
        Recursive      	= @Recursive,
        InCharge       	= @InCharge,
        FacPayment     	= @FacPayment,    
        Balance      	= @FacPayment,
        OrderType      	= @OrderType,
        ServePlace     	= @MasterServePlace,
        StationId      	= @StationId,
        ServiceTotal   	= @ServiceTotal,
        PackingTotal   	= @PackingTotal,
        ShiftNo        	= dbo.Get_Shift(GETDATE()),
		NvcDescription  = @NvcDescription ,
        TableNo        	= @TableNo ,
        GuestNo			= @GuestNo ,
        TempAddress     = @TempAddress
        --[Date]         	= @Date,
        --[Time]         	=dbo.SetTimeFormat(GETDATE()),
        --[User]         	= @User,
        --RegDate		= dbo.Shamsi(GETDATE())
    WHERE tFacMTemp.intSerialNo = @intSerialNo and Branch = @Branch 

    IF @@ERROR <>0
        GoTo EventHandler


    INSERT INTO tFacDTemp
              (
	intRow,
	Amount ,
	GoodCode  ,
	FeeUnit ,
	Discount ,
	ChairName ,
	ExpireDate ,
	intInventoryNo  ,
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
	tmpTable.ChairName ,
	tmpTable.ExpireDate ,
	tmpTable.intInventoryNo ,
	tmpTable.ServePlace ,
	tmpTable.DifferencesCode ,
	tmpTable.DifferencesDescription ,
	@intSerialNo , 
	@Branch
    FROM (SELECT * FROM Split(@D1)) tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode

    IF @@ERROR <>0
        GoTo EventHandler

	
Set @LastFacMNo = @No

COMMIT TRANSACTION

Return @LastFacMNo

EventHandler:

    ROLLBACK TRAN
    SET @LastFacMNo = -1
    RETURN @LastFacMNo



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE [dbo].[InsertFactorDetail]  (
	 @DetailsString NVARCHAR(Max) ,
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

	DECLARE @InventoryNo INT 
	select @InventoryNo=  (SELECT TOP 1  intInventoryNo FROM Split(@DetailsString))      
	DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

	If @Status = 6 AND @intSerialNo2 > 0 AND @DestinventoryNo > 0
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
	If @Status = 2 AND @intSerialNo2 > 0 
	Begin

        DECLARE @ReturnTable TABLE
            (
              Row INT IDENTITY(1, 1)
                      NOT NULL,
              Amount FLOAT NOT NULL,
              GoodCode INT NOT NULL,
              BuyPrice FLOAT NOT NULL
            )
 
        INSERT  INTO @ReturnTable
                (
                  Amount,
                  GoodCode,
                  BuyPrice
                )
                SELECT  CAST(SUM(T.Amount) AS DECIMAL(19,3)) ,
                        T.GoodCode,
                        T.BuyPrice
                FROM    ( SELECT    dbo.tUsePercent.GoodFirstCode AS GoodCode,
                                    ( dbo.tFacD.Amount
                                      * ( dbo.tUsePercent.fltUsedValue
                                          + ISNULL(dbo.tUsePercent.Pert, 0) ) ) AS Amount,
                                    ( SELECT    BuyPrice
                                      FROM      dbo.tGood
                                      WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                    ) AS BuyPrice
                          FROM      dbo.tFacM
                                    JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                      AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    JOIN dbo.tUsePercent ON dbo.tFacD.GoodCode = dbo.tUsePercent.GoodCode
                                                            AND dbo.tFacD.ServePlace = dbo.tUsePercent.intServePlace
                                    JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                          WHERE     tfacm.Branch = @Branch AND tfacM.intSerialNo = @intSerialNo
                                    AND ( SELECT    dbo.tGood.GoodType
                                          FROM      dbo.tGood
                                          WHERE     dbo.tGood.Code = dbo.tUsePercent.GoodFirstCode
                                        ) <> 4
                          UNION ALL
                          SELECT    dbo.tFacD.GoodCode,
                                    dbo.tFacD.Amount,
                                    dbo.tGood.BuyPrice
                          FROM      dbo.tFacM
                                    JOIN dbo.tFacD ON dbo.tFacM.Branch = dbo.tFacD.Branch
                                                      AND dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
                                    JOIN dbo.tGood ON dbo.tFacD.GoodCode = dbo.tGood.Code
                          WHERE     dbo.tFacM.Branch = @Branch AND tfacM.intSerialNo = @intSerialNo
                                    AND dbo.tFacD.GoodCode NOT IN (
                                    SELECT  dbo.tUsePercent.GoodCode
                                    FROM    dbo.tUsePercent
                                    WHERE   dbo.tUsePercent.intServePlace = dbo.tFacD.ServePlace )
                                    AND dbo.tGood.GoodType = 3
                        ) T
                GROUP BY T.GoodCode,
                        T.BuyPrice

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
				tmpTable.BuyPrice ,
				0 , --tmpTable.Discount ,
				1 , --tmpTable.Rate ,
				NULL , --tmpTable.ChairName ,
				'' , --tmpTable.[ExpireDate],
				@InventoryNo   ,
				NULL , --tmpTable.DestInventoryNo ,
				1 , --tmpTable.ServePlace ,
				'', --tmpTable.DifferencesCode ,
				'', --tmpTable.DifferencesDescription ,
				@intSerialNo2 , 
				@Branch --dbo.Get_Current_Branch()
		
		
			FROM @ReturnTable tmpTable INNER JOIN tGood ON tGood.code = tmpTable.GoodCode
	
	end

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER    PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

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
            @DetailsString NVARCHAR(4000),      
            @ds nText = '',      
            @Balance BIT ,      
            @AccountYear smallint = null  ,       
            @NvcDescription Nvarchar(150) = Null ,      
            @HavaleNo int = Null  ,      
            @TempAddress Nvarchar(255) = '',  
			@GuestNo INT,    
			@DetailsString2 NVARCHAR(4000) = NULL ,
			@DetailsString3 NVARCHAR(4000) = NULL ,
			@DetailsString4 NVARCHAR(4000) = NULL ,
            @lastFacMNo INT OUT  ,
		    @Person INT = NULL     
             )      

AS      

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''

DECLARE @D1 NVARCHAR(Max) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(Max))  +  CAST(@DetailsString2 AS NVARCHAR(Max))  + CAST(@DetailsString3 AS NVARCHAR(Max))  + CAST(@DetailsString4 AS NVARCHAR(Max)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  


Declare @intserialNo int      
Declare @intserialNo2 int      
--Declare @intserialNo3 Bigint    

SET @intserialNo = 0        
SET @intserialNo2   = 0      
--SET @intserialNo3   = 0      

DECLARE @No1  INT     
DECLARE @No2  INT     
--DECLARE @No3  INT     

DECLARE @SumPrice  FLOAT       
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
         FROM Split(@D1)      
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
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@D1))      

	select @destbranch=  @intBranch --   branch from tInventory where inventoryNo=(SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

    DECLARE @DestStatus INT 	
    IF @Status = 2
        AND dbo.AutoHavale() = 1 
 	SELECT  @DestStatus = 6 ,
                 @NO2 =   ISNULL(MAX([NO]), 0) + 1
                    FROM    tFacM
                    WHERE   Status = 6
                            AND Branch = @intBranch
                            AND AccountYear = @AccountYear
                        
   IF ( @Status = 6
             AND [dbo].[AutoResid]() = 1 AND @DestinventoryNo > 0
           ) 
	        SELECT  --@Customer = NULL , 
	        	@DestStatus = 7 ,
	                @NO2 = ISNULL(MAX([NO]), 0) + 1
	                    FROM    tFacM
	                    WHERE   Status = 7
	                            AND Branch = @destbranch
	                            AND AccountYear = @AccountYear
	
    IF ( @Status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1
             AND @DestinventoryNo > 0
           ) 
  BEGIN
 
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
                @DestStatus,      
                @Customer ,      
                @SumPrice ,      
                @OrderType ,      
                1 , --@MasterServePlace ,      
                @StationId ,      
                0 , --@ServiceTotal ,      
                @Recursive ,      
                0 , --@CarryFeeTotal ,      
                0 , --@PackingTotal ,      
                0 , --@DiscountTotal ,      
                0 ,
                0 ,      
                @newtime,      
                @User ,      
                @TableNo,      
                @ShiftNo , --dbo.Get_Shift(GETDATE()) ,      
                NULL , --@Incharge ,      
                @owner ,      
                @FacPayment ,      
                @Balance ,      
				@DestBranch ,     
				@AccountYear ,      
				@NvcDescription,      
				@TempAddress,
				0 , --@GuestNo ,
				NULL --@TempNo2    
		
 )      
		 IF @@ERROR <>0      
			GoTo EventHandler      
		SET @intserialNo2 = @IdentityNo + 1      

             IF @status = 2
                BEGIN 
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' حواله - '
                        + CAST(@No2 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo AND Branch = @intBranch
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' فاکتور فروش  - '
                        + CAST(@No1 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo2 AND Branch = @intBranch
                UPDATE  tfacm
                SET     RefrenceHavale = @intserialNo2
                WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch
				
				END 
            IF @status = 6
             BEGIN 
               UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' رسيد - '
                        + CAST(@No2 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo
                UPDATE  tfacm
                SET     NvcDescription = @NvcDescription + N' حواله  - '
                        + CAST(@No1 AS NVARCHAR(8))
                WHERE   intSerialNo = @intserialNo2 AND Branch = @intBranch
                UPDATE  tfacm
                SET     RefrenceHavale = @intserialNo2
                WHERE   intSerialNo = @intserialNo  AND Branch = @intBranch
			END 

end      


----------------------------------Fill Details Factor  --------------------------------------------------------------      
 exec InsertFactorDetail @D1 , @intserialNo , @intserialNo2, @Customer , @intBranch      

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

Declare @SumPrice2 FLOAT       
Set @SumPrice2 = (Select Cast(Sum(Amount * FeeUnit) as FLOAT ) From tFacd Where intSerialNo = @intserialNo2 And Branch = @DestBranch )        
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
IF @Status = 6 AND @DestinventoryNo > 0 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 7 , @User , 1 , @AccountYear , @destbranch      
     IF @@ERROR <>0      
   GoTo EventHandler      
IF @Status = 2 AND dbo.AutoHavale() = 1 --AND (@destbranch = @intBranch ) --or dbo.AutoResid() = 1     
 Exec InsertHistory  @No2, 6 , @User , 1 , @AccountYear , @destbranch      
     IF @@ERROR <>0      
   GoTo EventHandler      


----------------------------Cash ---------------------------      

------------------------Mojodi Control Online--------------------------------------------      

Exec InsertMojodiCalculate @Status ,  @intserialNo , @AccountYear , @intBranch      
IF @@ERROR <>0      
 GoTo EventHandler      
 IF ( @Status = 2
     AND dbo.AutoHavale() = 1
   )
    OR ( @Status = 6
         AND [dbo].[AutoResid]() = 1  AND @DestinventoryNo > 0
       ) 
 BEGIN      
	 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch      
	 IF @@ERROR <>0      
	 GoTo EventHandler      

	    EXEC InsertMojodiCalculate @DestStatus, @intserialNo2, @AccountYear, @intBranch
	 IF @@ERROR <>0      
	 GoTo EventHandler      
 END       
IF dbo.AutoHavale() = 1
        UPDATE  tfacm
        SET     [BitHavaleResid] = 1
        WHERE   [intSerialNo] = @intserialNo AND Branch = @intBranch

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
ALTER      PROCEDURE [dbo].[EditFactorMasterDetails]  (  


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
	@DetailsString  NVARCHAR(4000),  
	@ds nText = '',  
	@Balance Bit,  
	@AccountYear Smallint = Null ,  
	@NvcDescription Nvarchar(150) = Null ,  
	@TempAddress Nvarchar(255) = '', 
	@GuestNo INT,     
	@DetailsString2 NVARCHAR(4000) = NULL ,
	@DetailsString3 NVARCHAR(4000) = NULL ,
	@DetailsString4 NVARCHAR(4000) = NULL ,
	@LastFacMNo  INT OUT  ,
	@Person INT = NULL 
  )  


AS 

IF @DetailsString2 IS NULL SET @DetailsString2 = N''
IF @DetailsString3 IS NULL SET @DetailsString3 = N''
IF @DetailsString4 IS NULL SET @DetailsString4 = N''

DECLARE @D1 NVARCHAR(Max) 
SET @D1 = CAST(@DetailsString AS NVARCHAR(Max))  +  CAST(@DetailsString2 AS NVARCHAR(Max))  + CAST(@DetailsString3 AS NVARCHAR(Max))  + CAST(@DetailsString4 AS NVARCHAR(Max)) 
--SET @D1 = @DetailsString  +  @DetailsString2 + @DetailsString3 + @DetailsString4  
 
DECLARE @SumPrice FLOAT  
DECLARE @SumPrice2 FLOAT  
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
 (  SELECT DISTINCT ServePlace As SServePlace  FROM Split(@D1)) tmpTable  


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

DECLARE @DestinventoryNo INT 
	select @DestinventoryNo=  (SELECT TOP 1  DestInventoryNo FROM Split(@D1))      

    DECLARE @DestStatus INT 
    DECLARE @BitHavaleResid INT 
    IF ( @status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1  --AND  @DestinventoryNo > 0 
           ) 
        SELECT  @intSerialNo2 = ISNULL(RefrenceHavale , 0) ,
		        @BitHavaleResid = ISNULL(BitHavaleResid , 0) 
                               FROM      dbo.tFacM
                              WHERE     intSerialNo = @intSerialNo
                                        AND Branch = @intBranch

        SELECT  @DestStatus = Status 
                               FROM      dbo.tFacM
                              WHERE     intSerialNo = @intSerialNo2
                                        AND Branch = @intBranch


 select @destbranch= @intBranch -- branch from tInventory where inventoryNo=(SELECT TOP 1 DestInventoryNo FROM Split(@DetailsString))  

---------------------------------------Mojodi Control Online---------------------------------------------------------  
Exec DeleteMojodiCalculate @Status , @intserialNo  ,  1 , @AccountYear , @intBranch  
IF @@ERROR <>0  
 GoTo EventHandler  
    IF ( @status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1 --AND @DestinventoryNo > 0  Because AutoHavale is without destination
           ) 
        EXEC DeleteMojodiCalculate @DestStatus, @intserialNo2, 1, @AccountYear, @intBranch

    IF @@ERROR <> 0 
        GOTO EventHandler
 ----------------------------------------Delete Old Details -----------------------------------------------------------  
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
If  @intSerialNo2 > 0--And (@destbranch = @intBranch or dbo.AutoResid() = 1 )   
    DELETE FROM tFacD  
    WHERE tFacD.intSerialNo = @intSerialNo2 AND Branch =  @intBranch  
    IF @@ERROR <>0  
        GoTo EventHandler  
------------------------------------------------------------    
  Exec DeleteFactorChildren @intSerialNo , @intBranch  
  IF @@ERROR <>0  
         GoTo EventHandler  
 If @intSerialNo2 > 0--And (@destbranch = dbo.Get_Current_Branch() or dbo.AutoResid() = 1 )  
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

    IF ( @status = 2
         AND dbo.AutoHavale() = 1
       )
        OR ( @Status = 6
             AND [dbo].[AutoResid]() = 1 AND  @DestinventoryNo > 0 
           ) 
    BEGIN

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
    WHERE tFacM.intSerialNo = @intSerialNo2  AND Branch =  @intBranch  

END  

----------------------------------Fill Details Factor ----------------------------------------------------------------------  
 exec InsertFactorDetail @D1 , @intserialNo , @intserialNo2, @Customer , @intBranch  

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

If @intSerialNo2 > 0  --And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1 )  
   Set @SumPrice2 = (Select Cast (Sum(Amount * FeeUnit) as FLOAT )   From tFacd Where intSerialNo = @intSerialNo2 And Branch = @intBranch )    
   IF @@ERROR <>0  
        GoTo EventHandler  
PRINT @SumPrice2
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
If @intSerialNo2 > 0--And (@destbranch= dbo.Get_Current_Branch() Or dbo.AutoResid() = 1)   

    Update tFacm Set SumPrice = @SumPrice2 Where intSerialNo = @intserialNo2 And Branch = @intBranch  
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
	IF ( @status = 2  AND dbo.AutoHavale() = 1)
		OR ( @Status = 6   AND [dbo].[AutoResid]() = 1 AND  @DestinventoryNo > 0    ) 
	
	 BEGIN  
	 Exec Insert_tinventory_Good @intserialNo2 , @AccountYear , @intBranch  
	 IF @@ERROR <>0  
	 GoTo EventHandler  

	EXEC InsertMojodiCalculate @DestStatus, @intserialNo2, @AccountYear, @intBranch
    IF @@ERROR <> 0 
        GOTO EventHandler
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


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  PROCEDURE dbo.Update_BuyPrice_by_LastPrice
AS 

UPDATE tgood
SET BuyPrice = T2.FeeUnit

from (
SELECT T.intserialNo , T.GoodCode , FeeUnit
FROM tfacd INNER JOIN
(
SELECT MAX(dbo.tFacD.intSerialNo) AS intserialNo  , GoodCode  FROM tfacm
    INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
    WHERE Status = 1 AND AccountYear = dbo.Get_AccountYear()  AND tfacd.Branch = dbo.Get_Current_Branch()
    GROUP BY GoodCode
    )T
ON T.GoodCode = dbo.tFacD.GoodCode AND T.intserialNo = dbo.tFacD.intSerialNo 
--ORDER BY T.GoodCode
) T2

WHERE tGood.code = T2.GoodCode AND T2.FeeUnit > 0

GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO




ALTER  proc Transport_tblTotal_tGood_By_Prams 
(
	@Level1 int ,
	@strSelectedLevels nvarchar(4000) , 
	@InventoryNo int ,
	@Branch int ,
	@AccountYear SMALLINT,
	@CheckNotZeroMojodi BIT,
	@CheckFirstMojodi	BIT,
	@CountingNo	INT,
	@ToOtherAccountYear SMALLINT
)
	
as
BEGIN TRAN

	DELETE tInventory_Good
	FROM
	(
		SELECT vw_Good.* , tInventory_Good.* 
		
		FROM 
			[dbo].[vw_Good] 
			Inner Join  
			dbo.tInventory_Good ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode 
		WHERE 
			(LEVEL1 = @Level1 OR @Level1=-1)
			And (InventoryNo = @InventoryNo OR @InventoryNo=-1)
			And (Branch = @Branch OR @Branch=-1)
			And (AccountYear = @ToOtherAccountYear OR @ToOtherAccountYear=-1)
			AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels=N'')
			AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
			AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
	)AS T
	WHERE T.GoodCode=tInventory_Good.GoodCode 
		AND T.InventoryNo=tInventory_Good.InventoryNo
		AND T.Branch=tInventory_Good.Branch
		AND T.AccountYear=tInventory_Good.AccountYear

if @@Error <> 0 
	Goto ErrHandler

	INSERT INTO tInventory_Good
		(
		      InventoryNo, Branch, GoodCode, FirstMojodi, Mojodi, MojodiControl, OrderPoint, MinValue, MaxValue, [Date], [Time], BuyAmount, SaleAmount, 
                      LossAmount, BuyReturnAmount, SaleReturnAmount, FromStoreAmount, toStoreAmount, AccountYear, Counting1, 
                      Counting2, Counting3, CountDifference
		)

		SELECT     InventoryNo, Branch, GoodCode,CASE @CountingNo 
								WHEN 0 THEN CAST(ISNULL(T.Mojodi  ,0) AS DECIMAL(20,3))
								WHEN 1 THEN ISNULL(T.Counting1,0)
								WHEN 2 THEN ISNULL(T.Counting2,0)
								WHEN 3 THEN ISNULL(T.Counting3,0)
								ELSE ISNULL(T.Mojodi,0)
								END
							AS FirstMojodi, 0, 0, 0, MinValue, MaxValue, [Date], [Time], 0, 0, 
	                      0, 0, 0, 0, 0, @ToOtherAccountYear, 0, 
	                      0, 0, 0
		FROM
		(
			SELECT vw_Good.* , tInventory_Good.* 
			
			FROM 
				[dbo].[vw_Good] 
				Inner Join  
				dbo.tInventory_Good ON dbo.vw_Good.Code = dbo.tInventory_Good.GoodCode 
			WHERE 
				(LEVEL1 = @Level1 OR @Level1=-1)
				And (InventoryNo = @InventoryNo OR @InventoryNo=-1)
				And (Branch = @Branch OR @Branch=-1)
				And (AccountYear = @AccountYear OR @AccountYear=-1)
				AND (LEVEL2 in ( select cast( t.word as int) from dbo.SplitWithDelimiterNVarChar(@strSelectedLevels , ',')t ) OR @strSelectedLevels='')
				AND ((dbo.tInventory_Good.Mojodi <>CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckNotZeroMojodi WHEN 1 THEN 0 ELSE -1 END))
				AND ((dbo.tInventory_Good.FirstMojodi <>CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @CheckFirstMojodi WHEN 1 THEN 0 ELSE -1 END))
		)AS T
	
if @@Error <> 0 
	Goto ErrHandler

	Insert into dbo.tStation_Inventory_Good ( branch ,InventoryNo, AccountYear ,StationID,  GoodCode , Active)
	
	select Branch ,InventoryNo ,@ToOtherAccountYear ,StationID ,GoodCode ,Active 
		From tStation_Inventory_Good 
	        Where   inventoryno = @inventoryno and Branch = @Branch and AccountYear = @AccountYear

if @@Error <> 0 
	Goto ErrHandler

EXEC dbo.Update_BuyPrice_by_LastPrice
if @@Error <> 0 
	Goto ErrHandler

UPDATE tInventory_Good
SET FirstMojodi = t.Buyprice
FROM
(SELECT code , Buyprice FROM tgood)t
WHERE dbo.tInventory_Good.GoodCode = t.Code AND tInventory_Good.AccountYear  = @AccountYear AND tInventory_Good.InventoryNo = @InventoryNo


Commit Tran
Return

ErrHandler:
RollBack Tran
Return



GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  FUNCTION dbo.BarcodeGenerator
(
	@ServePlace INT,
	@FactorNo   INT,
	@Type	BIT = 0 
)
RETURNS  NVARCHAR(20)

AS

BEGIN

IF @Type IS NULL SET @Type = 0

	DECLARE @strServePlace NVARCHAR(10)
	DECLARE @strFactorNo     NVARCHAR(10)
	DECLARE @Tmp                NVARCHAR(20)
	DECLARE @ZeroCount      INT


	SET @ServePlace = @ServePlace + 10
	SET @ZeroCount = 9 - LEN(CAST(@FactorNo AS NVARCHAR(9)))

	SET @strFactorNo = (SELECT dbo.Repeater('0',@ZeroCount)) + CAST(@FactorNo AS NVARCHAR(9))
	

	IF ((@ServePlace / 10)<1)
		SET @strServePlace = '0' + CAST(@ServePlace AS NVARCHAR(10))
	ELSE
		SET @strServePlace = CAST(@ServePlace AS NVARCHAR(10))


	SET @Tmp = @strServePlace + CAST(@Type AS NVARCHAR(1)) + @strFactorNo


	SET @Tmp = '*' + @Tmp  + '*'




RETURN(@Tmp)

END

GO


-- Delete All Rows this Report
DELETE FROM tblTotal_ItemReports_Details WHERE intReportId = 71 
GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
          71 ,
          1 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 1 AND ROW = 1
GO


INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
          71 ,
          2 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 7 AND ROW = 1
GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
          71 ,
          3 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 7 AND ROW = 4
GO


INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
          71 ,
          4 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 9 AND ROW = 4
GO

INSERT INTO dbo.tblTotal_ItemReports_Details
        ( intReportId ,
          Row ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
        )
SELECT 
          71 ,
          5 ,
          FromText ,
          toText ,
          ParameterName ,
          ParameterType ,
          parameterLengh ,
          ObjectType ,
          Quantity ,
          MinValue ,
          MaxValue ,
          ComboQuery ,
          ComboFieldCode ,
          ComboFieldDescr ,
          RighttoLeft
FROM tblTotal_ItemReports_Details WHERE intReportId = 2 AND ROW = 5
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


--ems server\arya

ALTER   procedure GetInventoryGood_Mojodi_All(  
 @SystemDate  	NVARCHAR(50),
 @SystemDay   	NVARCHAR(50),
 @SystemTime  	NVARCHAR(50),
 @Date1 NVARCHAR(8),
 @Date2 NVARCHAR(8),
 @AccountYear1 SMALLINT,
 @Inventory1    Int ,  
 @Flag1 INT ,
 @Level11 INT ,
 @Level12 INT )  

as  

IF @Level11 = -1 SELECT @Level11 = Min(code) FROM dbo.tGoodLevel1
IF @Level12 = -1 SELECT @Level12 = Max(code) FROM dbo.tGoodLevel1

 DECLARE @Level11Name AS NVARCHAR(20)
 DECLARE @Level12Name AS NVARCHAR(20)
 
 SELECT  @Level11Name = [description] FROM tGoodLevel1 WHERE Code = @Level11
 SELECT @Level12Name = [description] FROM tGoodLevel1 WHERE Code = @Level12
 
 IF @Flag1 = 1 
 BEGIN 
  select @SystemDate AS SystemDate ,@SystemDay AS SystemDay , @SystemTime AS SystemTime , intInventoryNo ,InventoryName , GoodCode,  
	 	Name ,SUM(Amountv) AS Amountv, SUM(Amounts) AS Amounts ,SUM(Totalfeev) AS Totalfeev ,SUM(Totalfees) AS Totalfees, 
		(FirstMojodi + SUM(Amountv) - SUM(Amounts) ) AS  Mojodi ,FirstMojodi , FirstPrice , FirstMojodiPrice
--	  	, CASE (FirstMojodi + SUM(Amountv) - SUM(Amounts) ) WHEN 0 THEN 0 ELSE (FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) ) END AS MojodiPrice 
--	  	, CASE (FirstMojodi + SUM(Amountv) - SUM(Amounts) ) WHEN 0 THEN 0 ELSE ((FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) ) / (FirstMojodi + SUM(Amountv) - SUM(Amounts)))END  AS MojodiPriceFee
		,CASE t.Mojodi WHEN 0 THEN 0 ELSE FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) END  AS MojodiPrice  --t.Mojodi *  t.MojodiPrice
		, t.MojodiPrice AS MojodiPriceFee
	  	, CASE SUM(Amountv) WHEN 0 THEN 0 ELSE  (SUM(Totalfeev) / SUM(Amountv) ) END  AS Feev
	 	, CASE SUM(Amounts) WHEN 0 THEN 0 ELSE  (SUM(Totalfees) / SUM(Amounts) ) END  AS Fees
	   	,NamePrn,BarCode , @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay]  
		, @Level11Name AS Level11Name,  @Level12Name AS Level12Name , T.level1Code, T.level1Name
	  from 
	  (
	  	SELECT    
		dbo.tInventory_Good.InventoryNo AS  intInventoryNo
		,Isnull(dbo.tInventory.[Description],'') as InventoryName ,dbo.tInventory_Good.GoodCode ,tGood.Name ,  
		case when  tFacM.Status in (1,7)   then  Sum(tFacD.Amount ) else 0 end AS Amountv,  
		case when  tFacM.Status in (4,6,3)   then  Sum(tFacD.Amount ) else 0 end AS Amounts,    
		case when  tFacM.Status in (1,7)   then  Sum(tFacD.Amount * dbo.tFacD.FeeUnit * (1 - (tFacD.Discount/100)) ) else 0 end AS Totalfeev,  
		case when  tFacM.Status in (4,6,3)   then  Sum(tFacD.Amount * dbo.tFacD.FeeUnit * (1 - (tFacD.Discount/100)) ) else 0 end AS Totalfees,    
		
		tInventory_Good.FirstMojodi   ,  tInventory_Good.FirstPrice , (tInventory_Good.FirstMojodi  * tInventory_Good.FirstPrice) AS FirstMojodiPrice ,
		tGood.NamePrn,tGood.BarCode , Mojodi , MojodiPrice 
		, dbo.tGoodLevel1.Code AS level1Code, dbo.tGoodLevel1.Description AS level1Name
		
		FROM  dbo.tFacM   
		INNER join  dbo.tFacD ON  dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND  dbo.tFacM.Branch = dbo.tFacD.Branch  AND tFacD.intInventoryNo = @Inventory1 AND dbo.tFacM.Recursive = 0   AND dbo.tFacM.AccountYear = @AccountYear1
		INNER join dbo.tInventory_Good ON tInventory_Good.GoodCode = tFacD.GoodCode AND dbo.tInventory_Good.AccountYear = @AccountYear1 AND InventoryNo = @Inventory1
		INNER join tInventory On tInventory.InventoryNo = tInventory_Good.InventoryNo    
		INNER JOIN dbo.tGood ON  dbo.tGood.Code = dbo.tInventory_Good.GoodCode AND dbo.tGood.GoodType = 3 AND tGood.Level1 >= @Level11 AND tGood.Level1 <= @Level12 
		INNER JOIN dbo.tGoodLevel1 ON  dbo.tGood.Level1 = dbo.tGoodLevel1.Code  
		INNER join  dbo.tUnitGood ON  dbo.tGood.Unit = dbo.tUnitGood.Code  
		WHERE dbo.tFacM.Date >= @Date1 AND dbo.tFacM.Date <= @Date2
		Group By  FirstMojodi  , tGood.Name ,tinventory_Good.InventoryNo,tGood.NamePrn,   
		tInventory_Good.Goodcode ,tGood.BarCode   , dbo.tInventory.[Description], tInventory_Good.FirstPrice , tFacM.Status , Mojodi , MojodiPrice , dbo.tGoodLevel1.Code, dbo.tGoodLevel1.Description
	  	--, dbo.tblSale_FacD.FeeUnit
	  )T
	Group By  FirstMojodi  , [Name] , intInventoryNo,NamePrn,   
		Goodcode ,BarCode   , InventoryName , FirstPrice , FirstMojodiPrice , Mojodi , MojodiPrice  , T.level1Code, T.level1Name
	Order BY  GoodCode  

 END  
ELSE IF @Flag1 = 2 
 BEGIN 
	 select @SystemDate AS SystemDate ,@SystemDay AS SystemDay , @SystemTime AS SystemTime , intInventoryNo ,InventoryName , GoodCode,  
	 	Name ,SUM(Amountv) AS Amountv, SUM(Amounts) AS Amounts ,SUM(Totalfeev) AS Totalfeev ,SUM(Totalfees) AS Totalfees, 
		(FirstMojodi + SUM(Amountv) - SUM(Amounts) ) AS  Mojodi ,FirstMojodi , FirstPrice , FirstMojodiPrice
--	  	, CASE (FirstMojodi + SUM(Amountv) - SUM(Amounts) ) WHEN 0 THEN 0 ELSE (FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) ) END AS MojodiPrice 
--	  	, CASE (FirstMojodi + SUM(Amountv) - SUM(Amounts) ) WHEN 0 THEN 0 ELSE ((FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) ) / (FirstMojodi + SUM(Amountv) - SUM(Amounts)))END  AS MojodiPriceFee
		, CASE t.Mojodi WHEN 0 THEN 0 ELSE FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) END  AS MojodiPrice  --t.Mojodi *  t.MojodiPrice
		, T.MojodiPrice AS MojodiPriceFee
	  	, CASE SUM(Amountv) WHEN 0 THEN 0 ELSE  (SUM(Totalfeev) / SUM(Amountv) ) END  AS Feev
	 	, CASE SUM(Amounts) WHEN 0 THEN 0 ELSE  (SUM(Totalfees) / SUM(Amounts) ) END  AS Fees
	   	,NamePrn,BarCode  , @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay]    
		, @Level11Name AS Level11Name,  @Level12Name AS Level12Name , T.level1Code, T.level1Name
	  from 
	  (
	  	SELECT    
		dbo.tInventory_Good.InventoryNo AS  intInventoryNo
		,Isnull(dbo.tInventory.[Description],'') as InventoryName ,dbo.tInventory_Good.GoodCode, tGood.Name ,  
		case when  tfacm.Status in (1,7)   then  Sum(tFacd.Amount ) else 0 end AS Amountv,  
		case when  tfacm.Status in (4,6,3)   then  Sum(tFacd.Amount ) else 0 end AS Amounts,    
		case when  tfacm.Status in (1,7)   then  Sum(tFacd.Amount * dbo.tFacD.FeeUnit * (1 - (tFacd.Discount/100)) ) else 0 end AS Totalfeev,  
		case when  tfacm.Status in (4,6,3)   then  Sum(tFacd.Amount * dbo.tFacD.FeeUnit * (1 - (tFacd.Discount/100)) ) else 0 end AS Totalfees,    
		
		tInventory_Good.FirstMojodi   ,  tInventory_Good.FirstPrice , (tInventory_Good.FirstMojodi  * tInventory_Good.FirstPrice) AS FirstMojodiPrice ,
		tgood.NamePrn,tgood.BarCode   , Mojodi , MojodiPrice 
		, dbo.tGoodLevel1.Code AS level1Code, dbo.tGoodLevel1.Description AS level1Name
		
		FROM  dbo.tInventory_Good   
		INNER JOIN dbo.tGood ON  dbo.tGood.Code = dbo.tInventory_Good.GoodCode AND dbo.tgood.GoodType = 3 AND tGood.Level1 >= @Level11 AND tGood.Level1 <= @Level12 
		INNER JOIN dbo.tGoodLevel1 ON  dbo.tGood.Level1 = dbo.tGoodLevel1.Code  
		INNER join tInventory On tInventory.InventoryNo = tInventory_Good.InventoryNo AND dbo.tInventory_Good.InventoryNo = @Inventory1  
		INNER join  dbo.tUnitGood ON  dbo.tGood.Unit = dbo.tUnitGood.Code  
		left outer  JOIN dbo.tFacD ON tInventory_Good.GoodCode = tFacd.GoodCode  
		left outer JOIN dbo.tFacM ON  dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND  dbo.tFacM.Branch = dbo.tFacD.Branch   
			AND   dbo.tFacM.Recursive = 0   AND dbo.tFacM.AccountYear = @AccountYear1  And tFacD.intInventoryNo = @Inventory1   
			AND dbo.tFacM.Date >= @Date1 AND dbo.tFacM.Date <= @Date2	
		--AND tfacm.Status IN (1,3,4,5,6,7) 
		WHERE  	 dbo.tInventory_Good.InventoryNo = @Inventory1 --AND(Mojodi <> 0 ) -- OR FirstMojodi <>0 OR SaleAmount<>0 OR BuyAmount<>0)
			And tInventory_Good.AccountYear = @AccountYear1  
           -- AND ((dbo.tInventory_Good.Mojodi <>CASE @Flag1 WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @Flag1 WHEN 1 THEN 0 ELSE -1 END))  
		
		Group By  FirstMojodi  , tGood.Name ,  tInventory_Good.InventoryNo,tgood.NamePrn,   
		tInventory_Good.Goodcode ,tgood.BarCode   , dbo.tInventory.[Description], tInventory_Good.FirstPrice , tfacm.Status , Mojodi , MojodiPrice , dbo.tGoodLevel1.Code, dbo.tGoodLevel1.Description
	  	--, dbo.tFacD.FeeUnit
	  )T
	
	Group By  FirstMojodi  , [Name] , intInventoryNo,NamePrn,   
		Goodcode ,BarCode   , InventoryName , FirstPrice , FirstMojodiPrice , Mojodi , MojodiPrice , MojodiPrice  , T.level1Code, T.level1Name
	Order by GoodCode  

END 

 ELSE IF @Flag1 = 3 
 BEGIN 
	 select @SystemDate AS SystemDate ,@SystemDay AS SystemDay , @SystemTime AS SystemTime , intInventoryNo ,InventoryName , GoodCode,  
	 	Name ,SUM(Amountv) AS Amountv, SUM(Amounts) AS Amounts ,SUM(Totalfeev) AS Totalfeev ,SUM(Totalfees) AS Totalfees, 
		(FirstMojodi + SUM(Amountv) - SUM(Amounts) ) AS  Mojodi ,FirstMojodi , FirstPrice , FirstMojodiPrice
--	  	, CASE (FirstMojodi + SUM(Amountv) - SUM(Amounts) ) WHEN 0 THEN 0 ELSE (FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) ) END AS MojodiPrice 
--	  	, CASE (FirstMojodi + SUM(Amountv) - SUM(Amounts) ) WHEN 0 THEN 0 ELSE ((FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) ) / (FirstMojodi + SUM(Amountv) - SUM(Amounts)))END  AS MojodiPriceFee
		, CASE t.Mojodi WHEN 0 THEN 0 ELSE FirstMojodiPrice + SUM(Totalfeev) - SUM(Totalfees) END  AS MojodiPrice  --t.Mojodi *  t.MojodiPrice
		, T.MojodiPrice AS MojodiPriceFee
	  	, CASE SUM(Amountv) WHEN 0 THEN 0 ELSE  (SUM(Totalfeev) / SUM(Amountv) ) END  AS Feev
	 	, CASE SUM(Amounts) WHEN 0 THEN 0 ELSE  (SUM(Totalfees) / SUM(Amounts) ) END  AS Fees
	   	,NamePrn,BarCode  , @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay]    
		, @Level11Name AS Level11Name,  @Level12Name AS Level12Name , T.level1Code, T.level1Name
	  from 
	  (
	  	SELECT    
		dbo.tInventory_Good.InventoryNo AS  intInventoryNo
		,Isnull(dbo.tInventory.[Description],'') as InventoryName ,dbo.tInventory_Good.GoodCode, tGood.Name ,  
		case when  tfacm.Status in (1,7)   then  Sum(tFacd.Amount ) else 0 end AS Amountv,  
		case when  tfacm.Status in (4,6,3)   then  Sum(tFacd.Amount ) else 0 end AS Amounts,    
		case when  tfacm.Status in (1,7)   then  Sum(tFacd.Amount * dbo.tFacD.FeeUnit * (1 - (tFacd.Discount/100)) ) else 0 end AS Totalfeev,  
		case when  tfacm.Status in (4,6,3)   then  Sum(tFacd.Amount * dbo.tFacD.FeeUnit * (1 - (tFacd.Discount/100)) ) else 0 end AS Totalfees,    
		
		tInventory_Good.FirstMojodi   ,  tInventory_Good.FirstPrice , (tInventory_Good.FirstMojodi  * tInventory_Good.FirstPrice) AS FirstMojodiPrice ,
		tgood.NamePrn,tgood.BarCode   , Mojodi , MojodiPrice 
		, dbo.tGoodLevel1.Code AS level1Code, dbo.tGoodLevel1.Description AS level1Name
		
		FROM  dbo.tInventory_Good   
		INNER JOIN dbo.tGood ON  dbo.tGood.Code = dbo.tInventory_Good.GoodCode AND dbo.tgood.GoodType = 3  AND tGood.Level1 >= @Level11 AND tGood.Level1 <= @Level12
		INNER JOIN dbo.tGoodLevel1 ON  dbo.tGood.Level1 = dbo.tGoodLevel1.Code  
		INNER join tInventory On tInventory.InventoryNo = tInventory_Good.InventoryNo AND dbo.tInventory_Good.InventoryNo = @Inventory1  
		INNER join  dbo.tUnitGood ON  dbo.tGood.Unit = dbo.tUnitGood.Code  
		left outer  JOIN dbo.tFacD ON tInventory_Good.GoodCode = tFacd.GoodCode  
		left outer JOIN dbo.tFacM ON  dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo AND  dbo.tFacM.Branch = dbo.tFacD.Branch   
			AND   dbo.tFacM.Recursive = 0   AND dbo.tFacM.AccountYear = @AccountYear1  And tFacD.intInventoryNo = @Inventory1   
			AND dbo.tFacM.Date >= @Date1 AND dbo.tFacM.Date <= @Date2	
	--AND tfacm.Status IN (1,3,4,5,6,7) 
		WHERE  	 dbo.tInventory_Good.InventoryNo = @Inventory1 AND(Mojodi > 0) -- OR FirstMojodi <>0 OR SaleAmount<>0 OR BuyAmount<>0
			And tInventory_Good.AccountYear = @AccountYear1  
	                AND ((dbo.tInventory_Good.Mojodi <>CASE @Flag1 WHEN 1 THEN 0 ELSE -1 END) OR (-1=CASE @Flag1 WHEN 1 THEN 0 ELSE -1 END))  
		
		Group By  FirstMojodi  , tGood.Name ,  tInventory_Good.InventoryNo,tgood.NamePrn,   
		tInventory_Good.Goodcode ,tgood.BarCode   , dbo.tInventory.[Description], tInventory_Good.FirstPrice , tfacm.Status , Mojodi , MojodiPrice , dbo.tGoodLevel1.Code, dbo.tGoodLevel1.Description
	  	--, dbo.tFacD.FeeUnit
	  )T
	
	Group By  FirstMojodi  , [Name] , intInventoryNo,NamePrn,   
		Goodcode ,BarCode   , InventoryName , FirstPrice , FirstMojodiPrice , Mojodi , MojodiPrice  , T.level1Code, T.level1Name
	Order by GoodCode  

END 


GO




SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER    PROCEDURE [dbo].[InsertMojodiCalculate]
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

        END
--===============================================

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER PROCEDURE [dbo].[DeleteMojodiCalculate]
    (
      @Status INT ,
      @intserialNo BIGINT ,
      @Recursive INT ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS ---------------------------------------Mojodi Control Online---------------------------------------------------------
    IF @Recursive = 1 
        BEGIN
            IF @Status = 2 
                BEGIN
		--IF dbo.AutoHavale() = 0
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            SaleAmount = SaleAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + Amount ,
                            SaleAmount = SaleAmount - Amount
                    FROM    ( SELECT    SUM(( Amount * fltUsedValue )
                                            + ( [Amount] * [Pert] )) AS Amount ,
                                        GoodFirstCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      ( SELECT    *
                                          FROM      tFacd
                                                    INNER JOIN UsePercent ON tFacd.GoodCode = usepercent.code
                                                              AND tFacd.serveplace = usepercent.intserveplace
                                          WHERE     intserialNo = @intserialNo
                                                    AND Branch = @Branch
                                        ) FirstGoods
                                        INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code AND dbo.tGood.GoodType = 4 
                              GROUP BY  FirstGoods.GoodFirstCode ,
                                        FirstGoods.intInventoryNo ,
                                        FirstGoods.Branch
                            ) X
                    WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                            AND tInventory_Good.InventoryNo = X.intInventoryNo
                            AND tInventory_Good.Branch = X.Branch
                            AND tInventory_Good.AccountYear = @AccountYear

                END
            IF @Status = 1 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - t.Amount ,
                            BuyAmount = BuyAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 3 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            LossAmount = LossAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 4 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi + t.Amount ,
                            BuyReturnAmount = BuyReturnAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
            IF @Status = 5 
                BEGIN
                    UPDATE  tInventory_Good
                    SET     Mojodi = Mojodi - t.Amount ,
                            SaleReturnAmount = SaleReturnAmount - t.Amount
                    FROM    ( SELECT    SUM(Amount) AS Amount ,
                                        GoodCode ,
                                        intInventoryNo ,
                                        Branch
                              FROM      tFacD
                              WHERE     tFacD.intSerialNo = @intSerialNo
                                        AND Branch = @Branch
                              GROUP BY  GoodCode ,
                                        intInventoryNo ,
                                        Branch
                            ) t
                    WHERE   tInventory_Good.Goodcode = t.Goodcode
                            AND tInventory_Good.InventoryNo = t.intInventoryNo
                            AND tInventory_Good.Branch = t.Branch
                            AND tInventory_Good.AccountYear = @AccountYear
	
                END
        END

    ELSE 
        IF @Recursive = 0 
            BEGIN
                IF @Status = 2 
                    BEGIN
	   	    --IF dbo.AutoHavale() = 0
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - t.Amount ,
                                SaleAmount = SaleAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - Amount ,
                                SaleAmount = SaleAmount - Amount
                        FROM    ( SELECT    SUM(( Amount * fltUsedValue )
                                                + ( [Amount] * [Pert] )) AS Amount ,
                                            GoodFirstCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      ( SELECT    *
                                              FROM      tFacd
                                                        INNER JOIN UsePercent ON tFacd.GoodCode = usepercent.code
                                                              AND tFacd.serveplace = usepercent.intserveplace
                                              WHERE     intserialNo = @intserialNo
                                                        AND Branch = @Branch
                                            ) FirstGoods
                                            INNER JOIN tGood ON FirstGoods.GoodFirstCode = tGood.code  AND dbo.tGood.GoodType = 4  
                                  GROUP BY  FirstGoods.GoodFirstCode ,
                                            FirstGoods.intInventoryNo ,
                                            FirstGoods.Branch
                                ) X
                        WHERE   X.GoodFirstCode = tInventory_Good.Goodcode
                                AND tInventory_Good.InventoryNo = X.intInventoryNo
                                AND tInventory_Good.Branch = X.Branch
                                AND tInventory_Good.AccountYear = @AccountYear

                    END
                IF @Status = 1 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi + t.Amount ,
                                BuyAmount = BuyAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
                IF @Status = 3 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - t.Amount ,
                                LossAmount = LossAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
                IF @Status = 4 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi - t.Amount ,
                                BuyReturnAmount = BuyReturnAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
                IF @Status = 5 
                    BEGIN
                        UPDATE  tInventory_Good
                        SET     Mojodi = Mojodi + t.Amount ,
                                SaleReturnAmount = SaleReturnAmount + t.Amount
                        FROM    ( SELECT    Amount ,
                                            GoodCode ,
                                            intInventoryNo ,
                                            Branch
                                  FROM      tFacD
                                  WHERE     tFacD.intSerialNo = @intSerialNo
                                            AND Branch = @Branch
                                ) t
                        WHERE   tInventory_Good.Goodcode = t.Goodcode
                                AND tInventory_Good.InventoryNo = t.intInventoryNo
                                AND tInventory_Good.Branch = t.Branch
                                AND tInventory_Good.AccountYear = @AccountYear
	
                    END
            END
--===============================================

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  PROCEDURE [dbo].[Update_tblTotal_tInventory_tGood_For_Mojodi]
    (
      @intLanguage INT,
      @SystemDate NVARCHAR(50),
      @SystemDay NVARCHAR(50),
      @SystemTime NVARCHAR(50),
      @DateBefore NVARCHAR(50),
      @DateAfter NVARCHAR(50),
      @Type INT,
      @InventoryNo1 INT,
      @InventoryNo2 INT,
      @Branch INT,
      @UsePercentFlag INT,
      @AccountYear SMALLINT
    )
AS 
    BEGIN TRAN

    SET @SystemTime = dbo.SetTimeFormat(GETDATE())

    INSERT  INTO tInventory_Good
            (
              Branch,
              InventoryNo,
              GoodCode,
              BuyAmount,
              SaleAmount,
              LossAmount,
              BuyReturnAmount,
              SaleReturnAmount,
              FromStoreAmount,
              toStoreAmount,
              Mojodi,
              AccountYear 
            )
            SELECT  T1.Branch,
                    T1.intInventoryNo,
                    T1.GoodCode,
                    T1.BuyAmount,
                    T1.SaleAmount,
                    T1.LossAmount,
                    T1.BuyReturnAmount,
                    T1.SaleReturnAmount,
                    T1.FromStoreAmount,
                    T1.toStoreAmount,
					CASE WHEN @Type = 1
								THEN T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount
								  - T1.FromStoreAmount - T1.LossAmount
								  + T1.ToStoreAmount  
						 WHEN @Type = 3
								THEN T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount
								  - T1.FromStoreAmount - T1.LossAmount
								  + T1.ToStoreAmount
						ELSE T1.Mojodi  END  AS Mojodi,
                    @AccountYear
				FROM    dbo.tblTotal_tInventory_tGood_For_Mojodi(@intLanguage,
                                                             @SystemDate,
                                                             @SystemDay,
                                                             @SystemTime,
                                                             @DateBefore,
                                                             @DateAfter, @Type,
                                                             @InventoryNo1,
                                                             @InventoryNo2,
                                                             @Branch,
                                                             @UsePercentFlag,
                                                             @AccountYear) AS T1

--SELECT T1.GoodCode,T1.firstMojodi+T1.BuyAmount - T1.BuyReturnAmount - T1.FromStoreAmount - T1.LossAmount+ T1.ToStoreAmount FROM  dbo.tblTotal_tInventory_tGood_For_Mojodi(0,'','','',N'89/01/01',N'89/02/23',1,1,1,1,0,1389) T1
                                                          
            WHERE   0 = ( SELECT    COUNT(GoodCode)
                          FROM      tInventory_Good
                          WHERE     GoodCode = T1.GoodCode
                                    AND InventoryNo = T1.intInventoryNo
                                    AND Branch = T1.Branch
                                    AND AccountYear = @AccountYear
                        )
		

    IF @@Error <> 0 
        GOTO ErrHandler
------------------------------------------------------------------------

    UPDATE  tInventory_Good
    SET     BuyAmount = T2.BuyAmount,
            SaleAmount = T2.SaleAmount,
            LossAmount = T2.LossAmount,
            BuyReturnAmount = T2.BuyReturnAmount,
            SaleReturnAmount = T2.SaleReturnAmount,
            FromStoreAmount = T2.FromStoreAmount,
            toStoreAmount = T2.toStoreAmount,
            Mojodi = 	 CASE WHEN @Type = 1
								 THEN T2.firstMojodi+T2.BuyAmount - T2.BuyReturnAmount
									  - T2.FromStoreAmount - T2.LossAmount
									  + T2.ToStoreAmount
							 WHEN @Type = 3
								 THEN T2.firstMojodi+T2.BuyAmount - T2.BuyReturnAmount
									  - T2.FromStoreAmount - T2.LossAmount
									  + T2.ToStoreAmount
							ELSE T2.Mojodi    
					END  

			FROM    dbo.tblTotal_tInventory_tGood_For_Mojodi(@intLanguage, @SystemDate,
                                                     @SystemDay, @SystemTime,
                                                     @DateBefore, @DateAfter,
                                                     @Type, @InventoryNo1,
                                                     @InventoryNo2, @Branch,
                                                     @UsePercentFlag,
                                                     @AccountYear) AS T2
			WHERE   tInventory_Good.GoodCode = T2.GoodCode
					AND tInventory_Good.InventoryNo = T2.intInventoryNo
					AND tInventory_Good.Branch = T2.Branch
					AND tInventory_Good.AccountYear = @AccountYear
---------------------------------------------------------------------
    IF @@Error <> 0 
        GOTO ErrHandler

    UPDATE  tInventory_Good
    SET     Mojodi = 0
    FROM    ( SELECT    
                    GoodCode ,
                    InventoryNo ,
                    Branch
          FROM      tInventory_Good INNER JOIN dbo.tGood ON dbo.tInventory_Good.GoodCode = dbo.tGood.Code AND GoodType = 4
          WHERE     tInventory_Good.AccountYear = @AccountYear
                    AND tInventory_Good.Branch = @Branch
                                ) T3
    WHERE   tInventory_Good.GoodCode = T3.GoodCode
            AND tInventory_Good.InventoryNo = T3.InventoryNo
            AND tInventory_Good.Branch = T3.Branch
            AND tInventory_Good.AccountYear = @AccountYear
	    AND tInventory_Good.Mojodi < 0
---------------------------------------------------------------------
    IF @@Error <> 0 
        GOTO ErrHandler

    COMMIT TRAN 

    RETURN 1

    ErrHandler:
    ROLLBACK TRAN
    RETURN -1
	

GO
