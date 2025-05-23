
--Script_V26_16_Fix13
-- تغییر در گروههای منوهای کالاها
-- اضافه شدن تعداد ردیف گروهها در سایر تنظیمات
-- اضافه شدن منو های تاچ و غیر تاچ در سایر تنظیمات
--کنترل اینکه در ایستگاه فقط منوهای اکتیو ایستگاهها را نشان دهد 
--کنترل اینکه در ایستگاه فقط منوهای ایستگاهها را نشان دهد و برای تبلت ظاهر نشود تا شماره گروه 8
--کنترل گروههای تبلت که فقط گروههای تبلت را نشان دهد از شماره گروه 9 به بعد
-- کپی منو از یک ایستگاه به ایستگاه دیگر
--اضافه کردن میز به صورت گروهی  
--94/04/04

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
          13
        )
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Copy_tGood_Menu') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Copy_tGood_Menu
GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

CREATE PROC Copy_tGood_Menu

@StationId INT ,
@NewStationId INT ,
@intStatus INT OUT 
 
AS 

BEGIN TRAN
SET @intStatus = 0

DELETE FROM tNameDisp WHERE StationId = @NewStationId AND Branch = dbo.Get_Current_Branch()
if @@Error <> 0 
	Goto ErrHandler

INSERT INTO dbo.tNameDisp
        ( StationId ,
          BtnNum ,
          FactorType ,
          NameDisp ,
          LatinNameDisp ,
          PicturePath ,
          Branch
        )
SELECT 
		  @NewStationId ,
          BtnNum ,
          FactorType ,
          NameDisp ,
          LatinNameDisp ,
          PicturePath ,
          Branch
FROM dbo.tNameDisp WHERE StationId = @StationId AND Branch = dbo.Get_Current_Branch()
if @@Error <> 0 
	Goto ErrHandler

DELETE FROM dbo.tGood_Menu WHERE StationId = @NewStationId AND Branch = dbo.Get_Current_Branch()
if @@Error <> 0 
	Goto ErrHandler

INSERT INTO dbo.tGood_Menu
        ( GoodCode ,
          StationId ,
          FactorType ,
          BtnNum ,
          Branch
        )
SELECT 
		GoodCode ,
          @NewStationId ,
          FactorType ,
          BtnNum ,
          Branch
FROM dbo.tGood_Menu WHERE StationId = @StationId AND Branch = dbo.Get_Current_Branch()

if @@Error <> 0 
	Goto ErrHandler

Commit Tran

SET @intStatus= 1
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return



GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Copy_tTables') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Copy_tTables
GO

CREATE PROC Copy_tTables
(
    @FromTableNo    int  ,
    @NumberOfChair     INT,
    @Person     INT,
    @PartitionID    INT,   
    @Empty    Bit,
    @Reserve    Bit,
    @nvcMaxUseTime  NVARCHAR(10) ,
    @Branch INT ,
    @ToTableNo    int  ,
    @intStatus         INT Out

)
AS

BEGIN TRAN

SET @intStatus = -1

    IF @Person < 1
        SET @Person =null

DECLARE @MaxTableNo INT

    Set @MaxTableNo = (Select isnull(Max(No) , 0) + 1 from dbo.tTable      WHERE    Branch  = @Branch )
   
    IF @MaxTableNo < @Branch * 1000 SET @MaxTableNo = @Branch * 1000
    SET @FromTableNo = @FromTableNo + 1

    WHILE @FromTableNo <= @ToTableNo
    BEGIN
        BEGIN  
        PRINT @MaxTableNo
        PRINT CAST(@FromTableNo AS VARCHAR(3))
        PRINT @NumberOfChair
        PRINT @Person
        PRINT @PartitionID
        PRINT @Empty
        PRINT @Reserve
        PRINT @Branch
        PRINT @nvcMaxUseTime
		
		SET ROWCOUNT 1
		
        INSERT INTO dbo.tTable     ( No , [Name] , NumberOfChair , Person ,  PartitionID , Empty , Reserve ,Branch , nvcMaxUseTime)
                    SELECT         @MaxTableNo , CAST(@FromTableNo AS VARCHAR(3)) , @NumberOfChair , @Person ,  @PartitionID ,@Empty , @Reserve ,  @Branch ,@nvcMaxUseTime
					FROM dbo.tTable where @MaxTableNo NOT IN (SELECT No FROM dbo.tTable WHERE Branch = @Branch)
         
        IF @@Error <> 0 GOTO ErrHandler
        SET @MaxTableNo = @MaxTableNo + 1
        SET @FromTableNo = @FromTableNo + 1

        END

       IF @FromTableNo > @ToTableNo
          BREAK
       ELSE
          CONTINUE
    END
     
COMMIT TRAN

SET @intStatus = 1
RETURN @intStatus

ErrHandler:
    ROLLBACK TRAN
    SET @intStatus = -1
    RETURN @intStatus



GO


--declare @P1 int
--set @P1=-1
--exec Copy_tTables 27, 4, 4, 1, 1, 0, '60', 1, 40, @P1 output
--select @P1

