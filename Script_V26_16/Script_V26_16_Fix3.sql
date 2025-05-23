


--Script_V26_16_Fix3

--تغییر در فرم دریافت و پرداخت برای حسابداری 
--اضافه شدن دریافت از مشتریان و تامین کنندگان و پرداخت به مشتریان و تامین کنندگان و پرداخت هزینه از صندوق به تولید سند
-- اتصال حسابهای بانکی تعریف شده در حسابداری به پوز های بانکی
--ثبت گزارش وجوه نقدی و بانکی کاربران در هنگام زدن کلید گزارش صندوق
--امکان مشاهده یا ثبت مجدد گزارش کاربران در تولید سند حسابداری
--کنترل اتصال حسابهای بانکی به پوز بانکی
-- کنترل عدم ثبت تکراری گزارشات بر اساس کاربر و روز و شیفت و  پوزبانکی
-- دریافت وجه از صندوقدار در فرم تولید سند
--ثبت کسر و اضافات صندوق در قرم تولید سند
--امکان ثبت کالای تکراری در سطر جدید فاکتور خرید

-- 92/11/18

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
          3
        )
GO


ALTER TABLE dbo.tblAcc_Recieved
ADD SanadNo INT NULL 
GO

ALTER TABLE dbo.tblAcc_Cash
ADD SanadNo INT NULL 
GO



--=================Bank Pos====================================
ALTER TABLE [dbo].[tblPub_Pos]
ADD AccountId INT NULL 
go 


ALTER PROCEDURE [dbo].[Insert_tblPub_Pos] 
(
	@NvcPosNo nvarchar(20) , 
	@nvcBankName nvarchar(20) , 
	@nvcAccountNo nvarchar(20) , 
	@AccountId INT ,
	@intStatus int out)
AS

declare @PosId int

Begin Tran

SELECT @PosId = IsNull(Max(PosId) + 1, 1) FROM tblPub_Pos

Insert Into dbo.tblPub_Pos
        ( PosId ,
          NvcPosNo ,
          nvcBankName ,
          nvcAccountNo ,
          AccountId
        )
VALUES  ( @PosId , -- PosId - int
          @NvcPosNo , -- NvcPosNo - nvarchar(20)
          @nvcBankName , -- BankName - nvarchar(20)
          @nvcAccountNo , -- nvcAccountNo - nvarchar(20)
          @AccountId
        )

if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus=@PosId
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return


GO



ALTER PROCEDURE [dbo].[Update_tblPub_Pos] (
	@PosId INT ,
	@NvcPosNo nvarchar(20) , 
	@nvcBankName nvarchar(20) , 
	@nvcAccountNo nvarchar(20) ,
	@AccountId INT , 
	@intStatus int out)

AS

Begin Tran

UPDATE dbo.tblPub_Pos SET
	NvcPosNo = @NvcPosNo  , 
	nvcBankName = @nvcBankName , 
	nvcAccountNo = @nvcAccountNo ,
	AccountId = @AccountId 

   WHERE PosId = @PosId

if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus = 1
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return



GO



--==========================================================

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_PaymentType_Acc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_PaymentType_Acc
GO

CREATE PROCEDURE [dbo].[Get_PaymentType_Acc] AS
Select * from tblAcc_PaymentType 
WHERE Code = 0 OR Code = 5 OR Code = 6 
 


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_RecievedType_Acc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_RecievedType_Acc]
GO

CREATE PROCEDURE [dbo].[Get_RecievedType_Acc] AS
Select * from tblAcc_RecievedType 
WHERE Code = 3 OR Code = 4 
 



GO




ALTER   VIEW dbo.Vw_tblacc_Recieved_User
AS  
	SELECT  dbo.tblAcc_Recieved.Code,
            dbo.tblAcc_Recieved.[No],
            dbo.tblAcc_Recieved.List,
            dbo.tblAcc_Recieved.[Date],
            dbo.tblAcc_Recieved.RegDate,
            dbo.tblAcc_Recieved.RegTime,
            dbo.tblAcc_Recieved.UID,
            dbo.tblAcc_Recieved.Description,
            dbo.tblAcc_Recieved.Bestankar,
            dbo.tblacc_Recieved.Branch,
            dbo.tPer.Tafsili,
            dbo.tPer.nvcFirstName + ' ' + dbo.tPer.nvcSurName AS [User_Name],
            dbo.tPer.Gender,
            dbo.tblAcc_Recieved.RecieveType,
            dbo.tblacc_Recieved.Code_Bes,
            CASE WHEN RecieveType IN ( 0, 1, 2 )
                 THEN ( SELECT  dbo.tPer.Tafsili
                        FROM    dbo.tPer
                        WHERE   dbo.tper.Ppno = dbo.tblAcc_Recieved.Code_Bes
                               -- AND dbo.tPer.Branch = dbo.tblAcc_Recieved.Branch
                      )
                 WHEN RecieveType = 3
                 THEN ( SELECT  ISNULL(dbo.tCust.Tafsili, 0)
                        FROM    dbo.tCust
                        WHERE   dbo.tCust.Code = dbo.tblAcc_Recieved.Code_Bes
                               -- AND dbo.tCust.Branch = dbo.tblAcc_Recieved.Branch
                      )
                 WHEN RecieveType = 4
                 THEN ( SELECT  dbo.tSupplier.Tafsili
                        FROM    dbo.tSupplier
                        WHERE   dbo.tSupplier.Code = dbo.tblAcc_Recieved.Code_Bes
                               -- AND dbo.tSupplier.Branch = dbo.tblAcc_Recieved.Branch
                      )
            END AS Person_Tafsili,
            CASE WHEN RecieveType IN ( 0, 1, 2 )
                 THEN ( SELECT  dbo.tPer.nvcFirstName + ' '
                                + dbo.tPer.nvcSurName
                        FROM    dbo.tPer
                        WHERE   dbo.tper.Ppno = dbo.tblacc_Recieved.Code_Bes
                               -- AND dbo.tPer.Branch = dbo.tblAcc_Recieved.Branch
                      )
                 WHEN RecieveType = 3
                 THEN ( SELECT  dbo.vw_Customers.FullName
                        FROM    dbo.vw_Customers
                        WHERE   dbo.vw_Customers.Code = dbo.tblacc_Recieved.Code_Bes
                               -- AND dbo.vw_Customers.Branch = dbo.tblAcc_Recieved.Branch
                      )
                 WHEN RecieveType = 4
                 THEN ( SELECT  dbo.vw_Suppliers.FullName
                        FROM    dbo.vw_Suppliers
                        WHERE   dbo.vw_Suppliers.Code = dbo.tblacc_Recieved.Code_Bes
                              --  AND dbo.vw_Suppliers.Branch = dbo.tblAcc_Recieved.Branch
                      )
            END AS Person_Name,
            dbo.tblAcc_Recieved.AddUser,
            dbo.tblAcc_Recieved.AccountYear,
            ISNULL([tCust].[MembershipId], -1) AS MembershipId,
            dbo.tblAcc_Recieved.transferAccounting
    FROM    dbo.tblAcc_Recieved
            INNER JOIN dbo.tUser ON dbo.tblAcc_Recieved.UID = dbo.tUser.UID
            INNER JOIN dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno
            LEFT OUTER JOIN [tCust] ON [tCust].[Code] = [tblAcc_Recieved].[Code_Bes]
                                      -- AND dbo.tCust.Branch = dbo.tblAcc_Recieved.Branch 
	WHERE 	tblAcc_Recieved.intSerialNo IS NULL 



GO




ALTER  PROCEDURE [dbo].[Get_UserRecieve]
    (
      @RecieveType INT ,
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @Uid INT 
    )
AS 
    SELECT  SUM(Vw_tblacc_Recieved_User.Bestankar) AS TotalBestankar ,
            Vw_tblacc_Recieved_User.[Uid] ,
            Vw_tblacc_Recieved_User.[User_Name] ,
            Vw_tblacc_Recieved_User.Tafsili ,
            Vw_tblacc_Recieved_User.[Date] ,
            Vw_tblacc_Recieved_User.[RecieveType] ,
            Vw_tblacc_Recieved_User.[Code_Bes] ,
            Vw_tblacc_Recieved_User.[Person_Tafsili] ,
            Vw_tblacc_Recieved_User.[Person_Name]
    FROM    Vw_tblacc_Recieved_User
    WHERE   Vw_tblacc_Recieved_User.[Date] >= @DateBefore
            AND Vw_tblacc_Recieved_User.[Date] <= @DateAfter
            AND Vw_tblacc_Recieved_User.[RecieveType] = @RecieveType
			AND dbo.Vw_tblacc_Recieved_User.transferAccounting=0
			AND (Vw_tblacc_Recieved_User.[UID] = @Uid OR  @Uid = 0 )
    GROUP BY Vw_tblacc_Recieved_User.[Date] ,
            Vw_tblacc_Recieved_User.[Uid] ,
            Vw_tblacc_Recieved_User.[User_Name] ,
            Vw_tblacc_Recieved_User.Tafsili ,
            Vw_tblacc_Recieved_User.[RecieveType] ,
            Vw_tblacc_Recieved_User.[Code_Bes] ,
            Vw_tblacc_Recieved_User.[Person_Tafsili] ,
            Vw_tblacc_Recieved_User.[Person_Name]
    ORDER BY Vw_tblacc_Recieved_User.[Date]




GO




ALTER  PROCEDURE [dbo].[Get_UserPayment]
    (
      @PaymentType INT ,
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @Uid INT 
    )
AS 
    SELECT  SUM(Vw_tblAcc_Cash_User.Bestankar) AS TotalBestankar ,
            Vw_tblAcc_Cash_User.[Uid] ,
            Vw_tblAcc_Cash_User.[User_Name] ,
            Vw_tblAcc_Cash_User.Tafsili ,
            Vw_tblAcc_Cash_User.[Date] ,
            Vw_tblAcc_Cash_User.[PaymentType] ,
            Vw_tblAcc_Cash_User.[Uid_Bede] ,
            Vw_tblAcc_Cash_User.[Person_Tafsili] ,
            Vw_tblAcc_Cash_User.[Person_Name]
    FROM    Vw_tblAcc_Cash_User
    WHERE   Vw_tblAcc_Cash_User.[Date] >= @DateBefore
            AND Vw_tblAcc_Cash_User.[Date] <= @DateAfter
            AND Vw_tblAcc_Cash_User.[PaymentType] = @PaymentType
			AND dbo.Vw_tblacc_Cash_User.transferAccounting=0
			AND (Vw_tblacc_Cash_User.UID = @Uid OR @Uid = 0)
    GROUP BY Vw_tblAcc_Cash_User.[Date] ,
            Vw_tblAcc_Cash_User.[Uid] ,
            Vw_tblAcc_Cash_User.[User_Name] ,
            Vw_tblAcc_Cash_User.Tafsili ,
            Vw_tblAcc_Cash_User.[PaymentType] ,
            Vw_tblAcc_Cash_User.[Uid_Bede] ,
            Vw_tblAcc_Cash_User.[Person_Tafsili] ,
            Vw_tblAcc_Cash_User.[Person_Name]
    ORDER BY Vw_tblAcc_Cash_User.[Date]



GO




--اضافه شدن آپدیت دریافت و پرداخت 
ALTER  PROC Update_transferAccounting
(
  @Branch INT ,
  @DateBefore NVARCHAR(8) ,
  @DateAfter NVARCHAR(8),
  @SanadNo INT ,
  @Uid INT 
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
		--AND (Customer < 0 OR Customer IS NULL OR  (Customer > 0 AND Credit = 0))  --لازم نیست چون فاکتور مشتریان قبلا سند حسابداری خورده 
		AND (tfacm.[User] = @Uid OR @Uid = 0)

	UPDATE dbo.tblAcc_Recieved
	SET dbo.tblAcc_Recieved.transferAccounting=1	,
		dbo.tblAcc_Recieved.SanadNo = @SanadNo
	WHERE tblAcc_Recieved.Branch = @Branch
		AND tblAcc_Recieved.[Date] >= @DateBefore
		AND tblAcc_Recieved.[Date] <= @DateAfter
		AND transferAccounting = 0
		AND (RecieveType = 3 OR RecieveType = 4) 
		AND (tblAcc_Recieved.UID = @Uid OR @Uid = 0)
		AND intSerialNo IS NULL 


	UPDATE dbo.tblAcc_Cash
	SET dbo.tblAcc_Cash.transferAccounting = 1	,
		dbo.tblAcc_Cash.SanadNo = @SanadNo
	WHERE tblAcc_Cash.Branch = @Branch
		AND tblAcc_Cash.[Date] >= @DateBefore
		AND tblAcc_Cash.[Date] <= @DateAfter
		AND transferAccounting = 0
		AND (PaymentType = 0 OR PaymentType = 5 OR PaymentType = 6) 
		AND (tblAcc_Cash.UID = @Uid OR @Uid = 0)


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_tRecvType_Acc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_tRecvType_Acc]
GO



CREATE  PROCEDURE [dbo].[Get_All_tRecvType_Acc] AS
SELECT *
FROM tRecvType
WHERE (tintIsShow = 1)
AND tintType = 1 OR tintType = 5
ORDER BY tintType


GO

UPDATE  tRecvType
SET nvcDescription = N'کارت بانکی' 
WHERE tintType = 5


GO




ALTER    PROCEDURE [dbo].[Get_AccountDocument]
    (
      @Branch INT ,
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @Code INT ,
      @Uid INT 
    )
AS 
    IF ( @Code = 1 ) 
            --BEGIN
      --          SELECT  tFacM.[Date] ,
      --                  tPer.nvcFirstName + ' ' + tPer.nvcSurName AS UserFullName ,
      --                  tPer.Tafsili AS PersonTafsili 
      --                 --, ISNULL(SUM(ISNULL(tFacCash.intAmount , 0)), 0)  AS sp
      --                 ,CASE WHEN dbo.tFacM.Status =2 THEN SUM(dbo.tFacM.SumPrice) ELSE -1 * SUM(dbo.tFacM.SumPrice) END AS sp
      --          FROM    tFacM
      --                  INNER JOIN tUser ON tUser.UID = tFacM.[User]
      --                  INNER JOIN tPer ON tUser.pPno = tPer.pPno
						--INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
      --          WHERE    tFacM.Branch = @Branch 
      --                  AND tFacM.Recursive = 0 
      --                  AND (tFacM.Status = 2 OR tFacM.Status = 5 )
      --                  AND dbo.tFacM.transferAccounting=0  
						--AND (tfacm.[User] = @Uid OR @Uid = 0)
						--AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
      --          GROUP BY tFacM.[Date] ,tFacM.Status ,
      --                  tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
      --                  tPer.Tafsili 
      --          HAVING  tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
      --          ORDER BY tFacM.[Date] ,tPer.Tafsili
      --      END

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
                        INNER JOIN dbo.tCust ON tfacm.Customer = dbo.tCust.Code
						LEFT OUTER JOIN 
							(Select SUM(IsNull(Bestankar,0)) AS Bestankar1 , intSerialNo , Branch From   tblAcc_Recieved GROUP BY intSerialNo , Branch )t1
								ON  t1.intSerialNo = dbo.tFacM.intSerialNo  and t1.Branch = dbo.tFacM.Branch 	
                WHERE    tFacM.Branch = @Branch 
                        AND tFacM.Recursive = 0 
                        AND tFacM.Status = 2
                        AND dbo.tFacM.transferAccounting=0  
                        AND (tfacm.[User] = @Uid OR @Uid = 0)
                        AND (Customer < 0  OR Customer IS NULL OR (Customer > 0 AND Credit = 0))
                GROUP BY tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName ,
                        tPer.Tafsili 
                HAVING  tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,
                        tPer.Tafsili
            END

    IF ( @Code = 4 ) 
            BEGIN
                SELECT  tFacM.[Date] ,
                        tPer.nvcFirstName + ' ' + tPer.nvcSurName + ' - ' + tblPub_Pos.nvcAccountNo AS nvcDescription ,
                        ISNULL(tblPub_Pos.AccountId ,0) AS TafsiliId  ,
                        ISNULL(SUM(tFacCard.intAmount), 0) AS sp
                FROM    tFacM
                        INNER JOIN dbo.tUser ON tUser.UID = tFacM.[User]
                                           -- AND tUser.Branch = tFacM.Branch
                        INNER JOIN tPer ON tUser.pPno = tPer.pPno
                        INNER JOIN tFacCard ON tFacM.Branch = tFacCard.Branch
                                               AND tFacM.intSerialNo = tFacCard.intSerialNo
                        INNER JOIN tblPub_Pos ON dbo.tblPub_Pos.PosId = dbo.tFacCard.PosId
                WHERE    tFacM.Branch = @Branch 
                        AND  tFacM.Recursive = 0 
                        AND  tFacM.Status = 2
                        AND dbo.tFacM.transferAccounting=0
                        AND (tfacm.[User] = @Uid OR @Uid = 0)
                GROUP BY tFacM.[Date] ,
						tPer.nvcFirstName + ' ' + tPer.nvcSurName + ' - ' + tblPub_Pos.nvcAccountNo ,
                        tblPub_Pos.AccountId
                HAVING   tFacM.[Date] BETWEEN @DateBefore AND @DateAfter 
                ORDER BY tFacM.[Date] ,
                        tblPub_Pos.AccountId
            END



GO




--==================  خلاصه گزارش صندوق  ثبت======================

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblAcc_ReceivedSummary]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblAcc_ReceivedSummary]
GO


CREATE TABLE [dbo].[tblAcc_ReceivedSummary](
	SanadNo [int]  NOT NULL,
	intRow INT NOT NULL ,
	[Uid] [int] NOT NULL,
	[nvcDate] [nvarchar](8) NOT NULL,
	[ShiftNo] INT NOT NULL ,
	[ReceivedType] TINYINT NOT NULL,
	[PosId] [int] NULL,
	[Price] [bigint] NOT NULL ,
	nvcDescription NVARCHAR(255) NOT NULL 
) ON [PRIMARY]

GO


ALTER TABLE [dbo].[tblAcc_ReceivedSummary] ADD 
	CONSTRAINT [PK_tblAcc_ReceivedSummary] PRIMARY KEY  CLUSTERED 
	(
		SanadNo , intRow 	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE UNIQUE INDEX [IX_tblAcc_ReceivedSummary] ON [dbo].[tblAcc_ReceivedSummary]([Uid], nvcDate , ShiftNo , intRow  ASC ) ON [PRIMARY]
GO

 CREATE UNIQUE INDEX [IX_tblAcc_ReceivedSummary_PosId] ON [dbo].[tblAcc_ReceivedSummary]([Uid], nvcDate , ShiftNo , ReceivedType , PosId ASC ) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblAcc_ReceivedSummary] ADD 
	CONSTRAINT [FK_tblAcc_ReceivedSummary_tUser] FOREIGN KEY 
	(
		Uid
	) REFERENCES [dbo].tUser (
		Uid
	)  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblAcc_ReceivedSummary] ADD 
	CONSTRAINT [FK_tblAcc_ReceivedSummary_tPos] FOREIGN KEY 
	(
		PosId
	) REFERENCES [dbo].tblPub_Pos (
		PosId
	)  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblAcc_ReceivedSummary] ADD 
	CONSTRAINT [FK_tblAcc_ReceivedSummary_tRecvType] FOREIGN KEY 
	(
		ReceivedType
	) REFERENCES [dbo].tRecvType (
		tintType
	)  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblAcc_ReceivedSummary] ADD 
	CONSTRAINT [FK_tblAcc_ReceivedSummary_tShift] FOREIGN KEY 
	(
		ShiftNo
	) REFERENCES [dbo].tShift (
		Code
	)  ON UPDATE CASCADE 
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Insert_tblAcc_ReceivedSummary') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Insert_tblAcc_ReceivedSummary
GO


CREATE  PROCEDURE [dbo].Insert_tblAcc_ReceivedSummary
    (
      @nvcDate  NVARCHAR(8) ,
      @ShiftNo INT ,
      @Uid INT ,
      @intRow INT ,
      @ReceivedType TINYINT ,
      @Price BIGINT ,
      @PosId INT ,
      @nvcDescription NVARCHAR(255) ,
      @SanadNo INT ,
      @Result INT  OUT  
    )
AS 

IF @PosId = 0 SET @PosId = NULL 
INSERT INTO dbo.tblAcc_ReceivedSummary
        ( SanadNo ,
          Uid ,
          nvcDate ,
          ShiftNo ,
          intRow ,
          ReceivedType ,
          PosId ,
          Price ,
          nvcDescription 
        )
VALUES  ( @SanadNo ,
		  @Uid , -- Uid - int
          @nvcDate , -- nvcDate - nvarchar(8)
          @ShiftNo ,
          @intRow ,
          @ReceivedType , -- ReceivedType - tinyint
          @PosId , -- PosId - int
          @Price  ,-- Price - bigint
          @nvcDescription
        )
IF @@ERROR <> 0 GOTO ErrHandler

SET @Result = 1
RETURN @Result

ErrHandler:
SET @Result = -1
RETURN @Result

GO





if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_Max_tblAcc_ReceivedSummary') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_Max_tblAcc_ReceivedSummary
GO


CREATE  PROCEDURE [dbo].Get_Max_tblAcc_ReceivedSummary
AS

SELECT   IsNull(MAX(SanadNo),0) + 1 AS SanadNo FROM tblAcc_ReceivedSummary



GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Get_tblAcc_ReceivedSummary') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_tblAcc_ReceivedSummary
GO


CREATE  PROCEDURE [dbo].Get_tblAcc_ReceivedSummary
    (
      @nvcDate  NVARCHAR(8) ,
      @ShiftNo INT ,
      @Uid INT    )
AS 


SELECT * FROM dbo.tblAcc_ReceivedSummary
WHERE nvcDate = @nvcDate AND ShiftNo =  @ShiftNo 
AND (Uid = @Uid OR @Uid = 0)

GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].Delete_tblAcc_ReceivedSummary') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Delete_tblAcc_ReceivedSummary
GO


CREATE  PROCEDURE [dbo].Delete_tblAcc_ReceivedSummary
      @SanadNo INT 

AS


DELETE FROM tblAcc_ReceivedSummary WHERE SanadNo = @SanadNo 

GO


--INSERT INTO dbo.tObjects
--        ( intObjectCode ,
--          ObjectId ,
--          ObjectName ,
--          objectLatinName ,
--          intObjectType ,
--          ObjectParent
--        )
--VALUES  ( 326 , -- intObjectCode - int
--          N'FullSummaryReportSaved' , -- ObjectId - nvarchar(50)
--          N'تغییر گزارش ثبتی کاربران' , -- ObjectName - nvarchar(50)
--          N'FullSummaryReportSaved' , -- objectLatinName - nvarchar(50)
--          1 , -- intObjectType - int
--          102  -- ObjectParent - int
--        )
        
--GO

--INSERT INTO dbo.tAccess_Object
--        ( intAccessLevel, intObjectCode )
--VALUES  ( 1 ,-- intAccessLevel - int
--          326  -- intObjectCode - int
--          )
          
--GO

--DELETE FROM tObjects WHERE intObjectCode = 326


--========================================================================






ALTER  Proc Get_Customer_Code    
@ActDeact int ,    
@Code Bigint 
    
as    

Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where MemberShipId = @Code and actdeact <> @ActDeact-- AND branch = @Branch  
AND vw_Get_Cust.Code <> -1


GO





ALTER  Proc Get_Customer_Name
@ActDeact int ,
@Name Nvarchar(50)     
as    


Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where  CHARINDEX ( @Name , [Name] ) > 0 and actdeact <> @ActDeact  -- AND Branch = @Branch
AND vw_Get_Cust.Code <> -1
Order By [Name]



GO




ALTER  Proc Get_Customer_Tel
@ActDeact int ,
@Tel Nvarchar(20)    
as    


Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
	where (CHARINDEX ( @Tel , [Tel1] ) > 0 Or CHARINDEX ( @Tel , [Tel2] ) > 0
        Or CHARINDEX ( @Tel , [Tel3] ) > 0 Or CHARINDEX ( @Tel , [Tel4] ) > 0
        Or CHARINDEX ( @Tel , [Mobile] ) > 0 ) 
        and actdeact <> @ActDeact 
AND vw_Get_Cust.Code <> -1 



GO





ALTER  Proc Get_Customer_Address
@ActDeact int ,
@Address Nvarchar(200) 
as    


Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where  CHARINDEX ( @Address , [Address] ) > 0 and actdeact <> @ActDeact
AND vw_Get_Cust.Code <> -1 --AND Branch = @Branch


GO


ALTER  Proc Get_Customer_Prefix  
@ActDeact int ,    
@Code INT  

as    


Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where Prefix = @Code and actdeact <> @ActDeact-- AND branch = dbo.[Get_Current_Branch]()  
AND vw_Get_Cust.Code <> -1 --AND Branch = @Branch


GO




