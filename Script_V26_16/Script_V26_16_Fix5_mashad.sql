

--Script_V26_16_Fix5
--93/03/30

توجه توجه :  
 --Remove Identity FROM tables:   را به صورت دستی از جداول ذیل بردارید Identity 
  tfacm - tcust - tper - tuser - ttable - tpartitions - tinventory - tacc_Cash - tAcc_Received		


--امکان گرفتن سود و زیان بازرگانی کالاها
--ارزش موجودی اولیه انبار
--ارزش نهایی انبار
--اصلاح تفضیلی های حسابداری
--برداشتن آیدنتیتی از جداول برای رپلیکیشن و انتقال دیتا


IF NOT EXISTS ( SELECT * FROM tblPub_Script2 WHERE Version = 26 AND  Script = 16 AND FixNumber = 5 )

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
          5
        )
GO


IF NOT EXISTS ( SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME='tFacD' AND COLUMN_NAME='FinalPrice' )

	ALTER TABLE dbo.tFacD
	ADD FinalPrice INT NOT NULL DEFAULT(0)

GO

IF NOT EXISTS ( SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME='tFacM' AND COLUMN_NAME='RefrenceHavale' )
	ALTER TABLE dbo.tFacM
	ADD RefrenceHavale INT NULL

GO

IF NOT EXISTS ( SELECT * FROM tblAcc_Moein WHERE KolId = 17 AND MoeinId = 1702 )

INSERT INTO dbo.tblAcc_Moein
        ( KolId ,
          MoeinId ,
          MoeinName ,
          Kind ,
          Active
        )
VALUES  ( 17 , -- KolId - int
          1702 , -- MoeinId - int
          N'موجودي اوليه' , -- MoeinName - nvarchar(50)
          1 , -- Kind - tinyint
          1  -- Active - bit
        )
        
GO

IF NOT EXISTS ( SELECT * FROM tObjects WHERE intObjectCode = 344 )

INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 344 , -- intObjectCode - int
          N'frmBenefit' , -- ObjectId - nvarchar(50)
          N'سود و زیان کالاها' , -- ObjectName - nvarchar(50)
          N'frmBenefit' , -- objectLatinName - nvarchar(50)
          1 , -- intObjectType - int
          102  -- ObjectParent - int
        )
        
GO


ALTER   PROCEDURE [dbo].[Get_PC_Stations]

 AS
Select * from tStations Where (StationType  &  2  = 2 ) and IsActive =1 And Branch =  dbo.Get_Current_Branch()

GO




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

ALTER  PROCEDURE dbo.UpdatetGood
(
	@intLanguage	INT,
	@Goodname	NVARCHAR(50),
	@GoodNamePrn	NVARCHAR(50),
	@SellPrice	FLOAT,
	@BuyPrice	FLOAT,
	@Unit		INT,
	@GoodType	INT,
	@Barcode	NVARCHAR(50),
	@Code		INT,
	@Weight	Float,
	@NumberOfUnit 	INT,
	@SellPrice2 Float,
	@SellPrice3 Float ,
	@MainType Bit ,
	@Supplier Int ,
	@Level1 Int ,
	@Level2 Int ,
	@SellPrice4 Float,
	@SellPrice5 Float,
	@SellPrice6 Float,
	@CategoryShow INT ,
	@PicturePath NVARCHAR(100) ,
	@nvcDescription NVARCHAR(100) ,
	@Picture IMAGE ,
	@GoodNamePrn2	NVARCHAR(100),
	@GoodNamePrn3	NVARCHAR(100),
	@RealNewCode INT ,
	@Result Int Out
)

AS

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tUsePercent_tGood1]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
		ALTER TABLE [dbo].[tUsePercent] DROP CONSTRAINT [FK_tUsePercent_tGood1]



Declare  @NewCode INT
SET @NewCode = @RealNewCode
IF @RealNewCode = 0 
BEGIN 
	Set @NewCode = @Code
	Declare @Level2Code	INT
	Set @Level2Code = (Select Level2 From tGood Where Code = @Code)

	Begin Tran

	If @Level2 <>  @Level2Code
	Begin
	--	Set @NewCode =  (SELECT  ISNULL(MAX(RIGHT(RTRIM(LTRIM(STR(code))),LEN(RTRIM(LTRIM(STR(Code))))-4)),0) +1   
		Set @NewCode =  (SELECT  ISNULL(MAX(code),0) + 1   
		FROM dbo.tgood 
		WHERE Level2 = @Level2 )
		If Len(@NewCode) = 1 

		Set @NewCode = (@Level2 * 10000) + @NewCode 

	End

END 

IF @intLanguage = 0 
Begin		
		UPDATE dbo.tGood

		SET [Name]    = dbo.Get_ArabicToFarsiString(@GoodName) ,
		    NamePrn   = dbo.Get_ArabicToFarsiString(@GoodNamePrn) ,
		    SellPrice = @SellPrice ,
		    BuyPrice  = @BuyPrice ,
		    Unit      = @Unit ,
		    GoodType  = @GoodType ,
		    Barcode = @Barcode,
	                 Weight = @Weight,
		    NumberOfUnit=@NumberOfUnit,
		    SellPrice2 = @SellPrice2,
		    SellPrice3 = @SellPrice3 ,	    	
		    SellPrice4 = @SellPrice4 ,	    	
		    SellPrice5 = @SellPrice5 ,	    	
		    SellPrice6 = @SellPrice6 ,	    	
		    MainType = @MainType  ,
		    ProductCompany = @Supplier ,
		   Level1 = @Level1 ,
		   Level2 = @Level2 ,
		 Code = @NewCode ,
		 CategoryShow = @CategoryShow ,
		 PicturePath = @PicturePath ,
		 nvcDescription = @nvcDescription ,
	    GoodNamePrn2   = dbo.Get_ArabicToFarsiString(@GoodNamePrn2) ,
	    GoodNamePrn3   = dbo.Get_ArabicToFarsiString(@GoodNamePrn3) 
	WHERE Code = @Code		
        IF @@ERROR <>0
	        GoTo EventHandler

End
ELSE IF @intLanguage = 1 
Begin
		UPDATE dbo.tGood

		SET LatinName     = @GoodName ,
		    LatinNamePrn  = @GoodNamePrn ,
		    SellPrice     = @SellPrice ,
		    BuyPrice      = @BuyPrice ,
		    Unit          = @Unit ,
		    GoodType      = @GoodType,
		    Barcode = @Barcode,
		    Weight = @Weight,
		    NumberOfUnit=@NumberOfUnit,
		    SellPrice2 = @SellPrice2,
		    SellPrice3 = @SellPrice3 ,
		    SellPrice4 = @SellPrice4 ,	    	
		    SellPrice5 = @SellPrice5 ,	    	
		    SellPrice6 = @SellPrice6 ,	    	
		    MainType = @MainType ,
	 	    ProductCompany = @Supplier ,
		   Level1 = @Level1 ,
		   Level2 = @Level2 ,
		 Code = @NewCode ,
		 CategoryShow = @CategoryShow ,
		 PicturePath = @PicturePath ,
		 nvcDescription = @nvcDescription ,
	    GoodNamePrn2   = @GoodNamePrn2 ,
	    GoodNamePrn3   = @GoodNamePrn3
		WHERE Code = @Code

        IF @@ERROR <>0
	        GoTo EventHandler

End
Set @Result = 1
	update  [dbo].[tGood]  set [name] = latinname where ([Name] is null or [Name] = ''  ) And Code = @NewCode

	update  [dbo].[tGood] set latinname = [name] where ([latinName] is null or latinname = '') And Code = @NewCode

	update  [dbo].[tGood] set [nameprn]=[latinnameprn] where ([Nameprn] is null or [Nameprn] = '') And Code = @NewCode 

	update  [dbo].[tGood] set [GoodNamePrn2]=[nameprn] where ([GoodNamePrn2] is null or [GoodNamePrn2] = '' ) And Code = @NewCode

	update  [dbo].[tGood] set [GoodNamePrn3]=[nameprn] where ([GoodNamePrn3] is null or [GoodNamePrn3] = '' ) And Code = @NewCode

	update  [dbo].[tGood] set [latinnameprn] = [nameprn] where ([latinNameprn] is null or latinnameprn = '') And Code = @NewCode

	UPDATE dbo.tGood SET Picture = @Picture WHERE Code = @Code

	ALTER TABLE [dbo].[tUsePercent]  WITH CHECK ADD  CONSTRAINT [FK_tUsePercent_tGood1] FOREIGN KEY([GoodFirstCode])
		REFERENCES [dbo].[tGood] ([Code])

	ALTER TABLE [dbo].[tUsePercent] CHECK CONSTRAINT [FK_tUsePercent_tGood1]

	update  [dbo].[tGood] SET nvcDate = dbo.shamsi(GETDATE()) WHERE Code = @Code

COMMIT TRANSACTION


Return @Result


EventHandler:
    ROLLBACK TRAN
    Set @Result = 0
    RETURN @Result



GO



ALTER  PROCEDURE [dbo].[Get_Previous_Factor_Detail] (@intLanguage  int , @Code INT , @Branch INT ) AS
SELECT     dbo.tFacD2.*, case @intLanguage when  0 then dbo.tGood.Name 
when 1 then dbo.tGood.LatinName end AS Name
FROM         dbo.tFacD2 INNER JOIN
                      dbo.tGood ON dbo.tFacD2.GoodCode = dbo.tGood.Code
where tFacD2.Code = @Code and Branch =  @Branch 

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
--exec Get_EditedFactors 0, N'85/12/28', N'85/12/28'

ALTER  VIEW vw_EditedFactors
AS
SELECT     dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tFacM.Status, dbo.tFacM.Owner, dbo.tFacM.Customer, dbo.tFacM.DiscountTotal, 
                      dbo.tFacM.CarryFeeTotal, dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.InCharge, dbo.tFacM.FacPayment, dbo.tFacM.OrderType, 
                      dbo.tFacM.ServePlace, dbo.tFacM.StationID, dbo.tFacM.ServiceTotal, dbo.tFacM.PackingTotal, dbo.tFacM.BascoleNo, dbo.tFacM.TableNo, 
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description, dbo.tShift.LatinDescription , dbo.tFacM.Branch-- , dbo.tRepFacEditM.Sumprice as Sumpriceold
FROM         dbo.tFacM INNER JOIN
                      dbo.tRepFacEditM ON dbo.tFacM.intSerialNo = dbo.tRepFacEditM.intSerialNo and dbo.tFacM.Branch = dbo.tRepFacEditM.Branch INNER JOIN
                      dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
                      dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  --and  dbo.tFacM.Branch = dbo.tShift.Branch
WHERE     (dbo.tFacM.intSerialNo IN
                          (SELECT     dbo.tRepFacEditM.intSerialNo
                             FROM         dbo.tRepFacEditM )) --Where Branch = dbo.Get_Current_Branch()




GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER  VIEW dbo.vw_EditedFactors1
AS
SELECT     dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tStatusType.NvcDescription as statusDescription ,dbo.tStatusType.NvcLatinDescription as statusLatinDescription, dbo.tFacM.Owner, dbo.tFacM.Customer, dbo.tFacM.DiscountTotal, 
                      dbo.tFacM.CarryFeeTotal, dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.InCharge, dbo.tFacM.FacPayment, dbo.tFacM.OrderType, 
                      dbo.tServePlace.Description as tServePlaceDescription, dbo.tServePlace.LatinDescription as tServePlaceLatinDescription, dbo.tFacM.StationID, dbo.tFacM.ServiceTotal, dbo.tFacM.PackingTotal, dbo.tFacM.BascoleNo, dbo.tFacM.TableNo, 
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description, dbo.tShift.LatinDescription
FROM         dbo.tFacM INNER JOIN
                      dbo.tRepFacEditM ON dbo.tFacM.intSerialNo = dbo.tRepFacEditM.intSerialNo and dbo.tFacM.Branch = dbo.tRepFacEditM.Branch INNER JOIN
                      dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
                      dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  --and  dbo.tFacM.Branch = dbo.tShift.Branch 
                      inner JOIN   dbo.tstatustype on dbo.tfacm.status=tStatusType.intStatusNo inner join
					dbo.tServePlace ON dbo.tFacM.ServePlace = dbo.tServePlace.intServePlace
			
			WHERE     (dbo.tFacM.intSerialNo IN
                          (SELECT     dbo.tRepFacEditM.intSerialNo
                             FROM         dbo.tRepFacEditM Where Branch = dbo.Get_Current_Branch()))



GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  VIEW dbo.vw_EditedFactors1_Less
AS
SELECT     dbo.tFacM.intSerialNo, dbo.tFacM.[No], dbo.tStatusType.NvcDescription as statusDescription ,dbo.tStatusType.NvcLatinDescription as statusLatinDescription, dbo.tFacM.Owner, dbo.tFacM.Customer, dbo.tFacM.DiscountTotal, 
                      dbo.tFacM.CarryFeeTotal, dbo.tFacM.SumPrice, dbo.tFacM.Recursive, dbo.tFacM.InCharge, dbo.tFacM.FacPayment, dbo.tFacM.OrderType, 
                      dbo.tServePlace.Description as tServePlaceDescription, dbo.tServePlace.LatinDescription as tServePlaceLatinDescription, dbo.tFacM.StationID, dbo.tFacM.ServiceTotal, dbo.tFacM.PackingTotal, dbo.tFacM.BascoleNo, dbo.tFacM.TableNo, 
                      dbo.tFacM.[Date], dbo.tFacM.[Time], dbo.tFacM.[User], dbo.tFacM.RegDate, dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, 
                      dbo.tFacM.ShiftNo, dbo.tShift.Description, dbo.tShift.LatinDescription 
FROM         dbo.tFacM INNER JOIN
                      dbo.tRepFacEditM ON dbo.tFacM.intSerialNo = dbo.tRepFacEditM.intSerialNo and dbo.tFacM.Branch = dbo.tRepFacEditM.Branch INNER JOIN
                      dbo.tUser ON dbo.tFacM.[User] = dbo.tUser.UID and dbo.tFacM.Branch = dbo.tUser.Branch  INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
                      dbo.tShift ON dbo.tFacM.ShiftNo = dbo.tShift.Code  --and  dbo.tFacM.Branch = dbo.tShift.Branch 
                      inner JOIN  dbo.tstatustype on dbo.tfacm.status=tStatusType.intStatusNo inner join
		      dbo.tServePlace ON dbo.tFacM.ServePlace = dbo.tServePlace.intServePlace 
		      
			WHERE     (dbo.tFacM.intSerialNo IN
                          (SELECT     dbo.tRepFacEditM.intSerialNo
                             FROM         dbo.tRepFacEditM Where dbo.tRepFacEditM.Branch=dbo.tFacM.Branch)) And  dbo.tFacM.SumPrice < dbo.tRepFacEditM.SumPrice




GO


ALTER  VIEW dbo.vw_EditedFactors2
AS
SELECT     dbo.tRepFacEditM.code,dbo.tRepFacEditM.intSerialNo, dbo.tRepFacEditM.[No], dbo.tStatusType.NvcDescription as statusDescription ,dbo.tStatusType.NvcLatinDescription as statusLatinDescription, dbo.tRepFacEditM.Owner, dbo.tRepFacEditM.Customer, dbo.tRepFacEditM.DiscountTotal, 
                      dbo.tRepFacEditM.CarryFeeTotal, dbo.tRepFacEditM.SumPrice, dbo.tRepFacEditM.Recursive, dbo.tRepFacEditM.InCharge, dbo.tRepFacEditM.FacPayment, dbo.tRepFacEditM.OrderType, 
                                            dbo.tServePlace.Description as tServePlaceDescription, dbo.tServePlace.LatinDescription as tServePlaceLatinDescription, dbo.tRepFacEditM.StationID, dbo.tRepFacEditM.ServiceTotal, dbo.tRepFacEditM.PackingTotal, dbo.tRepFacEditM.BascoleNo, dbo.tRepFacEditM.TableNo, 
                      dbo.tRepFacEditM.[Date],(SELECT dbo.tFacM.[Date] FROM dbo.tFacM 
					WHERE dbo.tRepFacEditM.intSerialNo=dbo.tFacM.intSerialNo
					AND dbo.tRepFacEditM.Branch=dbo.tFacM.Branch
					) AS Date1, dbo.tRepFacEditM.[Time], dbo.tRepFacEditM.[User], dbo.tRepFacEditM.RegDate, dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, 
                      dbo.tRepFacEditM.ShiftNo, dbo.tShift.Description, dbo.tShift.LatinDescription
FROM         dbo.tRepFacEditM INNER JOIN
                      dbo.tUser ON dbo.tRepFacEditM.[User] = dbo.tUser.UID and dbo.tRepFacEditM.Branch = dbo.tUser.Branch  INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
			
                      dbo.tShift ON dbo.tRepFacEditM.ShiftNo = dbo.tShift.Code  --and  dbo.tRepFacEditM.Branch = dbo.tShift.Branch
			inner join
		      dbo.tstatustype on dbo.trepfaceditm.status=tStatusType.intStatusNo
			inner join
		      dbo.tServePlace ON dbo.trepfaceditm.ServePlace = dbo.tServePlace.intServePlace

GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  VIEW dbo.vw_EditedFactors2_Less
AS
SELECT     dbo.tRepFacEditM.code,dbo.tRepFacEditM.intSerialNo, dbo.tRepFacEditM.[No], dbo.tStatusType.NvcDescription as statusDescription ,dbo.tStatusType.NvcLatinDescription as statusLatinDescription, dbo.tRepFacEditM.Owner, dbo.tRepFacEditM.Customer, dbo.tRepFacEditM.DiscountTotal, 
                      dbo.tRepFacEditM.CarryFeeTotal, dbo.tRepFacEditM.SumPrice, dbo.tRepFacEditM.Recursive, dbo.tRepFacEditM.InCharge, dbo.tRepFacEditM.FacPayment, dbo.tRepFacEditM.OrderType, 
                                            dbo.tServePlace.Description as tServePlaceDescription, dbo.tServePlace.LatinDescription as tServePlaceLatinDescription, dbo.tRepFacEditM.StationID, dbo.tRepFacEditM.ServiceTotal, dbo.tRepFacEditM.PackingTotal, dbo.tRepFacEditM.BascoleNo, dbo.tRepFacEditM.TableNo, 
                      dbo.tRepFacEditM.[Date],(SELECT dbo.tFacM.[Date] FROM dbo.tFacM 
					WHERE dbo.tRepFacEditM.intSerialNo=dbo.tFacM.intSerialNo
					AND dbo.tRepFacEditM.Branch=dbo.tFacM.Branch 
					) AS Date1, dbo.tRepFacEditM.[Time], dbo.tRepFacEditM.[User], dbo.tRepFacEditM.RegDate, dbo.tPer.nvcFirstName, dbo.tPer.nvcSurName, 
                      dbo.tRepFacEditM.ShiftNo, dbo.tShift.Description, dbo.tShift.LatinDescription 
FROM         dbo.tRepFacEditM INNER JOIN
                      dbo.tUser ON dbo.tRepFacEditM.[User] = dbo.tUser.UID and dbo.tRepFacEditM.Branch = dbo.tUser.Branch  INNER JOIN
                      dbo.tPer ON dbo.tUser.pPno = dbo.tPer.pPno and dbo.tUser.Branch = dbo.tPer.Branch INNER JOIN
			
                      dbo.tShift ON dbo.tRepFacEditM.ShiftNo = dbo.tShift.Code  --and  dbo.tRepFacEditM.Branch = dbo.tShift.Branch
			inner join
		      dbo.tstatustype on dbo.trepfaceditm.status=tStatusType.intStatusNo
			inner join
		      dbo.tServePlace ON dbo.trepfaceditm.ServePlace = dbo.tServePlace.intServePlace
		 	INNER JOIN
                      dbo.tFacM ON dbo.tFacM.intSerialNo = dbo.tRepFacEditM.intSerialNo and dbo.tFacM.Branch = dbo.tRepFacEditM.Branch And dbo.tFacM.SumPrice < dbo.tRepFacEditM.SumPrice
			


 
GO


IF NOT EXISTS ( SELECT * FROM tAccess_Object WHERE intAccessLevel = 1 AND  intObjectCode = 344 )

INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          344  -- intObjectCode - int
          )
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_FacDFinalPrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_FacDFinalPrice]
GO

CREATE  PROCEDURE [dbo].[Update_FacDFinalPrice]
    (
	@InventoryNo INT ,
	@AccountYear SMALLINT ,
	@GoodCode INT ,
	@Flag INT ,
	@BeforeDate NVARCHAR(8),
	@AfterDate NVARCHAR(8),
	@NumberOfRecords INT OUT 
    )
AS 

-- DECLARE @GoodCode INT
-- SELECT  @GoodCode = 4
-- 
-- DECLARE @InventoryNo INT 
-- DECLARE @Branch INT 
-- DECLARE @AccountYear SMALLINT
-- -- 
-- SELECT  @InventoryNo = 100
-- SELECT  @Branch = 1
-- SELECT  @AccountYear = 1389
-- 

--PRINT '###########*'

DECLARE @BuyPrice INT 
IF @Flag = 0 
    SELECT @NumberOfRecords = ISNULL(COUNT(GoodCode), 0)  --, [TIME] 
    FROM [dbo].[tFacM]
    INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
    WHERE [Status] IN( 2,5) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
	AND (GoodCode = @GoodCode OR @GoodCode = 0) 
	AND tFacM.AccountYear = @AccountYear 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate

ELSE 
BEGIN

    SET  @NumberOfRecords = 0			
    DECLARE  GoodsList CURSOR	 
    FOR 

 SELECT DISTINCT T2.GoodCode , dbo.tGood.BuyPrice FROM 
( SELECT DISTINCT  T1.GoodCode FROM 
(   SELECT  DISTINCT      [GoodCode]

    FROM    dbo.tFacM
    INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		AND [dbo].[tFacM].Branch = dbo.tFacD.Branch
    WHERE   dbo.tFacM.Status IN ( 1, 2 ,3, 4,5 , 6, 7 )
    AND tFacM.AccountYear = @AccountYear
    AND tFacM.Recursive = 0
    AND [dbo].[tFacD].intInventoryNo = @InventoryNo
    AND (GoodCode = @GoodCode OR @GoodCode = 0)  -- For One Good(FrmGoodTurnOver) or AllGood(FrmFinalPrice) 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate
UNION all
    SELECT  [GoodCode]
	FROM      tInventory_Good
	WHERE     (tInventory_Good.GoodCode = @GoodCode OR @GoodCode = 0)
	AND dbo.tInventory_Good.AccountYear = @AccountYear
	AND [InventoryNo] = @InventoryNo AND tInventory_Good.FirstMojodi <> 0
) T1
GROUP BY GoodCode 
)T2
INNER JOIN dbo.tGood ON T2.GoodCode = dbo.tGood.Code

	
    OPEN GoodsList
    FETCH FROM GoodsList INTO @GoodCode , @BuyPrice

    WHILE @@FETCH_STATUS = 0 
        BEGIN
            DECLARE @intSerialNo INT
	    DECLARE @Branch INT 
            DECLARE @fDate NVARCHAR(8)
            --DECLARE @fTime NVARCHAR(8)
            DECLARE Havale CURSOR 
            FOR 
            SELECT DISTINCT tFacM.Branch,tFacM.intSerialNo,[Date] --, GoodCode  
            FROM [dbo].[tFacM]
            INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
            WHERE [Status] IN( 2,5) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
            AND GoodCode = @GoodCode --AND (GoodCode = @GoodCode OR @GoodCode = 0) 
            AND tFacM.AccountYear = @AccountYear 
            AND dbo.tFacM.Date <=@AfterDate-- N'88/06/31'  --*****************
            AND dbo.tFacM.Date>=@BeforeDate
            ORDER BY [Date] ASC , [dbo].[tFacM].intSerialNo ASC  

            OPEN Havale
	
            FETCH  FROM Havale INTO @Branch ,@intSerialNo,@fDate --,@GoodCode 
	
            WHILE @@FETCH_STATUS = 0 
                BEGIN
                    DECLARE @priceTamam INT ;
                    DECLARE @Mablagh BIGINT ;
                    DECLARE @Tedad INT ;
                    --SELECT @Tedad = ISNULL(FirstMojodi , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    --SELECT @Mablagh = ISNULL(FirstPrice , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    SELECT  @Mablagh = SUM(T.FirstMojodi * T.FirstPrice) + SUM(T.Amount * T.Flag * T.FeeUnit) ,                                          
                            @Tedad = SUM(T.FirstMojodi) + Sum(T.Amount * T.Flag)
                    FROM    (
                              SELECT    dbo.tInventory_Good.FirstMojodi ,
                                        dbo.tInventory_Good.FirstPrice ,
                                        tInventory_Good.GoodCode ,
                                        0 AS Amount ,
                                        0 AS Flag ,
                                        0 AS FeeUnit 
                                        FROM      tInventory_Good
                              WHERE     tInventory_Good.GoodCode = @GoodCode
                                        AND dbo.tInventory_Good.AccountYear = @AccountYear
                                        AND [InventoryNo] = @InventoryNo
                              UNION ALL
                              SELECT    0 AS FirstMojodi ,
                              		0 AS FirstPrice ,
                              		Goodcode ,
                                        Amount ,
                                        Flag ,
                                        FeeUnit 
                              FROM      dbo.[tFacM]
                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
                                        INNER JOIN dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo
                              WHERE     dbo.tFacM.[Date] <= @fDate 
                              		--( dbo.tFacM.[Date] + ' ' + dbo.tFacM.[Time] ) <= ( @fDate + ' ' + @fTime )
                                        AND tFacM.Status IN ( 1, 3, 4 ) --, 6, 7
                                        AND (dbo.tFacM.intSerialNo < @intSerialNo OR status = 1 OR status = 4)
                                        AND dbo.tFacM.Branch = @Branch
                                        AND dbo.tFacM.AccountYear = @AccountYear
                                        AND tFacD.GoodCode = @GoodCode
                                        AND dbo.tFacM.Recursive = 0
                                        AND dbo.tfacD.[intInventoryNo] = @InventoryNo
					AND dbo.tFacM.Date <=@AfterDate
					AND dbo.tFacM.Date>=@BeforeDate
                            ) T
                    GROUP BY GoodCode 
                    IF @Tedad <= 0 
                    	SET @priceTamam = @BuyPrice
             	    ELSE
             	        SET @priceTamam = CAST((@Mablagh/@Tedad) AS INT)
             	        
--PRINT @GoodCode 
--PRINT   @priceTamam          	           	
--PRINT @NumberOfRecords
 --                   IF @priceTamam >= 0 
                        UPDATE  dbo.tFacD
                        SET     FinalPrice = @priceTamam
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
--                     IF @priceTamam < 0 			--Negative Price set to Zero
--                         UPDATE  dbo.tFacD
--                         SET     FinalPrice = 0
--                         WHERE   dbo.tFacD.intSerialNo = @intSerialNo
--                                 AND dbo.tFacD.Branch = @Branch
--                                 AND dbo.tFacD.GoodCode = @GoodCode
			
                    SET @NumberOfRecords = @NumberOfRecords + 1
                    FETCH NEXT FROM Havale INTO @Branch ,@intSerialNo,@fDate --, @GoodCode 
	
                END
	
            CLOSE Havale
            DEALLOCATE Havale
           
	FETCH NEXT  FROM GoodsList INTO @GoodCode , @BuyPrice

        END
    CLOSE GoodsList
    DEALLOCATE GoodsList
--=====================================================

	END
	RETURN @NumberOfRecords


IF @@ERROR <> 0
    AND @@TRANCOUNT > 0 
    ROLLBACK TRANSACTION ;


GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_FirstPriceByBuyPrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_FirstPriceByBuyPrice
GO


CREATE  PROC Update_FirstPriceByBuyPrice
	(
	@AccountYear INT,
	@Flag BIT,
	@InventoryNO INT
	)
AS
BEGIN
IF @Flag = 0
	BEGIN
		UPDATE dbo.tInventory_Good
		SET dbo.tInventory_Good.FirstPrice=dbo.tGood.BuyPrice
		FROM dbo.tGood 
			JOIN dbo.tInventory_Good ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
		WHERE dbo.tInventory_Good.FirstPrice = 0--@Flag
			AND dbo.tInventory_Good.AccountYear = @AccountYear
			AND dbo.tInventory_Good.InventoryNo = @InventoryNO--ISNULL(@InventoryNO,dbo.tInventory_Good.InventoryNo)
			AND dbo.tGood.BuyPrice <> 0
				
	END
ELSE 
	BEGIN
		UPDATE dbo.tInventory_Good
		SET dbo.tInventory_Good.FirstPrice  =dbo.tGood.BuyPrice
		FROM dbo.tGood 
			JOIN dbo.tInventory_Good ON dbo.tGood.Code = dbo.tInventory_Good.GoodCode
		WHERE  dbo.tInventory_Good.AccountYear = @AccountYear
			AND dbo.tInventory_Good.InventoryNo = @InventoryNO--ISNULL(@InventoryNO,dbo.tInventory_Good.InventoryNo)
			AND dbo.tGood.BuyPrice <> 0

	END
END



GO




if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_HavalehResid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Update_HavalehResid]
GO

CREATE PROCEDURE [dbo].[Update_HavalehResid]
    (
	@InventoryNo INT ,
	@AccountYear SMALLINT ,
	@GoodCode INT ,
	@Flag INT ,
	@BeforeDate NVARCHAR(8),
	@AfterDate NVARCHAR(8),
	@NumberOfRecords INT OUT 
    )
AS 
DECLARE @BuyPrice INT  
DECLARE @GoodCode1 INT 
SET @GoodCode1 = @GoodCode
-- DECLARE @GoodCode INT
-- SELECT  @GoodCode = 4
-- 
-- DECLARE @InventoryNo INT 
-- DECLARE @Branch INT 
-- DECLARE @AccountYear SMALLINT
-- -- 
-- SELECT  @InventoryNo = 100
-- SELECT  @Branch = 1
-- SELECT  @AccountYear = 1389
-- 

IF @Flag = 0 
    SELECT @NumberOfRecords = ISNULL(COUNT(GoodCode), 0)  --, [TIME] 
    FROM [dbo].[tFacM]
    INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
    WHERE [Status] IN( 6,7) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
	AND (GoodCode = @GoodCode OR @GoodCode = 0) 
	AND tFacM.AccountYear = @AccountYear 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate

ELSE 
BEGIN

    SET  @NumberOfRecords = 0			
    DECLARE  GoodsList CURSOR	 
    FOR 

 SELECT DISTINCT T2.GoodCode , dbo.tGood.BuyPrice FROM 
( SELECT DISTINCT T1.GoodCode  FROM 
(   SELECT  DISTINCT      [GoodCode]

    FROM    dbo.tFacM
    INNER JOIN dbo.tFacD ON dbo.tFacM.intSerialNo = dbo.tFacD.intSerialNo
		AND [dbo].[tFacM].Branch = dbo.tFacD.Branch
    WHERE   dbo.tFacM.Status IN ( 1, 2 ,3, 4,5 , 6, 7 )
    AND tFacM.AccountYear = @AccountYear
    AND tFacM.Recursive = 0
    AND [dbo].[tFacD].intInventoryNo = @InventoryNo
    AND (GoodCode = @GoodCode OR @GoodCode = 0)  -- For One Good(FrmGoodTurnOver) or AllGood(FrmFinalPrice) 
	AND dbo.tFacM.Date<=@AfterDate
	AND dbo.tFacM.Date>=@BeforeDate
UNION all
    SELECT  [GoodCode]
	FROM      tInventory_Good
	WHERE     (tInventory_Good.GoodCode = @GoodCode OR @GoodCode = 0)
	AND dbo.tInventory_Good.AccountYear = @AccountYear
	AND [InventoryNo] = @InventoryNo AND tInventory_Good.FirstMojodi <> 0
) T1
GROUP BY GoodCode 
)T2
INNER JOIN dbo.tGood ON T2.GoodCode = dbo.tGood.Code


	
    OPEN GoodsList
    FETCH FROM GoodsList INTO @GoodCode , @BuyPrice

    WHILE @@FETCH_STATUS = 0 
        BEGIN
            DECLARE  @intSerialNo INT
            DECLARE @Branch INT 
	        DECLARE @fDate NVARCHAR(8)
            --DECLARE @fTime NVARCHAR(8)
            DECLARE Havale CURSOR 
            FOR 
            SELECT DISTINCT tFacM.Branch,tFacM.intSerialNo,[Date] --, GoodCode  
            FROM [dbo].[tFacM]
            INNER JOIN [dbo].[tFacD] ON dbo.tFacM.intSerialNo = [dbo].[tFacD].intSerialNo AND [dbo].[tFacM].Branch = [dbo].[tFacD].Branch
            WHERE [Status] IN( 6,7) AND [dbo].[tFacD].intInventoryNo = @InVentoryNo 
            AND GoodCode = @GoodCode --AND (GoodCode = @GoodCode OR @GoodCode = 0) 
            AND tFacM.AccountYear = @AccountYear 
            AND dbo.tFacM.Date <=@AfterDate-- N'88/06/31'  --*****************
            AND dbo.tFacM.Date>=@BeforeDate
            ORDER BY [Date] ASC , [dbo].[tFacM].intSerialNo ASC  

            OPEN Havale
	
            FETCH  FROM Havale INTO @Branch ,@intSerialNo,@fDate --,@GoodCode 
	
            WHILE @@FETCH_STATUS = 0 
                BEGIN
                    DECLARE @priceTamam INT ;
                    DECLARE @Mablagh BIGINT ;
                    DECLARE @Tedad INT ;
                    --SELECT @Tedad = ISNULL(FirstMojodi , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    --SELECT @Mablagh = ISNULL(FirstPrice , 0) FROM dbo.tInventory_Good WHERE Branch = @Branch
                    --          		AND InventoryNo = @InventoryNo AND AccountYear = @AccountYear AND GoodCode = @GoodCode 
                    SELECT  @Mablagh = SUM(T.FirstMojodi * T.FirstPrice) + SUM(T.Amount * T.Flag * T.FeeUnit) ,                                          
                            @Tedad = SUM(T.FirstMojodi) + Sum(T.Amount * T.Flag)
                    FROM    (
                              SELECT    dbo.tInventory_Good.FirstMojodi ,
                                        dbo.tInventory_Good.FirstPrice ,
                                        tInventory_Good.GoodCode ,
                                        0 AS Amount ,
                                        0 AS Flag ,
                                        0 AS FeeUnit 
                                        FROM      tInventory_Good
                              WHERE     tInventory_Good.GoodCode = @GoodCode
                                        AND dbo.tInventory_Good.AccountYear = @AccountYear
                                        AND [InventoryNo] = @InventoryNo
                              UNION ALL
                              SELECT    0 AS FirstMojodi ,
                              		0 AS FirstPrice ,
                              		Goodcode ,
                                        Amount ,
                                        Flag ,
                                        FeeUnit 
                              FROM      dbo.[tFacM]
                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
                                        INNER JOIN dbo.tStatusType ON dbo.tFacM.Status = dbo.tStatusType.intStatusNo
                              WHERE     dbo.tFacM.[Date] <= @fDate 
                              		--( dbo.tFacM.[Date] + ' ' + dbo.tFacM.[Time] ) <= ( @fDate + ' ' + @fTime )
                                        AND tFacM.Status IN ( 1, 3, 4 ) -- , 6, 7  براي موجودي منفي
                                        AND (dbo.tFacM.intSerialNo < @intSerialNo OR Status = 1 OR Status = 4)
                                        AND dbo.tFacM.Branch = @Branch
                                        AND dbo.tFacM.AccountYear = @AccountYear
                                        AND tFacD.GoodCode = @GoodCode
                                        AND dbo.tFacM.Recursive = 0
                                        AND [intInventoryNo] = @InventoryNo
										AND dbo.tFacM.Date <=@AfterDate
										AND dbo.tFacM.Date>=@BeforeDate
                            ) T
                    GROUP BY GoodCode 
                    IF @Tedad <= 0 
                    	SET @priceTamam = @BuyPrice
             	    ELSE
             	        SET @priceTamam = CAST((@Mablagh/@Tedad) AS INT)
             	        
                    DECLARE @Status1 INT 
                    SET @Status1 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1)         
                    DECLARE @Status2 INT 
                    SET @Status2 = (SELECT Status FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo)         
                    DECLARE @HavaleNo INT 
                    SET @HavaleNO = ISNULL((SELECT RefrenceHavale FROM dbo.tFacM WHERE Branch = @Branch 
                                AND intSerialNo = @intSerialNo - 1 AND Status = 6) , 0)         
--                    PRINT @GoodCode
--                     PRINT @Mablagh
--                     PRINT @Tedad
--                     PRINT @priceTamam
--                    PRINT @intSerialNo
--		    PRINT @NumberOfRecords + 1
                    --PRINT @HavaleNO
                    IF @priceTamam >= 0 
                        UPDATE  dbo.tFacD
                        SET     FeeUnit = @priceTamam
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
                                AND intSerialNo <> @HavaleNo -- we don,t need update Resid  from Havale
                    IF @priceTamam < 0 			--Negative Price set to Zero
                        UPDATE  dbo.tFacD
                        SET     FeeUnit = 0
                        WHERE   dbo.tFacD.intSerialNo = @intSerialNo
                                AND dbo.tFacD.Branch = @Branch
                                AND dbo.tFacD.GoodCode = @GoodCode
                                AND intSerialNo <> @HavaleNo -- we don,t need update Resid  from Havale

--Update Resid From Havale With Havale Fee
			IF @intSerialNo = @HavaleNo
				UPDATE dbo.tFacD
				SET FeeUnit = X.feeUnit		
	                        FROM (SELECT feeUnit FROM  dbo.[tFacM]
	                                        INNER JOIN dbo.tFacD ON dbo.tFacD.Branch = dbo.tFacM.Branch
	                                                              AND dbo.tFacM.intSerialNo = tFacD.intSerialNo
	                                        WHERE   tFacM.Status = 6
	                                        AND dbo.tFacM.intSerialNo = @intSerialNo -1
	                                        AND dbo.tFacM.Branch = @Branch
	                                        AND dbo.tFacM.AccountYear = @AccountYear
	                                        AND dbo.tFacM.Recursive = 0
	                                        AND dbo.tFacD.GoodCode = @GoodCode  
	                                       -- AND [intInventoryNo] = @InventoryNo  --No Inventory Needed because is resid from other inventory
	                                )X  
	                  	WHERE  dbo.tFacD.intSerialNo = @intSerialNo
	                                AND dbo.tFacD.Branch = @Branch
	                                AND dbo.tFacD.GoodCode = @GoodCode             		      
			
--Update Resid when Mojodi is zero or negative
				IF @Tedad <= 0 AND @Status1 = 5 AND @Status2 = 7 
					UPDATE dbo.tFacD
					SET FeeUnit = X.BuyPrice		
	                        FROM (SELECT ISNULL(BuyPrice ,0) AS BuyPrice FROM  dbo.[tGood]
	                                        WHERE dbo.tGood.Code = @GoodCode  
	                                )X  
	                  	WHERE  dbo.tFacD.intSerialNo = @intSerialNo
	                                AND dbo.tFacD.Branch = @Branch
	                                AND dbo.tFacD.GoodCode = @GoodCode             		      

                    SET @NumberOfRecords = @NumberOfRecords + 1
                    FETCH NEXT FROM Havale INTO @Branch ,@intSerialNo,@fDate --, @GoodCode 
	
                END
	
            CLOSE Havale
            DEALLOCATE Havale
           
	FETCH NEXT  FROM GoodsList INTO @GoodCode , @BuyPrice

        END
    CLOSE GoodsList
    DEALLOCATE GoodsList
--=====================================================
PRINT '***********'
	EXEC dbo.Update_FacDFinalPrice 
	    @InventoryNo ,
	    @AccountYear, -- smallint
	    @GoodCode1 , -- int
	    @Flag , -- int
	    @BeforeDate , -- nvarchar(8)
	    @AfterDate , -- nvarchar(8)
	    @NumberOfRecords  -- int
PRINT '***********'
	END
	RETURN @NumberOfRecords


IF @@ERROR <> 0
    AND @@TRANCOUNT > 0 
    ROLLBACK TRANSACTION ;



GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Benefit_Loss]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Benefit_Loss]
GO


CREATE  PROCEDURE [dbo].[Get_Benefit_Loss]
    (
      @DateBefore NVARCHAR(8) ,
      @DateAfter NVARCHAR(8) ,
      @AccountYear SMALLINT ,
      @InventoryNo INT ,
      @GoodLevel1 INT ,
      @SelectedLevelsString NVARCHAR(4000)
    )
AS 
     BEGIN
--SET NOCOUNT ON added to prevent extra result sets FROM interfering with SELECT statements.
SET NOCOUNT ON ;
      
-- DECLARE @DateBefore NVARCHAR(50) ;
-- DECLARE @DateAfter NVARCHAR(50) ;
-- DECLARE @AccountYear SMALLINT ;
-- DECLARE @Branch INT ;
-- DECLARE @InventoryNo INT ;
-- DECLARE      @GoodLevel1 INT 
-- DECLARE      @SelectedLevelsString NVARCHAR(4000)
-- 
-- SELECT  @DateBefore = N'88/01/01' ;
-- SELECT  @DateAfter = N'88/12/30' ;
-- SELECT  @AccountYear = 1388 ;
-- SELECT  @Branch = 1 ;
-- SELECT  @InventoryNo = 100 ;
-- SELECT @GoodLevel1 = -1
-- SELECT @SelectedLevelsString = N''
DECLARE @SaleDiscountTotal BIGINT
DECLARE @DiscountFacD BIGINT
SELECT @SaleDiscountTotal = SUM(ISNULL(DiscountTotal , 0))  
	                      FROM      [dbo].[tFacM] 
	                        WHERE     [tFacM].[Date] >= @DateBefore
                                AND [tFacM].[Date] <= @DateAfter
                                AND [tFacM].[AccountYear] = @AccountYear
                                AND [tFacM].[Recursive] = 0
				AND tFacM.Status = 2
				AND dbo.tFacM.intSerialNo IN 
				(
				SELECT intSerialNo FROM dbo.tFacD 
				INNER JOIN dbo.vw_Good ON dbo.tFacD.GoodCode = dbo.vw_Good.Code
                                WHERE  [tFacD].[intInventoryNo] = @InventoryNo
		                AND ( [vw_Good].[Level1] = @GoodLevel1
		                      OR @GoodLevel1 = -1
		                    )
		                AND ( [vw_Good].[Level2] IN (
		                      SELECT    CAST(Word AS INT)
		                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
		                                                              N',') )
		                      OR @SelectedLevelsString = N''
		                    )
				)

	SELECT 	@DiscountFacD = SUM(( [tFacD].[Amount] * [tFacD].[FeeUnit] ) * ( [tFacD].[Discount] / 100 )) 
                      FROM      [dbo].[tFacM] 
                                INNER JOIN [dbo].[tFacD]  ON [tFacM].[Branch] = [tFacD].[Branch]
                                                      AND [tFacM].[intSerialNo] = [tFacD].[intSerialNo]
				INNER JOIN [vw_Good] ON vw_Good.Code = tFacD.GoodCode
	                        WHERE     [tFacM].[Date] >= @DateBefore
                                AND [tFacM].[Date] <= @DateAfter
                                AND [tFacM].[AccountYear] = @AccountYear
                                AND [tFacM].[Recursive] = 0
                                AND [tFacD].[intInventoryNo] = @InventoryNo
				AND tFacM.Status = 2
		                AND ( [vw_Good].[Level1] = @GoodLevel1
		                      OR @GoodLevel1 = -1
		                    )
		                AND ( [vw_Good].[Level2] IN (
		                      SELECT    CAST(Word AS INT)
		                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
		                                                              N',') )
		                      OR @SelectedLevelsString = N''
		                    )

	SET @SaleDiscountTotal = @SaleDiscountTotal - @DiscountFacD

        SELECT  [tInventory_Good].GoodCode  , 
                [dbo].[vw_Good].[Name] ,
                [dbo].[vw_Good].[BarCode] ,
				FirstMojodi , FirstPrice ,
				CAST([dbo].[tInventory_Good].[FirstMojodi]
                  * [dbo].[tInventory_Good].[FirstPrice] AS BIGINT) AS TotalFirstPrice ,
            ISNULL(CAST(T3.TotalBuyAmount AS BIGINT), 0) AS TotalBuyAmount ,
            ISNULL(CAST(T3.TotalBuyReturnAmount AS BIGINT), 0) AS TotalBuyReturnAmount ,
            ISNULL(CAST(T3.TotalLossAmount AS BIGINT), 0) AS TotalLossAmount ,
            ISNULL(CAST(T3.TotalHavalehAmount AS BIGINT), 0) AS TotalHavalehAmount ,
            ISNULL(CAST(T3.TotalResidAmount AS BIGINT), 0) AS TotalResidAmount ,
            Mojodi , [dbo].[tInventory_Good].[MojodiPrice] ,
            CAST([dbo].[tInventory_Good].[Mojodi]
                  * [dbo].[tInventory_Good].[MojodiPrice] AS BIGINT) AS TotalMojodiPrice ,
            ISNULL(CAST(T3.TotalSellAmount AS BIGINT), 0) AS TotalSellAmount ,
            ISNULL(CAST(T3.TotalFinalAmount AS BIGINT), 0) AS TotalFinalAmount ,
            ISNULL(CAST(T3.TotalSellReturnAmount AS BIGINT), 0) AS TotalSellReturnAmount ,
            ISNULL(CAST(T3.TotalFinalReturnAmount AS BIGINT), 0) AS TotalFinalReturnAmount ,
		(ISNULL(CAST(T3.TotalSellAmount AS BIGINT), 0) - ISNULL(CAST(T3.TotalFinalReturnAmount AS BIGINT), 0) ) -
		(ISNULL(CAST(T3.TotalFinalAmount AS BIGINT), 0) - ISNULL(CAST(T3.TotalFinalReturnAmount AS BIGINT), 0)) AS GoodBenefitLoss ,
		@SaleDiscountTotal AS SaleDiscountTotal ,@DiscountFacD AS DiscountFacD ,
		
                [dbo].[vw_Good].[TechnicalNo] ,
                [dbo].[vw_Good].[Unit] ,
                [dbo].[vw_Good].[UnitDescription] ,
                CAST([dbo].[tInventory_Good].[FirstMojodi]
                  * [dbo].[tInventory_Good].[FirstPrice] +  ISNULL(T3.TotalBuyAmount, 0) - ISNULL(T3.TotalLossAmount, 0)
				- ISNULL(T3.TotalBuyReturnAmount, 0) - ISNULL(T3.TotalHavalehAmount, 0) +ISNULL(T3.TotalResidAmount, 0)  AS BIGINT) AS TotalMojodiPrice2 
                FROM    [dbo].[vw_Good]
                INNER JOIN [dbo].[tInventory_Good] ON [dbo].[vw_Good].[Code] = [dbo].[tInventory_Good].[GoodCode] 
                AND [dbo].[tInventory_Good].[InventoryNo] = @InventoryNo
                AND [dbo].[tInventory_Good].[AccountYear] = @AccountYear

	FULL OUTER JOIN 
		(
                  SELECT    GoodCode ,
                            ISNULL(SUM(TotalBuyAmount), 0) AS TotalBuyAmount ,
                            ISNULL(SUM(TotalSellAmount), 0) AS TotalSellAmount ,
                            ISNULL(SUM(TotalLossAmount), 0) AS TotalLossAmount ,
                            ISNULL(SUM(TotalBuyReturnAmount), 0) AS TotalBuyReturnAmount ,
                            ISNULL(SUM(TotalSellReturnAmount), 0) AS TotalSellReturnAmount ,
                            ISNULL(SUM(TotalHavalehAmount), 0) AS TotalHavalehAmount ,
                            ISNULL(SUM(TotalResidAmount), 0) AS TotalResidAmount ,
                            ISNULL(SUM(TotalFinalAmount), 0) AS TotalFinalAmount ,
                            ISNULL(SUM(TotalFinalReturnAmount), 0) AS TotalFinalReturnAmount
                  FROM      (
                              SELECT    [D].[GoodCode] ,
                                        CASE WHEN [M].Status = 1
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalBuyAmount ,
                                        CASE WHEN [M].Status = 2
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalSellAmount ,
                                        CASE WHEN [M].Status = 2
                                             THEN SUM([D].[Amount]
                                                      * [D].[FinalPrice])
                                             ELSE 0
                                        END AS TotalFinalAmount ,
                                        CASE WHEN [M].Status = 3
                                             THEN SUM([D].[Amount]
                                                      * [D].[FeeUnit])
                                             ELSE 0
                                        END AS TotalLossAmount ,
                                        CASE WHEN [M].Status = 4
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalBuyReturnAmount ,
                                        CASE WHEN [M].Status = 5
                                             THEN SUM(( [D].[Amount]
                                                        * [D].[FeeUnit] )
                                                      * ( 1 - ( [D].[Discount]
                                                              / 100 ) ))
                                             ELSE 0
                                        END AS TotalSellReturnAmount ,
                                        CASE WHEN [M].Status = 5
                                             THEN SUM([D].[Amount]
                                                      * [D].[FinalPrice])
                                             ELSE 0
                                        END AS TotalFinalReturnAmount ,
                                        CASE WHEN [M].Status = 6
                                             THEN SUM([D].[Amount]
                                                      * [D].[FeeUnit])
                                             ELSE 0
                                        END AS TotalHavalehAmount ,
                                        CASE WHEN [M].Status = 7
                                             THEN SUM([D].[Amount]
                                                      * [D].[FeeUnit])
                                             ELSE 0
                                        END AS TotalResidAmount
                              FROM      [dbo].[tFacM] M
                                        INNER JOIN [dbo].[tFacD] D ON [M].[Branch] = [D].[Branch]
                                                              AND [M].[intSerialNo] = [D].[intSerialNo]
										INNER JOIN dbo.tGood ON [D].GoodCode = dbo.tGood.Code
                              WHERE     [M].[Date] >= @DateBefore
                                        AND [M].[Date] <= @DateAfter
                                        AND [M].[AccountYear] = @AccountYear
                                        AND [M].[Recursive] = 0
                                        AND [D].[intInventoryNo] = @InventoryNo
                                        AND (dbo.tGood.GoodType = 1 OR dbo.tGood.GoodType = 3)
                              GROUP BY  [D].[GoodCode] ,
                                        [M].[Status]
                            ) T1
                  GROUP BY  [T1].[GoodCode]
                ) T3
                ON [T3].[GoodCode] = [dbo].[tInventory_Good].[GoodCode]
        	WHERE   [dbo].[tInventory_Good].[InventoryNo] = @InventoryNo
                AND [dbo].[tInventory_Good].[AccountYear] = @AccountYear
                AND (vw_Good.GoodType = 1 OR vw_Good.GoodType = 3)
                AND ( [vw_Good].[Level1] = @GoodLevel1
                      OR @GoodLevel1 = -1
                    )
                AND ( [vw_Good].[Level2] IN (
                      SELECT    CAST(Word AS INT)
                      FROM      [dbo].[SplitWithDelimiterNVarChar](@SelectedLevelsString,
                                                              N',') )
                      OR @SelectedLevelsString = N''
                    )
        ORDER BY [tInventory_Good].[GoodCode] ASC
    END
--===============================================




GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_Good_Store_FirstMojodi]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_Good_Store_FirstMojodi
GO


CREATE  PROCEDURE dbo.Update_Good_Store_FirstMojodi
(
	@FirstMojodi	Float,
	@FirstPrice	INT,
	@Code		INT ,
	@InventoryNo INT ,
	@AccountYear Smallint
	
)

AS

    UPDATE dbo.tInventory_Good

	SET   
  	       FirstMojodi    = @FirstMojodi ,
	       FirstPrice = @FirstPrice ,
	       [Time] = dbo.setTimeFormat(getdate()) ,
	       [Date] = dbo.Shamsi(GETDATE()) 
	Where GoodCode = @Code And InventoryNo = @InventoryNo  And AccountYear = @AccountYear



GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Update_tblTotal_tInventory_tGood_For_FinalPrice]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Update_tblTotal_tInventory_tGood_For_FinalPrice
GO

CREATE  PROCEDURE dbo.Update_tblTotal_tInventory_tGood_For_FinalPrice
(  
 @SystemDate   NVARCHAR(50),  
 @SystemDay    NVARCHAR(50),  
 @SystemTime   NVARCHAR(50),   
 @DateBefore   NVARCHAR(50),  
 @DateAfter    NVARCHAR(50),  
 @Type  int  ,  
 @InventoryNo Int ,  
 @AccountYear Smallint ,
 @ZeroNegative BIT 
)   
  
AS  
BEGIN TRAN  
/*
INSERT INTO   tInventory_Good  
  (Branch ,   InventoryNo ,  GoodCode , BuyAmount , SaleAmount , BuyReturnAmount , SaleReturnAmount  ,  
  FromStoreAmount , toStoreAmount  , Mojodi , AccountYear )   
 SELECT  
    T1.Branch , T1.InventoryNo  ,T1.GoodCode , T1.BuyAmount,  T1.SaleAmount ,  T1.BuyReturnAmount ,  T1.SaleReturnAmount ,  
  T1.FromStoreAmount , T1.toStoreAmount ,T1.Mojodi , @AccountYear  
  
 FROM dbo.tblTotal_tInventory_tGood_For_FinalPrice  
  (  
   @DateBefore   ,  
   @DateAfter    ,  
   @Type    ,  
   @InventoryNo  ,  
   @Branch  ,  
   @AccountYear  
  )  
  AS T1  
    
  WHERE 0=(Select Count(GoodCode) From tInventory_Good Where GoodCode = T1.GoodCode And InventoryNo = T1.InventoryNo and Branch = T1.Branch And AccountYear = @AccountYear)  
    
  
IF  @@Error <> 0   
 goto ErrHandler  
*/  

	UPDATE dbo.tInventory_Good
	SET BuyPriceAverage = FirstPrice , MojodiPrice = FirstPrice
	WHERE InventoryNo = @InventoryNo AND AccountYear = @AccountYear
	
	UPDATE dbo.tInventory_Good
	SET BuyPriceAverage = T.AverageBuyFee , MojodiPrice = T.AverageBuyFee  FROM (
	Select IsNull(Sum(FeeUnit * Amount) ,0) + ISNULL(FirstPrice * FirstMojodi , 0) /(ISNULL(Sum(Amount),1) + ISNULL(FirstMojodi ,1)) AS AverageBuyFee , dbo.tInventory_Good.GoodCode , tInventory_Good.AccountYear , tInventory_Good.InventoryNo
--	Select IsNull(Sum(FeeUnit * Amount) ,0) /(ISNULL(Sum(Amount),1) ) AS AverageBuyFee , dbo.tInventory_Good.GoodCode , tInventory_Good.AccountYear , tInventory_Good.InventoryNo
	From tFacM inner join tfacd On tFacM.intSerialNo = tfacd.intSerialNo and tFacM.Branch = tfacd.Branch AND dbo.tFacD.intInventoryNo = @InventoryNo
	INNER JOIN dbo.tInventory_Good ON tfacD.intInventoryNo = dbo.tInventory_Good.InventoryNo AND tfacd.GoodCode = dbo.tInventory_Good.GoodCode AND dbo.tInventory_Good.AccountYear = @AccountYear
	Where tfacm.Status = 1 and Recursive = 0 And tfacm.AccountYear = @AccountYear AND tfacD.intInventoryNo = @InventoryNo 
	GROUP BY tfacd.GoodCode ,tInventory_Good.GoodCode, tInventory_Good.AccountYear , tInventory_Good.InventoryNo ,  tInventory_Good.FirstMojodi  ,  tInventory_Good.FirstPrice)T
	WHERE tInventory_Good.AccountYear = t.AccountYear  AND dbo.tInventory_Good.InventoryNo = t.InventoryNo AND tInventory_Good.GoodCode = t.GoodCode

UPDATE  tInventory_Good  
    
 Set    BuyAmount = T2.BuyAmount,  
		SaleAmount = T2.SaleAmount ,  
		LossAmount = T2.LossAmount ,
		BuyReturnAmount = T2.BuyReturnAmount ,  
		SaleReturnAmount = T2.SaleReturnAmount ,  
		FromStoreAmount = T2.FromStoreAmount ,  
		toStoreAmount = T2.toStoreAmount ,  
		Mojodi = T2.Mojodi , 
 	    MojodiPrice = CASE tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount  WHEN 0 THEN 0 ELSE ( firstMojodiRial + BuyRial ) / (tInventory_Good.FirstMojodi + tInventory_Good.BuyAmount) END 
   
 FROM dbo.tblTotal_tInventory_tGood_For_FinalPrice  
  (  
   @DateBefore   ,  
   @DateAfter    ,  
   @Type    ,  
   @InventoryNo  ,  
   @AccountYear  
  )  
   AS T2    
     Where tInventory_Good.GoodCode = T2.GoodCode And tInventory_Good.InventoryNo = T2.InventoryNo and tInventory_Good.AccountYear = @AccountYear  
	if @@Error <> 0   
	 goto ErrHandler  
  

--	UPDATE  dbo.tInventory_Good 
--	Set  MojodiPrice = ISNULL(( firstMojodiRial + BuyRial - FromStoreRial - LossRial + toStoreRial ) / ABS(CASE tInventory_Good.Mojodi WHEN 0 THEN 1 ELSE tInventory_Good.Mojodi END ) , 0 )
-- 	
-- FROM dbo.tblTotal_tInventory_tGood_For_FinalPrice  
--  (  
--   @DateBefore   ,  
--   @DateAfter    ,  
--   @Type    ,  
--   @InventoryNo  ,  
--   @AccountYear  
--  )  
--   AS T3    
--     Where tInventory_Good.GoodCode = T3.GoodCode And tInventory_Good.InventoryNo = T3.InventoryNo and tInventory_Good.AccountYear = @AccountYear  
--     	AND t3.Mojodi <> 0
--	if @@Error <> 0   
--	 goto ErrHandler  
 	  
UPDATE  tInventory_Good  
	SET MojodiPrice = 0 WHERE Mojodi = 0

IF @ZeroNegative = 1
UPDATE  tInventory_Good  
	SET MojodiPrice = 0 WHERE MojodiPrice < 0

Commit Tran   
  
return  
  
ErrHandler:  
RollBack Tran  
return 
  




GO


ALTER PROCEDURE [dbo].[InsertFactorMasterDetails]  (      

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
	select @destbranch=branch from tInventory where inventoryNo=(SELECT TOP 1  DestInventoryNo FROM Split(@DetailsString))      

          SET @NO2 = (SELECT ISNULL(MAX([NO]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @destbranch AND AccountYear = @AccountYear)      
		  SET @TempNo2 = (SELECT ISNULL(MAX([TempNo]),0)+1 FROM tFacM WHERE Status=7  And Branch =  @destbranch AND Date = @Date AND ShiftNo = @ShiftNo)      

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
				@DestBranch , --dbo.Get_Current_Branch(),      
				@AccountYear ,      
				@NvcDescription,      
				@TempAddress,
				@GuestNo ,
				@TempNo2    
		
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


ALTER  Procedure dbo.Insert_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int, 
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Code Bigint out 

)  

as  

Begin Tran  

--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  
if @MasterCode = 0   
  Set @MasterCode = Null  
if @MasterCode is not Null  
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  )  --AND (Branch = @Branch )
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode   )--AND (Branch = @Branch )
 end   
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId ) <> 0 --AND Branch = @Branch)   
  Goto ErrHandler   

Set @Code = (Select  isnull(Max(Code),0) + 1 from tCust where code > 0  And  ( Branch = @Branch  Or Branch Is NULL ) )
If @Code < (@Branch * 100000 ) Set @Code = (@Branch * 100000 )

Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

if @nvcRFID = N''  
  SET @nvcRFID=N'-999'  

insert Into dbo.tCust  
(   
	Code,   
	MembershipId,   
	MasterCode,   
	Owner,   
	Name,   
	Family,   
	Sex,   
	WorkName,   
	InternalNo,   
	Unit,   
	City,   
	ActKind,   
	ActDeAct,  
	Prefix,   
	Assansor,   
	Address,   
	PostalCode,   
	Tel1,   
	Tel2,   
	Tel3,   
	Tel4,   
	Mobile,   
	Fax,   
	Email,   
	Flour,   
	CarryFee,   
	PaykFee,   
	Distance,   
	Credit,   
	Discount,   
	BuyState,   
	[Description],   
	[Date],   
	[Time],   
	[User],  
	FamilyNo ,  
	Member ,  
	State ,  
	Central ,  
	Branch,  
	nvcRFID,  
	sellprice ,
	EconomicCode ,
	nvcBirthDate ,
	TotalRemainingAmount
	
)  
values  
(   
	@Code ,  
	@MembershipId,   
	@MasterCode,   
	@Owner,   
	@Name,   
	@Family,   
	@Sex,   
	@WorkName,   
	@InternalNo,   
	@Unit,   
	@City,   
	@ActKind,   
	@ActDeAct,  
	@Prefix,   
	@Assansor,   
	@Address,   
	@PostalCode,   
	@Tel1,   
	@Tel2,   
	@Tel3,   
	@Tel4,   
	@Mobile,   
	@Fax,   
	@Email,   
	@Flour,   
	@CarryFee,   
	@PaykFee,   
	@Distance,   
	@Credit,   
	@Discount,   
	@BuyState,   
	@Description,   
	@Date,   
	@Time,   
	@User ,  
	@FamilyNo ,  
	@Member ,  
	@State ,  
	@Central ,  
	@Branch,  
	@nvcRFID,  
	@sellprice   ,
	@EconomicCode ,
	@nvcBirthDate ,
	@TotalRemainingAmount
	
)  
if @@Error <> 0   
 goto ErrHandler  

--Set @Code = @@Identity  
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
  and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address)  
 , nvcRFID=CAST(Branch AS NVARCHAR(1))+CAST(Code AS NVARCHAR(8))  
  where code=@code  AND Branch = @Branch 



Commit Tran   
return @Code  

ErrHandler:  
RollBack Tran  
Set @Code = -1  
return @Code




GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER  Procedure dbo.Update_Cust  
( 
	@MembershipId nVarChar(50) ,   
	@MasterCode int,  
	@Owner int ,   
	@Name nVarChar(50),   
	@Family nVarChar(50),   
	@Sex int,   
	@WorkName nVarChar(50),   
	@InternalNo nVarChar(50),   
	@Unit nVarChar(50),   
	@City int,   
	@ActKind int,   
	@ActDeAct int,  
	@Prefix int,   
	@Assansor int,   
	@Address nVarChar(255),   
	@PostalCode nVarChar(50),   
	@Tel1 nVarChar(50),   
	@Tel2 nVarChar(50),   
	@Tel3 nVarChar(50),   
	@Tel4 nVarChar(50),   
	@Mobile nVarChar(50),   
	@Fax nVarChar(50),   
	@Email nVarChar(50),   
	@Flour nVarChar(50),   
	@CarryFee Float,   
	@PaykFee Float,   
	@Distance int,   
	@Credit Float,   
	@Discount Float,   
	@BuyState int,   
	@Description nVarChar(255),   
	@User int ,   
	@Code Bigint ,  
	@FamilyNo int ,  
	@Member Bit ,  
	@State int ,  
	@Central Bit ,  
	@Sellprice smallint,  
	@EconomicCode NVARCHAR(20) ,
	@nvcRFID NVARCHAR(20)=N''  ,
	@nvcBirthDate NVARCHAR(10)=N''  ,
	@TotalRemainingAmount INT  ,
	@Branch INT ,
	@Updated Bigint out  

)  

as  

Begin Tran  
--IF  @Branch IS NULL Set @Branch = dbo.Get_Current_Branch()  

if @MasterCode = 0   
 Set @MasterCode = Null  
if @MasterCode is not Null    
 begin  
   Set @MembershipId = (Select top 1 MembershipId from  dbo.tCust where  Code = @MasterCode  ) --AND (Branch = @Branch ) )  
   Set @BuyState = (Select top 1 BuyState from  dbo.tCust where  Code = @MasterCode )   --AND (Branch = @Branch ) )  
 end  
else   

 if (Select top 1 isnull(Code , 0) from tCust where MembershipId = @MembershipId and Code <> @Code and MasterCode <> @Code  ) <> 0  -- AND (Branch = @Branch )    
  Goto ErrHandler   
 else  

  Update dbo.tCust     
   Set MembershipId = @MembershipId   

   Where MasterCode = @Code   AND (Branch = @Branch )  



Declare @Time nVarchar(50)  
Set @Time = (select dbo.setTimeFormat(GETDATE()))  

Declare @Date nVarchar(50)  
Set @Date =(Select dbo.Shamsi(GETDATE()))  

Update dbo.tCust  

 Set MembershipId = @MembershipId ,  
	MasterCode  = @MasterCode ,    
	Owner = @Owner ,  
	Name = @Name ,  
	Family = @Family ,  
	Sex = @Sex ,  
	WorkName = @WorkName ,   
	InternalNo = @InternalNo ,  
	Unit = @Unit ,  
	City = @City ,  
	ActKind = @ActKind ,  
	ActDeAct = @ActDeAct ,  
	Prefix = @Prefix ,  
	Assansor = @Assansor ,  
	Address = @Address ,  
	PostalCode = @PostalCode ,  
	Tel1 = @Tel1 ,  
	Tel2 = @Tel2 ,  
	Tel3 = @Tel3 ,  
	Tel4 = @Tel4 ,  
	Mobile = @Mobile ,  
	Fax = @Fax ,  
	Email = @Email ,  
	Flour = @Flour ,  
	CarryFee = @CarryFee ,  
	PaykFee = @PaykFee ,  
	Distance = @Distance ,  
	Credit = @Credit ,  
	Discount = @Discount ,  
	BuyState = @BuyState ,  
	[Description] = @Description ,  
	[Date] = @Date ,  
	[Time] = @Time ,  
	[User] = @User ,  
	FamilyNo = @FamilyNo ,  
	Member = @Member ,  
	State = @State ,  
	Central = @Central,  
	Sellprice=@Sellprice  ,
	EconomicCode = @EconomicCode ,
	nvcRFID = @nvcRFID ,
	nvcBirthDate = @nvcBirthDate ,
	TotalRemainingAmount = @TotalRemainingAmount
	
Where Code = @Code   AND (Branch = @Branch )   

if @@Error <> 0   
 goto ErrHandler  


Set @Updated = @Code   
 UPDATE dbo.tCust  
 SET Address = tmpCust.Address  
 FROM dbo.tCust  , dbo.tCust tmpCust  
 WHERE dbo.tCust.MasterCode = tmpCust.Code  
 and (dbo.tCust.[Branch] = tmpCust.[Branch] )  
update tcust set address=dbo.addressedit(address) where code=@code  AND Branch = @Branch 
 


Commit Tran  
return @Updated  

ErrHandler:  
RollBack Tran  
return -1



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  PROCEDURE [dbo].[InsertPersonel]( 
	@PersonnelNumber nvarchar(50),
	@nvcFirstName nvarchar(50),
	@nvcSurName nvarchar(50),
	@Gender bit,
	@IdNumber nvarchar(50),
	@Job int,
	@InsuranceNo nvarchar(50) ,
	@Address nvarchar(300),
	@Tel nvarchar(30),
	@User int , 
	@UserName nvarchar(50) ,
	@Password nvarchar(50) ,
	@intAccessLevel int ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno int out

	)
 AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

set @Time = dbo.SetTimeFormat(getdate())

select @pPno = isnull(max(Ppno),0) + 1 from tper Where Branch = @Branch 
If @pPno < (@Branch * 1000 ) Set @pPno = (@Branch * 1000 )


begin Tran
insert into dbo.tper (
	pPno ,
	PersonnelNumber,
	nvcFirstName,
	nvcSurName,
	Gender ,
	IdNumber,
	Job ,
	InsuranceNo  ,
	Address ,
	Tel ,
	[Date] ,
	[Time] ,			
	[User] ,
	Branch,
	MaxCredit,
	ActDeAct

)
values(
	@pPno ,
	@PersonnelNumber,
	@nvcFirstName,
	@nvcSurName ,
	@Gender ,
	@IdNumber ,
	@Job ,
	@InsuranceNo ,
	@Address ,
	@Tel ,
	@Date,
	@Time ,
	@User ,
	@Branch,
	@MaxCredit,
	@ActDeAct
)
if @@Error <> 0 
		GOTO EventHandler	


--set @pPno=@@identity
DECLARE @UID INT
if @intAccessLevel<>0 and @UserName <> '' and @Password<>''


select @Uid = isnull(max(Uid),0) + 1 from tUser Where Branch = @Branch 
If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )

BEGIN
	insert into dbo.tUser 
	(
		[Uid] ,
	 UserName ,
	 [Password] ,
	 intAccessLevel ,
	 pPno ,
	 addUser , 
	 Branch, 
	 CountRePrint, 
	 CountInvoicePrint,
	 CountInvoiceEditable,
	 CountInvoiceRefferable
	)
 values (
	@UID ,					
	@UserName  ,
	@Password  ,
	@intAccessLevel ,
	@pPno , 
	@User ,
	@Branch,
	@CountRePrint,
	@CountInvoicePrint,
	@CountInvoiceEditable,
	@CountInvoiceRefferable
	)

if @@Error <> 0 
		GOTO EventHandler	
--SET @UID = @@IDENTITY		
END	



commit Tran



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1




GO


SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO



ALTER  PROCEDURE [dbo].[UpdatePersonel]( 
	@CurrentPPNO 		INT,
	@PersonnelNumber 	NVARCHAR(50),
	@nvcFirstName 		NVARCHAR(50),
	@nvcSurName	 	NVARCHAR(50),
	@Gender 		BIT,
	@IdNumber 		NVARCHAR(50),
	@Job 			INT,
	@InsuranceNo 		NVARCHAR(50) ,
	@Address 		NVARCHAR(300),
	@Tel 			NVARCHAR(30),
	@User 			INT , 
	@UID 			INT ,
	@UserName 		NVARCHAR(50) ,
	@Password 		NVARCHAR(50) ,
	@intAccessLevel 	INT ,
	@MaxCredit INT=0,
	@ActDeAct BIT,
	@CountRePrint INT=0,
	@CountInvoicePrint INT=0,
	@CountInvoiceEditable INT=0,
	@CountInvoiceRefferable INT=0,
	@Branch INT = 1 ,
	@pPno 			INT OUT
	       )
AS
Declare @Date nvarchar(50)
Declare @Time nvarchar(50)

SET @Date = (SELECT GETDATE())
SET @Date = dbo.Shamsi(@Date)

SET @Time= dbo.SetTimeFormat(getdate())

BEGIN TRANSACTION

	UPDATE tPer
		SET PersonnelNumber 	= @PersonnelNumber,
		    nvcFirstName    	= @nvcFirstName,
		    nvcSurName	    	= @nvcSurName,
		    Gender	    	= @Gender,
		    IdNumber       	= @IdNumber,
		    Job		    	= @Job,
		    InsuranceNo     	= @InsuranceNo,
		    Address	    	= @Address,
		    Tel   	    	= @Tel,
		    [Date]	    	= @Date,
		    [Time]	    	= @Time,
		    [User]	    	= @User,
		    MaxCredit		=@MaxCredit,
		    ActDeAct 		=@ActDeAct
	WHERE       pPNO = @CurrentPPNO And Branch = @Branch 


	IF @@ERROR <> 0 
		GOTO EventHandler	

	set @pPno = @CurrentPPNO

	IF @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID<>0
		UPDATE tUser
			SET 		UserName       	= @UserName,
	        	   		 	[Password]     	= @Password,
			    		intAccessLevel 	= @intAccessLevel,
			    		pPno           	= @pPno,
			    		addUser        	= @User,
					 CountRePrint		=@CountRePrint,
		  			 CountInvoicePrint	=@CountInvoicePrint,
					 CountInvoiceEditable		=@CountInvoiceEditable,
		  			 CountInvoiceRefferable	=@CountInvoiceRefferable
		WHERE   UID = @UID And Branch = @Branch  
	else 
		if @intAccessLevel <>0 AND @UserName <> '' AND @Password<>'' and @UID=0
		BEGIN 
			select @Uid = isnull(max(Uid),0) + 1 from tUser Where Branch = @Branch   
			If @Uid < (@Branch * 1000 ) Set @Uid = (@Branch * 1000 )  
			insert into dbo.tUser (
						UID ,
						UserName , 
						[Password] , 
						intAccessLevel , 
						pPno , 
						addUser , 
						Branch,
						CountRePrint,
						CountInvoicePrint,
						CountInvoiceEditable,
		  			 	CountInvoiceRefferable
			) values (	
						@UID ,				
						@UserName  ,
						@Password  ,
						@intAccessLevel ,
						@pPno , 
						@User , 
						@Branch,
						@CountRePrint,
						@CountInvoicePrint,
						@CountInvoiceEditable,
		  			 	@CountInvoiceRefferable
			)			
			END 
	IF @@ERROR <> 0 
		GOTO EventHandler	



COMMIT TRANSACTION



RETURN

EventHandler: 

	ROLLBACK TRAN
	SET @pPno = -1
	RETURN -1



GO



SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[Insert_tinventory] (
					@Description nvarchar(50) , 
					@Active bit ,
					@Branch int , 
					@InventoryNo int out )

AS

Begin Tran
set @InventoryNo=-1
Set @InventoryNo = (Select isnull(Max(InventoryNo) , 0) + 1 as InventoryNo from dbo.tinventory  
	WHERE    Branch  = @Branch )
IF @InventoryNo < @Branch * 100 SET @InventoryNo = @Branch * 100

declare @MasterCode int
select @MasterCode=InventoryNo from tinventory where branch=@Branch  and MasterCode is null 
--if  ( @MasterCode is null) or ( @MasterCode  is not null)
--	Goto ErrHandler

Insert Into dbo.tinventory
(InventoryNo , [Description] ,MasterCode,  Active , Branch)
values
( @InventoryNo , @Description ,@MasterCode,  @Active , @Branch)
 --set @InventoryNo=@@identity

if @@Error <> 0 
	Goto ErrHandler

Commit Tran




Return

ErrHandler:
RollBack Tran
Set @InventoryNo = -1
Return




GO

SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO


ALTER   PROCEDURE [dbo].[InsertTable] 
(
	@Name 		NVARCHAR(50), 
	@NumberOfChair 	INT, 
	@Person 	INT, 
	@PartitionID	INT,	
	@Empty	Bit,
	@Reserve	Bit,
	@nvcMaxUseTime  NVARCHAR(10) ,
	@Branch INT ,
	@No 		INT Out 

) 
AS

	IF @Person < 1 
		SET @Person =null
	BEGIN TRAN

	Set @No = (Select isnull(Max(No) , 0) + 1 from dbo.tTable  
		WHERE    Branch  = @Branch )
	IF @No < @Branch * 1000 SET @No = @Branch * 1000

		INSERT INTO dbo.tTable 	( No , [Name] , NumberOfChair , Person ,  PartitionID , Empty , Reserve ,Branch , nvcMaxUseTime) 
		VALUES 
					( @No , @Name , @NumberOfChair , @Person ,  @PartitionID ,@Empty , @Reserve ,  @Branch ,@nvcMaxUseTime ) 

		IF @@Error <> 0 GOTO ErrHandler

		--SET @No=@@Identity
	COMMIT TRAN


	RETURN @No

ErrHandler:
	ROLLBACK TRAN
	SET @No= -1
	RETURN @No

GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO




ALTER   PROCEDURE dbo.InsertPartition 
(
	@intLanguage	INT,
	@ServicePercentDefault 	INT,
	@PartitionName	NVARCHAR(50),	
	@ReserveServiceRate 	INT,
	@Branch INT ,
	@PartitionID	INT OUT
) 
AS
	BEGIN TRAN
		
	Set @PartitionID = (Select isnull(Max(PartitionID) , 0) + 1 from dbo.tPartitions  
		WHERE    Branch  = @Branch )
	IF @PartitionID < @Branch * 100 SET @PartitionID = @Branch * 100

		IF @intLanguage = 0 
			INSERT INTO dbo.tPartitions ( PartitionID , PartitionDescription , PartitionLatinDescription ,DefaultServicePercent,ReserveServiceRate, Branch )
				             VALUES (@PartitionID , @PartitionName , @PartitionName ,@ServicePercentDefault, @ReserveServiceRate, @Branch )
		IF @intLanguage = 1
			INSERT INTO dbo.tPartitions ( PartitionID ,PartitionLatinDescription , PartitionDescription, DefaultServicePercent,ReserveServiceRate, Branch )
				             VALUES ( @PartitionID ,@PartitionName , @PartitionName , @ServicePercentDefault , @ReserveServiceRate, @Branch )
		
		IF @@Error <> 0 GOTO ErrHandler	
		--SET @PartitionID = @@Identity

	COMMIT TRAN
	RETURN @PartitionID
	
ErrHandler:
	ROLLBACK TRAN
	SET @PartitionID = -1
	RETURN @PartitionID



GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO



ALTER    PROCEDURE [dbo].[Insert_tblAcc_Cash]( 
	@No INT ,
	@List int,
	@Date nvarchar(50),
	@Uid int,
	@Description nvarchar(300) ,
	@Bestankar Bigint,
	@PaymentType int,
	@Uid_Bede int ,
	@AddUser int ,
	@AccountYear SMALLINT,
	@Branch INT 		
)

 AS

Declare @RegDate nvarchar(50)
Declare @Time nvarchar(50)

SET @RegDate = (SELECT GETDATE())
SET @RegDate = dbo.Shamsi(@RegDate)

set @Time = dbo.SetTimeFormat(getdate())

DECLARE @Code INT 
begin Tran

IF @No = 0
	set @No = (Select isnull(max([No]),0)+ 1  From tblAcc_Cash  Where  Branch = @Branch And AccountYear = @AccountYear)

	Set @Code = (Select isnull(Max(Code) , 0) + 1 from dbo.tblAcc_Cash  
		WHERE    Branch  = @Branch )
	IF @Code < @Branch * 1000000 SET @Code = @Branch * 1000000

insert into dbo.tblAcc_Cash (
	Code ,
	[No],
	List,
	[Date],
	RegDate ,
	RegTime,
	Uid ,
	[Description]  ,
	Bestankar ,
	PaymentType ,
	Uid_Bede ,
	AddUser ,
	AccountYear ,
	Branch
)
values(
	@Code ,
	@No,
	@List,
	@Date,
	@RegDate ,
	@Time ,
	@Uid ,
	@Description  ,
	@Bestankar ,
    @PaymentType ,
	@Uid_Bede ,
	@AddUser ,
	@AccountYear ,
	@Branch
)
if @@Error <> 0 
		GOTO EventHandler	



commit Tran


RETURN

EventHandler: 

	ROLLBACK TRAN

	RETURN -1



GO


SET QUOTED_IDENTIFIER OFF
SET ANSI_NULLS ON
GO

ALTER     PROCEDURE [dbo].[Insert_tblAcc_Recieved]( 
	@No Bigint,
	@List int,
	@Date nvarchar(50),
	@Uid int,
	@Description nvarchar(300) ,
	@Bestankar Bigint,
	@RecieveType int,
	@Code_Bes Bigint,
	@AddUser int ,
	@AccountYear Smallint = Null,
	@intSerialNo BigInt = NULL,
	@Branch INT 
		)

 AS
IF @AccountYear IS NULL
	SET @AccountYear = dbo.get_AccountYear() --Left(N'13' + dbo.shamsi(Getdate()) ,4)

Declare @RegDate nvarchar(50)
Declare @Time nvarchar(50)

SET @RegDate = (SELECT GETDATE())
SET @RegDate = dbo.[Get_ShamsiDate_For_Current_Shift](@RegDate)

set @Time = dbo.SetTimeFormat(getdate())

DECLARE @Code INT 

begin Tran
If @intSerialNo = 0 SET @intSerialNo = Null

IF @No = 0 set @No = (Select isnull(max([No]),0)+ 1  From tblAcc_Recieved  Where  Branch =  @Branch And AccountYear = @AccountYear)
	Set @Code = (Select isnull(Max(Code) , 0) + 1 from dbo.tblAcc_Recieved  
		WHERE    Branch  = @Branch )
	IF @Code < @Branch * 1000000 SET @Code = @Branch * 1000000

insert into dbo.tblAcc_Recieved (
	Code ,
	[No],
	List,
	[Date],
	RegDate ,
	RegTime,
	Uid ,
	[Description]  ,
	Bestankar ,
	RecieveType ,
	Code_Bes ,
	AddUser ,
	AccountYear , 
	intSerialNo ,
	Branch

)
values(
	@Code ,
	@No,
	@List,
	@Date,
	@RegDate ,
	@Time ,
	@Uid ,
	@Description  ,
	@Bestankar ,
    @RecieveType ,
	@Code_Bes,
	@AddUser ,
	@AccountYear ,
	@intSerialNo ,
	@Branch
)
if @@Error <> 0 
		GOTO EventHandler	



commit Tran


RETURN

EventHandler: 

	ROLLBACK TRAN

	RETURN -1




GO


