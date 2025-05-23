
--ScriptV26_16_SmsKiosk.sql

IF NOT EXISTS(SELECT * FROM tblPub_SmsActionList WHERE SmsActionNo = 13 )

	INSERT INTO dbo.tblPub_SmsActionList
			( SmsActionNo, SmsActionText )
	VALUES  ( 13, -- SmsActionNo - int
			  N'ارسال پیامک برای کیوسک'  -- SmsActionText - nvarchar(255)
			  )
          
--SELECT * FROM dbo.tblPub_SmsText
--GO

DECLARE @Maxid AS INT 

INSERT INTO dbo.tblPub_SmsText
        ( SmsText )
VALUES  ( N'مشتری محترم جهت تحویل سفارش خود به پیتزا خاتون مراجعه نمایید'  -- SmsText - nvarchar(255)
          )


SET @Maxid = @@IDENTITY

IF NOT EXISTS(SELECT * FROM tblPub_SmsAction WHERE SmsActionCode = 13 )

	INSERT INTO dbo.tblPub_SmsAction
			( SmsTextId ,
			  SmsActionCode ,
			  SmsConfigNoId ,
			  SmsRepSaleNo ,
			  SmsRepSaleTime ,
			  IsActive ,
			  IsPaykName ,
			  IsPaykMobile
			)
	VALUES  ( @Maxid , -- SmsTextId - int
			  13 , -- SmsActionCode - int
			  1 , -- SmsConfigNoId - int
			  '' , -- SmsRepSaleNo - varchar(1000)
			  '' , -- SmsRepSaleTime - varchar(50)
			  1 , -- IsActive - bit
			  0 , -- IsPaykName - bit
			  0  -- IsPaykMobile - bit
			)

GO

