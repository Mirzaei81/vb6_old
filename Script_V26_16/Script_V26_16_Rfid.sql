

--ScriptV26_16_Rfid.sql
--اضافه شدن کارت خوان مایفر
--در سایر تنظیمات امکان اکتیو کردن کارت خوان وجوددارد
--در وسایل جانبی باید کارت خوان مایفر با سرعت 9600انتخاب شود
--هیچگونه داده ای در کارت نوشته نمی شود فقط شماره سریال کارت در دیتابیس مشتریان نوشته می شود
--در قسمت مشترکین هنگام ثبت کد یونیک کارت در دیتابیس باید کارت روی دستگاه کارت ریدر قرارداشته باشد
--افزودن قابلیت اعمال تخفیفات لویالتی بر روی غیر مشترک در سایر تنطیمات . در نتیجه کارت خوان لازم ندارد.
--94/05/03


UPDATE dbo.tCust SET nvcRFID = '' WHERE ISNUMERIC(nvcRFID) = 1 --AND  Code = CAST(nvcRFID AS INT)
GO


INSERT INTO dbo.tDevice
        ( DeviceCode ,
          DeviceName ,
          DeviceLatinName ,
          Active ,
          DeviceTypeCode ,
          BufferSize ,
          RThreshold
        )
VALUES  ( 68 , -- DeviceCode - int
          N'کارت خوان مایفر' , -- DeviceName - nvarchar(50)
          N'RfId Reader' , -- DeviceLatinName - nvarchar(50)
          1 , -- Active - bit
          6 , -- DeviceTypeCode - int
          40 , -- BufferSize - int
          10  -- RThreshold - int
        )
        
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Customer_Rfid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Get_Customer_Rfid
GO


CREATE Proc Get_Customer_Rfid    
@ActDeact int ,    
@nvcRfid NVARCHAR(20) 
    
as    

Select TOP 1 * from dbo.tCust 
where nvcRFID = @nvcRfid and actdeact <> @ActDeact-- AND branch = @Branch  


GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Check_RFIDSerialExist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Check_RFIDSerialExist
GO


CREATE Proc Check_RFIDSerialExist    
@nvcRfid NVARCHAR(20) 
    
as    

Select TOP 1 * from dbo.tCust 
where nvcRFID = @nvcRfid  


GO




