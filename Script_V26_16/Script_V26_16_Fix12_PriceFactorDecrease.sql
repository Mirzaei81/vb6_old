
--Script_V26_16_Fix12_PriceFactorDecrease.sql
--اضافه شدن دسترسی کاهش مبلغ هنگام اصلاح
--اصلاح کنترل دسترسی کم کردن تعدادی کالاها
--حذف متغیراز آریا ستینگ EnableUpperAmountGood
-- برای تبلت نیز همین دسترسی ها استفاده خواهد شد
--94/02/12


INSERT INTO dbo.tObjects
        ( intObjectCode ,
          ObjectId ,
          ObjectName ,
          objectLatinName ,
          intObjectType ,
          ObjectParent
        )
VALUES  ( 315 , -- intObjectCode - int
          N'PriceFactorDecrease' , -- ObjectId - nvarchar(50)
          N'کاهش مبلغ هنگام اصلاح' , -- ObjectName - nvarchar(50)
          N'PriceFactorDecrease' , -- objectLatinName - nvarchar(50)
          2 , -- intObjectType - int
          110  -- ObjectParent - int
        )
GO


INSERT INTO dbo.tAccess_Object
        ( intAccessLevel, intObjectCode )
VALUES  ( 1, -- intAccessLevel - int
          315  -- intObjectCode - int
          )
          
go


