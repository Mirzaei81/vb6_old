

--Script V5_1_V26_16  جدید مجددا روی همه زده شود
--Script V26_16_Fix5_Added2 جدید مجددا روی همه زده شود 
--Branch = 6 in tBranch ساخته شود
--در tInventory Branch 3==>6 , 4==>3 , 6==>4
-- Branch = 6 in tBranch حذف گردد

--بعد دستورهای ذیل اجرا گردد
--tFacm در سرور مرکزی رکوردها پاک شوند تا دوباره انتقال اطلاعات صورت گیرد

UPDATE dbo.tFacM
SET No = No + 6
WHERE Branch = 3 AND Status = 7
Go

UPDATE dbo.tFacM
SET No = No + 205
WHERE Branch = 5 AND Status = 7
Go

DELETE FROM dbo.tHistory
GO

UPDATE tfacM SET branch = 1
GO


