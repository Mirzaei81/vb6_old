
declare @p7 int
set @p7=396
exec Update_HavalehResid 1,1395,0,1,N'95/01/01',N'95/06/27',@p7 output
select @p7
GO


exec Update_tblTotal_tInventory_tGood_For_FinalPrice N'95/06/27',N'‘‰»Â',N'01:46',N'95/01/01',N'95/06/27',3,1,1395,1
GO

