

--اضافه شدن دسته بندی به آپشن های کالاها
--برای یوگوبری و مشابه
--Script_V26_16_Fix10_4_OptionCategory
--93/10/23


IF COL_LENGTH('tDifferences','CategoryType') IS NULL
BEGIN
	ALTER TABLE dbo.tDifferences
	ADD CategoryType INT NULL
END

GO



