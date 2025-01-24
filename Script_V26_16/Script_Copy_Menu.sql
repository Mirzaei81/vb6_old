
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
		10 ,
          BtnNum ,
          FactorType ,
          NameDisp ,
          LatinNameDisp ,
          PicturePath ,
          Branch
FROM tNameDisp WHERE StationId = 5
GO


INSERT INTO dbo.tGood_Menu
        ( GoodCode ,
          StationId ,
          FactorType ,
          BtnNum ,
          Branch
        )
SELECT 
			GoodCode ,
          10 ,
          FactorType ,
          BtnNum ,
          Branch
FROM dbo.tGood_Menu WHERE StationId = 5
GO


