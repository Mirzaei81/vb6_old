
--Script_V26_16_Fix10_10_ResidMovaghat
--چاپ رسید موقت
--نام ریپورت :  ResidMovaghat.rpt
--93/10/29


DECLARE @MaxPrintFormat INT 
SELECT @MaxPrintFormat = ISNULL(MAX(PrintFormat) ,0) + 1 FROM tPrintFormat

INSERT INTO dbo.tPrintFormat
        ( PrintFormat ,
          PrintFormatName ,
          RptFilePath ,
          NoticeNo ,
          LatinRptFilePath ,
          PrintFormatLatinName ,
          Active
        )
VALUES  ( @MaxPrintFormat , -- PrintFormat - int
          N'رسید موقت' , -- PrintFormatName - nvarchar(50)
          N'A4\ResidMOvaghat.rpt' , -- RptFilePath - nvarchar(50)
          NULL  , -- NoticeNo - int
          N'A4\ResidMOvaghat.rpt' , -- LatinRptFilePath - nvarchar(50)
          N'ResidMOvaghat.rpt' , -- PrintFormatLatinName - nvarchar(50)
          1  -- Active - bit
        )
        
GO





