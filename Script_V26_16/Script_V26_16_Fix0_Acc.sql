/*
Run this script on:

        (local).Total_V26_15    -  This database will be modified

to synchronize it with:

        (local).Total_V26_16

You are recommended to back up your database before running this script

Script created by SQL Compare version 10.4.8 from Red Gate Software Ltd at 2014/01/05 10:56:14 ب.ظ

*/
SET NUMERIC_ROUNDABORT OFF
GO
SET ANSI_PADDING, ANSI_WARNINGS, CONCAT_NULL_YIELDS_NULL, ARITHABORT, QUOTED_IDENTIFIER, ANSI_NULLS ON
GO
IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE id=OBJECT_ID('tempdb..#tmpErrors')) DROP TABLE #tmpErrors
GO
CREATE TABLE #tmpErrors (Error int)
GO
SET XACT_ABORT ON
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO
BEGIN TRANSACTION
GO

PRINT N'Add Refrence_Acc to [dbo].[tFacM]'
GO
ALTER TABLE dbo.tFacM
	ADD Refrence_Acc INT NULL 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO

PRINT N'Creating [dbo].[tblAcc_Moein]'
GO
CREATE TABLE [dbo].[tblAcc_Moein]
(
[KolId] [int] NOT NULL,
[MoeinId] [int] NOT NULL,
[MoeinName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Moein_MoeinName] DEFAULT (''),
[Kind] [tinyint] NOT NULL CONSTRAINT [DF_tblAcc_Moein_Kind] DEFAULT ((0)),
[Active] [bit] NOT NULL CONSTRAINT [DF_tMoein_Active] DEFAULT ((1))
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Moein] on [dbo].[tblAcc_Moein]'
GO
ALTER TABLE [dbo].[tblAcc_Moein] ADD CONSTRAINT [PK_tblAcc_Moein] PRIMARY KEY CLUSTERED  ([KolId], [MoeinId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Moein]'
GO


CREATE PROCEDURE [dbo].[Delete_tblAcc_Moein] (
@KolID int, 		
@MoeinId int
) AS
DELETE [tblAcc_Moein]
WHERE
	[KolID] = @KolID AND 
	[MoeinId] = @MoeinId

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_DocumentDetail]'
GO
CREATE TABLE [dbo].[tblAcc_DocumentDetail]
(
[AccountYear] [smallint] NOT NULL,
[Branch] [int] NOT NULL,
[DocumentId] [int] NOT NULL,
[RowId] [int] NOT NULL,
[KolId] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_KolId] DEFAULT ((0)),
[MoeinId] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_MoeinId] DEFAULT ((0)),
[TafsiliId] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_TafsiliId] DEFAULT ((0)),
[RowDes] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL CONSTRAINT [DF_tblAcc_DocumentDetail_RowDes] DEFAULT (''),
[Bedehkar] [bigint] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_Bedehkar] DEFAULT ((0)),
[Bestankar] [bigint] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_Bestankar] DEFAULT ((0)),
[kind] [tinyint] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_kind] DEFAULT ((0)),
[SaveDate] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_SaveDate] DEFAULT ((0)),
[UserId] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentDetail_UserId] DEFAULT ((0)),
[CheckNo] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[CheckDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[intRefrenceCheque] [int] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_DocumentDetail] on [dbo].[tblAcc_DocumentDetail]'
GO
ALTER TABLE [dbo].[tblAcc_DocumentDetail] ADD CONSTRAINT [PK_tblAcc_DocumentDetail] PRIMARY KEY CLUSTERED  ([AccountYear], [Branch], [DocumentId], [RowId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_DocumentHeader]'
GO
CREATE TABLE [dbo].[tblAcc_DocumentHeader]
(
[AccountYear] [smallint] NOT NULL,
[Branch] [int] NOT NULL,
[DocumentId] [int] NOT NULL,
[DocumentDate] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentHeader_DocumentDate] DEFAULT ((0)),
[DocumentDes] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL CONSTRAINT [DF_tblAcc_DocumentHeader_DocumentDes] DEFAULT (''),
[State] [tinyint] NOT NULL CONSTRAINT [DF_tblAcc_DocumentHeader_State] DEFAULT ((0)),
[DocumentId2] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentHeader_DocumentId2] DEFAULT ((0)),
[DocumentKind] [tinyint] NOT NULL CONSTRAINT [DF_tblAcc_DocumentHeader_DocumentKind] DEFAULT ((1)),
[SaveDate] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentHeader_SaveDate] DEFAULT ((0)),
[UserId] [int] NOT NULL CONSTRAINT [DF_tblAcc_DocumentHeader_UserId] DEFAULT ((0)),
[Refrence_Sale] [int] NULL,
[Refrence_Khazane] [int] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_DocumentHeader] on [dbo].[tblAcc_DocumentHeader]'
GO
ALTER TABLE [dbo].[tblAcc_DocumentHeader] ADD CONSTRAINT [PK_tblAcc_DocumentHeader] PRIMARY KEY CLUSTERED  ([AccountYear], [Branch], [DocumentId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[ConvIntToDateFormat]'
GO


CREATE FUNCTION [dbo].[ConvIntToDateFormat] (@dt int)  
RETURNS nvarchar(10) AS  
BEGIN 
 declare @r varchar(10)
 declare @y int
 declare @m int
 declare @d int

 if (@dt > 0)
 begin
     set @d=@dt % 100 + 100
     set @dt=@dt / 100
     
     set @m=@dt % 100 + 100
     set @dt=@dt / 100
     
     set @y=@dt
     
     set @r = right(str(@y),4)+'/'+right(str(@m),2)+'/'+right(str(@d),2)
 end
 else
 begin
     set @r=''
 end
 return @r
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_AsnadTarazNashodeh]'
GO
CREATE PROCEDURE [dbo].[Get_All_AsnadTarazNashodeh](@AccountYear smallint, @Branch int) AS
SELECT     tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdt, 
                      tblAcc_DocumentHeader.DocumentDes, t.sBedehkar, t.sBestankar, 
			CASE WHEN t.sBedehkar >= t.sBestankar THEN t.sBedehkar - t.sBestankar ELSE t.sBestankar - t.sBedehkar END AS diff
FROM         tblAcc_DocumentHeader INNER JOIN
                          (SELECT     AccountYear, Branch, DocumentId, SUM(Bedehkar) AS sBedehkar, SUM(Bestankar) AS sBestankar
                             FROM         tblAcc_DocumentDetail
                             GROUP BY AccountYear, Branch, DocumentId
                             HAVING      SUM(Bedehkar) <> SUM(Bestankar)) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND 
                      tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)
ORDER BY tblAcc_DocumentHeader.DocumentId
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Moein_Atf]'
GO
CREATE TABLE [dbo].[tblAcc_Moein_Atf]
(
[KolId] [int] NOT NULL,
[MoeinId] [int] NOT NULL,
[AtfId] [int] NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Moein_Atf] on [dbo].[tblAcc_Moein_Atf]'
GO
ALTER TABLE [dbo].[tblAcc_Moein_Atf] ADD CONSTRAINT [PK_tblAcc_Moein_Atf] PRIMARY KEY CLUSTERED  ([KolId], [MoeinId], [AtfId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atf_ByID]'
GO


-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atf_ByID] (
		 	
				
		@KolID int, 		
		@MoeinId int, 		
		@AtfID int

		) AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[AtfID]
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[KolID] = @KolID AND 
			[MoeinId] = @MoeinId AND 
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Tafsili]'
GO
CREATE TABLE [dbo].[tblAcc_Tafsili]
(
[Branch] [int] NOT NULL,
[TafsiliId] [int] NOT NULL,
[TafsiliName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Tafsili_TafsiliName] DEFAULT (''),
[Active] [bit] NOT NULL CONSTRAINT [DF_tTafsili_Active] DEFAULT ((1)),
[AccountYear] [smallint] NULL,
[RemainingAmount] [int] NULL,
[SanadNo] [int] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Tafsili] on [dbo].[tblAcc_Tafsili]'
GO
ALTER TABLE [dbo].[tblAcc_Tafsili] ADD CONSTRAINT [PK_tblAcc_Tafsili] PRIMARY KEY CLUSTERED  ([Branch], [TafsiliId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_CountTafsiliById]'
GO

CREATE proc [dbo].[Get_CountTafsiliById]
(
@TafsiliId int	,
@Result int OUT	
)
as
begin

select @Result=(select Count(*) from tblAcc_Tafsili where TafsiliId=@TafsiliId)


return @Result

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Kol]'
GO
CREATE TABLE [dbo].[tblAcc_Kol]
(
[KolId] [int] NOT NULL,
[GroupId] [int] NOT NULL CONSTRAINT [DF_tblAcc_Kol_GroupId] DEFAULT ((0)),
[KolName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Kol_KolName] DEFAULT (''),
[Active] [bit] NOT NULL CONSTRAINT [DF_tKol_Active] DEFAULT ((1)),
[ShenaseId] [int] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Kol] on [dbo].[tblAcc_Kol]'
GO
ALTER TABLE [dbo].[tblAcc_Kol] ADD CONSTRAINT [PK_tblAcc_Kol] PRIMARY KEY CLUSTERED  ([KolId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_KolByName]'
GO
CREATE PROC [dbo].[Get_KolByName](@Search NVARCHAR(20))
AS
BEGIN
	SELECT * 
	FROM tblAcc_Kol 
	WHERE CHARINDEX(@Search,dbo.tblAcc_Kol.KolName)>0
END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Rooznameh]'
GO

CREATE PROCEDURE [dbo].[Get_All_Rooznameh](@AccountYear smallint,
										  @Branch int,
										  @d1 int = 0,
										  @d2 int = 0,
										  @RowDes nvarchar(255)) AS
SELECT     0 AS DocumentId,
		 '       ' AS sdate,
		 @RowDes AS RowDes,
		 0 AS kind,
		 0 AS KolId,
		 0 AS MoeinId,
		 0 AS TafsiliId,
         SUM(tblAcc_DocumentDetail.Bedehkar) AS bd,
		 SUM(tblAcc_DocumentDetail.Bestankar) AS bs
FROM         tblAcc_DocumentHeader 
		INNER JOIN   tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear 
				AND	tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch 
				AND	tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE (State > 1) 
		AND (tblAcc_DocumentHeader.DocumentDate < @d1)
UNION
SELECT   tblAcc_DocumentHeader.DocumentId,
		 dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdate,
		 tblAcc_DocumentDetail.RowDes,
		 tblAcc_DocumentDetail.kind,
		 tblAcc_DocumentDetail.KolId,
		 tblAcc_DocumentDetail.MoeinId,
		 tblAcc_DocumentDetail.TafsiliId,
         tblAcc_DocumentDetail.Bedehkar AS bd,
		 tblAcc_DocumentDetail.Bestankar AS bs
FROM         tblAcc_DocumentHeader 
		INNER JOIN tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear 
				AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch 
				AND	tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE (State > 1) 
		AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)
ORDER BY tblAcc_DocumentHeader.DocumentId,
		 kind,
		 tblAcc_DocumentDetail.KolId,
		 tblAcc_DocumentDetail.MoeinId,
		 tblAcc_DocumentDetail.TafsiliId,
		 bd,
		 bs



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[tAccountYears]'
GO
ALTER TABLE [dbo].[tAccountYears] ADD
[nvcDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[ClosingSanad] [int] NULL,
[OpeningSanad] [int] NULL,
[FirstMojodi] [int] NULL
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_MoeinByName]'
GO
CREATE PROC [dbo].[Get_MoeinByName](@Search NVARCHAR(20),
							@KolId INT)
AS
BEGIN
	SELECT * 
	FROM [dbo].[tblAcc_Moein] 
	WHERE CHARINDEX(@Search,dbo.tblAcc_Moein.MoeinName)>0
			AND dbo.tblAcc_Moein.KolId=@KolId
END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_AtfID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_AtfID] (
			
			
			@AtfID int
				
		) AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[AtfID]
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_CountKolById]'
GO
CREATE proc [dbo].[Get_CountKolById]
(
@KolId int,
@Result int OUT	
)
as
begin

select @Result=(select Count(*) from tblAcc_Kol where KolId=@KolId)


return @Result
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Moein]'
GO
CREATE	PROC [dbo].[Get_All_Moein]
AS
BEGIN
SELECT *,
(SELECT [KolName] FROM [dbo].[tblAcc_Kol] WHERE [KolId] =[dbo].[tblAcc_Moein].[KolId])AS KolName
FROM [dbo].[tblAcc_Moein]
ORDER BY [MoeinId],[MoeinName]



end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_PaymentSanad]'
GO
CREATE TABLE [dbo].[tblAcc_PaymentSanad]
(
[intSerialNo] [int] NOT NULL IDENTITY(1, 1),
[CheckNo] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[DateS] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Price] [bigint] NOT NULL,
[Descs] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[BankAccountTafsili] [int] NULL,
[PaymentTypeId] [int] NOT NULL,
[DateT] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[RecKol] [int] NULL,
[RecMoein] [int] NULL,
[RecTafsili] [int] NULL,
[Taraf] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[PayKol] [int] NULL,
[PayMoein] [int] NULL,
[PayTafsili] [int] NULL,
[PayTafsiliName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Vosouli_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[BargashtiMoshtari_Date] [nvarchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Bargashti_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Sanad_Cash] [int] NULL,
[Sanad_Pardakhti] [int] NULL,
[Sanad_Vosouli] [int] NULL,
[Sanad_BargashtiMoshtari] [int] NULL,
[Sanad_Bargashti] [int] NULL,
[Resid] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Void] [bit] NOT NULL CONSTRAINT [DF_tblAcc_PaymentSanad_Void] DEFAULT ((0)),
[CheckBookId] [int] NULL,
[Void_Date] [nvarchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Cash_Date] [nvarchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_PaymentSanad] on [dbo].[tblAcc_PaymentSanad]'
GO
ALTER TABLE [dbo].[tblAcc_PaymentSanad] ADD CONSTRAINT [PK_tblAcc_PaymentSanad] PRIMARY KEY CLUSTERED  ([intSerialNo])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating index [IX_tblAcc_PaymentSanad] on [dbo].[tblAcc_PaymentSanad]'
GO
CREATE NONCLUSTERED INDEX [IX_tblAcc_PaymentSanad] ON [dbo].[tblAcc_PaymentSanad] ([CheckNo])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_VosouliPayChequeByChequeNo]'
GO
Create PROC [dbo].[Get_VosouliPayChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_PaymentSanad]
WHERE [CheckNo]=@ChequeNo AND [PaymentTypeId]=2
END 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_ByID] (
		 	
				
		@KolID int, 		
		@MoeinId int

		) AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[MoeinName],
				[Kind],
				[Active]
		
		FROM 
		
		[tblAcc_Moein]
		
		WHERE
		
		
			[KolID] = @KolID AND 
			[MoeinId] = @MoeinId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Moein_Atf]'
GO

CREATE PROCEDURE [dbo].[Insert_tblAcc_Moein_Atf] (
				
		@KolID int, 		
		@MoeinId int, 		
		@AtfID int
	) 
	
	AS
		
	INSERT INTO [tblAcc_Moein_Atf]
		
	(
		[KolID],
		[MoeinId],
		[AtfID]
	)		
		
	VALUES		
	(
		@KolID,
		@MoeinId,
		@AtfID
	)




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_CountMoeinById]'
GO
CREATE proc [dbo].[Get_CountMoeinById]
(
@MoeinId int	,
@Result int OUT
)
as
begin
select @Result=(select Count(*) from tblAcc_Moein where MoeinId=@MoeinId)


return @Result

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Kols_ByFK_GroupID_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Kols_ByFK_GroupID_Count] (
			
			
			@GroupID int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Kol]
		
		WHERE
		
		
			[GroupID] = @GroupID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_KartHesab]'
GO



CREATE PROCEDURE [dbo].[Get_All_KartHesab](@AccountYear smallint, @Branch int, @KolId int, @MoeinId int, @TafsiliId int, @d1 int, @d2 int, @title nvarchar(255)) AS
SELECT TOP 100 PERCENT * FROM (
SELECT     DocumentId, dbo.ConvIntToDateFormat(MAX(dt)) AS sdate, SUM(Bedehkar) AS Bedehkar, SUM(Bestankar) AS Bestankar, MAX(RowDes) AS RowDes, 
                      KolId, MoeinId, TafsiliId, MAX(DocumentDate) AS DocumentDate, MAX(t1) AS t1, MAX(t2) AS t2, kind
FROM         (SELECT     0 AS RowId, 0 AS DocumentId, 0 AS dt, SUM(tblAcc_DocumentDetail.Bedehkar) AS Bedehkar, SUM(tblAcc_DocumentDetail.Bestankar) AS Bestankar, 
                                              @title AS RowDes, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS DocumentDate, MAX(tblAcc_Kol.KolName) + ' - ' + MAX(tblAcc_Moein.MoeinName) AS t1, 
                                              MAX(tblAcc_Tafsili.TafsiliName) AS t2, 0 AS kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < @d1)
                        GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, RowId
                        UNION ALL
                        SELECT     RowId, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate AS dt, tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar, tblAcc_DocumentDetail.RowDes, 
                                              dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, tblAcc_DocumentHeader.DocumentDate, tblAcc_Kol.KolName + ' - ' + tblAcc_Moein.MoeinName AS t1, tblAcc_tafsili.TafsiliName AS t2, 
                                              dbo.tblAcc_DocumentDetail.kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) t
WHERE ((KolId = @KolId) OR (@KolId = 0)) AND ((MoeinId = @MoeinId) OR (@MoeinId = 0)) AND ((TafsiliId = @TafsiliId) OR (@TafsiliId = 0))
GROUP BY DocumentId, RowId, KolId, MoeinId, TafsiliId, Kind) dt
ORDER BY KolId, MoeinId, TafsiliId, DocumentDate, DocumentId, kind


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentDetails_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentDetails_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AccountYear smallint, Branch int, DocumentId int, RowId int, KolId int, MoeinId int, TafsiliId int, RowDes nvarchar(100), Bedehkar int, Bestankar int, kind tinyint, SaveDate int, UserId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		
	FROM [tblAcc_DocumentDetail] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Moein_Atf]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Update_tblAcc_Moein_Atf] (
				
				
		@KolID int, 		
		@MoeinId int, 		
		@AtfID int

		
		) AS
		
		UPDATE [tblAcc_Moein_Atf]
		
		SET
		
		
				[KolID] = @KolID,
				[MoeinId] = @MoeinId,
				[AtfID] = @AtfID

		
		WHERE
		
		
		
			[KolID] = @KolID AND 
			[MoeinId] = @MoeinId AND 
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Tafsili_Atf]'
GO
CREATE TABLE [dbo].[tblAcc_Tafsili_Atf]
(
[Branch] [int] NOT NULL,
[TafsiliId] [int] NOT NULL,
[AtfId] [int] NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Tafsili_Atf] on [dbo].[tblAcc_Tafsili_Atf]'
GO
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] ADD CONSTRAINT [PK_tblAcc_Tafsili_Atf] PRIMARY KEY CLUSTERED  ([Branch], [TafsiliId], [AtfId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_TafsiliByRelatedAtf]'
GO
Create Proc [dbo].[Insert_TafsiliByRelatedAtf]
(
@Branch int,
@TafsiliId int,
@TafsiliName NVarChar(50),
@Active bit,
@AtfId int

)
As
Begin
INSERT INTO [tblAcc_Tafsili]
		
	(
		[Branch],
		[TafsiliId],
		[TafsiliName],
		[Active]
	)		
		
	VALUES		
	(
		@Branch,
		@TafsiliId,
		@TafsiliName,
		@Active
	)

DECLARE @count INT
SET @count=(SELECT  COUNT(*) FROM [dbo].[tblAcc_Tafsili_Atf] WHERE [TafsiliId]=@TafsiliId AND [AtfId]=@AtfId)

IF (@count=0 )
INSERT INTO [dbo].[tblAcc_Tafsili_Atf]
        ( [Branch], [TafsiliId], [AtfId] )
VALUES  ( @Branch, -- Branch - int
          @TafsiliId, -- TafsiliId - int
          @AtfId  -- AtfId - int
          )

End




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Atf]'
GO
CREATE TABLE [dbo].[tblAcc_Atf]
(
[AtfId] [int] NOT NULL,
[AtfName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Atf_AtfName] DEFAULT (''),
[Active] [bit] NOT NULL CONSTRAINT [DF_tAtf_Active] DEFAULT ((1))
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Atf] on [dbo].[tblAcc_Atf]'
GO
ALTER TABLE [dbo].[tblAcc_Atf] ADD CONSTRAINT [PK_tblAcc_Atf] PRIMARY KEY CLUSTERED  ([AtfId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_AtfMoein]'
GO
CREATE PROC [dbo].[Get_All_AtfMoein]
(
@KolId int
)
as
begin
SELECT *,
(SELECT COUNT(*) FROM [dbo].[tblAcc_Moein_Atf] WHERE [AtfId] =[dbo].[tblAcc_Atf].[AtfId] AND [MoeinId]=[dbo].[tblAcc_Moein].[MoeinId] )AS Relation
FROM [dbo].[tblAcc_Moein]
CROSS JOIN [dbo].[tblAcc_Atf]
WHERE [dbo].[tblAcc_Moein].[KolId]=@KolId
ORDER BY [dbo].[tblAcc_Moein].[MoeinId]

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[fnDocumentState]'
GO

CREATE FUNCTION [dbo].[fnDocumentState] (@a tinyint)  
RETURNS nvarchar(25) AS  
BEGIN 
	declare @d nvarchar(25)
	set @d = case @a
		when 1 then N'يادداشت'
		when 2 then N'ثبت موقت'
		when 3 then N'ثبت قطعي'
		else N''
	end
	return @d
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_KholasehAsnad]'
GO

CREATE PROCEDURE [dbo].[Get_All_KholasehAsnad](@AccountYear smallint, @Branch int, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0, @DocumentId21 int, @DocumentId22 int, @State tinyint, @DocumentKind tinyint, @b1 int, @b2 int) AS
SELECT     CASE WHEN State = 3 THEN tblAcc_DocumentHeader.DocumentId ELSE 0 END AS DocumentId, tblAcc_DocumentHeader.DocumentDes, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sDocumentDate, 
                      tblAcc_DocumentHeader.DocumentId2, dbo.fnDocumentState(State) AS StatusTitle,
                      sdBedehkar, CASE WHEN t.ct IS NOT NULL THEN t.ct ELSE 0 END AS ct
FROM         tblAcc_DocumentHeader LEFT OUTER JOIN
                      (SELECT AccountYear, Branch, DocumentId, COUNT(DocumentId) AS ct, CASE WHEN SUM(Bedehkar) IS NOT NULL THEN SUM(Bedehkar) ELSE 0 END AS sdBedehkar FROM tblAcc_DocumentDetail GROUP BY AccountYear, Branch, DocumentId) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
WHERE (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND ((@d2 = 0) OR (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) AND ((@DocumentId2 = 0) OR (tblAcc_DocumentHeader.DocumentId BETWEEN @DocumentId1 AND @DocumentId2))
	AND ((@State = 0) OR (tblAcc_DocumentHeader.State = @State)) AND ((@DocumentKind = 0) OR (tblAcc_DocumentHeader.DocumentKind = @DocumentKind)) AND ((@DocumentId22 = 0) OR (tblAcc_DocumentHeader.DocumentId2 BETWEEN @DocumentId21 AND @DocumentId22)) AND ((@b2 = 0) OR (t.sdBedehkar BETWEEN @b1 AND @b2))
ORDER BY tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_ByID_Count]'
GO




CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_ByID_Count] (
		 	
				
		@KolID int, 
		@MoeinID int

		) AS
		
		SELECT 
		
		
				COUNT([MoeinID]) AS ct
		
		FROM 
		
		[tblAcc_Moein]
		
		WHERE
		
		
			[KolID] = @KolID AND [MoeinID] = @MoeinID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moein_Atfs]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moein_Atfs]
				
		AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[AtfID]
		
		FROM 
		
		[tblAcc_Moein_Atf]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Kols_ByPK_KolID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Kols_ByPK_KolID] (
			
			
			@KolID int
				
		) AS
		
		SELECT 
		
		
				[KolID],
				[GroupID],
				[KolName],
				[Active]
		
		FROM 
		
		[tblAcc_Kol]
		
		WHERE
		
		
			[KolID] = @KolID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_PayType]'
GO
CREATE TABLE [dbo].[tblAcc_PayType]
(
[PaymentTypeId] [int] NOT NULL,
[PaymentTypeName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_PayType] on [dbo].[tblAcc_PayType]'
GO
ALTER TABLE [dbo].[tblAcc_PayType] ADD CONSTRAINT [PK_tblAcc_PayType] PRIMARY KEY CLUSTERED  ([PaymentTypeId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Cheque_Payment_ByDate]'
GO


--براي ريپورت هست 
CREATE  PROCEDURE [dbo].[Get_All_Cheque_Payment_ByDate] 
(
@SystemDate NVARCHAR(10) ,
@SystemDay AS NVARCHAR(20) ,
@SystemTime NVARCHAR(5) ,
@PaymentTypeId TINYINT ,
@BankTafsili INT ,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) ,
@OrderView INT ,
@AscDesc INT ,
@ChequeType NVARCHAR(20) ,
@AccountDesc NVARCHAR(50) ,
@OrderDesc NVARCHAR(20) ,
@SortDesc NVARCHAR(20)
)
AS

		
IF @OrderView = 0 OR @OrderView = 3 OR @OrderView = 4
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		,@FromDate AS FromDate ,@ToDate AS ToDate , @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		, ISNULL(dbo.tblAcc_Tafsili.TafsiliName , N'' ) AS TafsiliName
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId
		LEFT OUTER JOIN dbo.tblAcc_Tafsili ON tblAcc_PaymentSanad.BankAccountTafsili = tblAcc_Tafsili.TafsiliId 
		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId] > 1 AND [tblAcc_PaymentSanad].[PaymentTypeId] < 6 AND [CheckNo] is NOT null			
			AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			WHEN 4 THEN BankAccountTafsili
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		,@FromDate AS FromDate ,@ToDate AS ToDate , @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		, ISNULL(dbo.tblAcc_Tafsili.TafsiliName , N'' ) AS TafsiliName
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId
		LEFT OUTER JOIN dbo.tblAcc_Tafsili ON tblAcc_PaymentSanad.BankAccountTafsili = tblAcc_Tafsili.TafsiliId 

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId] > 1 AND [tblAcc_PaymentSanad].[PaymentTypeId] < 6 AND [CheckNo] is NOT null			
			AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			WHEN 4 THEN BankAccountTafsili
			END DESC 
		END 
END 

ELSE IF @OrderView = 1 OR @OrderView = 2 OR @OrderView = 5

BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		,@FromDate AS FromDate ,@ToDate AS ToDate , @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		, ISNULL(dbo.tblAcc_Tafsili.TafsiliName , N'' ) AS TafsiliName
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId
		LEFT OUTER JOIN dbo.tblAcc_Tafsili ON tblAcc_PaymentSanad.BankAccountTafsili = tblAcc_Tafsili.TafsiliId 

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId] > 1 AND [tblAcc_PaymentSanad].[PaymentTypeId] < 6 AND [CheckNo] is NOT null			
			AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 5 THEN Taraf
			END ASC  
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		,@FromDate AS FromDate ,@ToDate AS ToDate , @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		, ISNULL(dbo.tblAcc_Tafsili.TafsiliName , N'' ) AS TafsiliName
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId
		LEFT OUTER JOIN dbo.tblAcc_Tafsili ON tblAcc_PaymentSanad.BankAccountTafsili = tblAcc_Tafsili.TafsiliId

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId] > 1 AND [tblAcc_PaymentSanad].[PaymentTypeId] < 6 AND [CheckNo] is NOT null			
			AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 5 THEN Taraf
			END DESC 
		END 
END 





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliByKolMoeinAtfFull]'
GO




Create PROC [dbo].[Get_TafsiliByKolMoeinAtfFull]( @AtfId INT )
AS	
    BEGIN
        SELECT DISTINCT
                dbo.tblAcc_Tafsili.TafsiliName ,
                dbo.tblAcc_Tafsili.TafsiliId
		FROM    dbo.tblAcc_Tafsili
				JOIN dbo.tblAcc_Tafsili_Atf ON dbo.tblAcc_Tafsili.TafsiliId = dbo.tblAcc_Tafsili_Atf.TafsiliId 
		WHERE dbo.tblAcc_Tafsili_Atf.AtfId=@AtfId AND [dbo].[tblAcc_Tafsili].[TafsiliId]<>0

    END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_PayCash]'
GO
CREATE PROC [dbo].[Get_All_PayCash]
(
@PayType INT,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) 
)
AS
BEGIN
SELECT [intSerialNo],
		[DateT],
		[Sanad_Cash],
		[PayTafsili],
		[Taraf],
		[PayTafsiliName],
		[RecTafsili],
		[Resid],
		[Price]
 FROM [dbo].[tblAcc_PaymentSanad]
WHERE [PaymentTypeId]=@PayType AND [DateT]>=@FromDate AND [DateT]<=@ToDate
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moein_Atfs_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moein_Atfs_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Moein_Atf]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tAccountYears_FirstMojodi]'
GO

CREATE Proc [dbo].[Update_tAccountYears_FirstMojodi]
@SanadNo INT ,
@AccountYear SMALLINT   
AS 
UPDATE dbo.tAccountYears
	SET FirstMojodi = @SanadNo WHERE AccountYear = @AccountYear
	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[shamsi]'
GO
----------------------------------------------------------

ALTER  FUNCTION [dbo].[Shamsi](
@sDate  DATETIME) RETURNS CHAR(8)
AS  

BEGIN
	DECLARE @Date VARCHAR(10)
	declare @SEPARATOR AS CHAR(1) 
	DECLARE @year CHAR(4)
	DECLARE @month CHAR(2)
	DECLARE @day CHAR(2)

	IF dbo.MiladiDate() = 0 
	BEGIN 

		DECLARE @Result CHAR(10)
		DECLARE @IYear INT 
		DECLARE @IMonth INT 
		DECLARE @IDay INT 
		DECLARE @OYear INT 
		DECLARE @OMonth INT 
		DECLARE @ODay INT 
	    
		
		SET @year=CAST(datepart(yyyy,@sDate)AS CHAR(4))
		SET @month=CAST(datepart(mm,@sDate)AS CHAR(2))
		SET @day=CAST(datepart(dd,@sDate)AS CHAR(2))
		
		SET @month =REPLACE(SPACE(2 - LEN(@month)), ' ', '0') + @month
		SET @day =REPLACE(SPACE(2 - LEN(@day)), ' ', '0') + @day
		SET @Date=@year+'/'+@month +'/'+@day

		SET @SEPARATOR= '/'
		-- read date
		DECLARE @temp VARCHAR(10)
		DECLARE @i INT
		DECLARE @j INT
	    
		SET @i = CHARINDEX(@SEPARATOR, @Date)
		IF @i > 1
		BEGIN
			SET @temp = LEFT(@Date, @i - 1)
			IF ISNUMERIC(@temp) = 1
				SET @IYear = CAST(@temp AS INT)
			ELSE
				SET @IYear = 0
		END
		ELSE
			SET @IYear = 0
	        
		SET @j = CHARINDEX(@SEPARATOR, @Date, @i + 1)
		IF @j > 0
		BEGIN
			SET @temp = SUBSTRING(@Date,@i + 1,@j - @i - 1)
			IF ISNUMERIC(@temp) = 1
				SET @IMonth = CAST(@temp AS INT)
			ELSE
				SET @IMonth = 0
	        
			IF @j < LEN(@Date)
			BEGIN
				SET @temp = RIGHT(@Date,LEN(@Date) - @j)
				IF ISNUMERIC(@temp) = 1
					SET @IDay = CAST(@temp AS INT)
				ELSE
					SET @IDay = 0
			END
			ELSE
				SET @IDay = 0
	        
			IF @IMonth <= 0 SET @IMonth = 1
			IF @IMonth > 12 SET @IMonth = 12
	        
			IF @IDay <= 0 SET @IDay = 1
			IF @IDay > 31 SET @IDay = 31
		END
		ELSE
		BEGIN
			SET @IMonth = 0
			SET @IDay = 0
		END
	    
		IF @IYear = 0 AND @IMonth = 0 AND @IDay = 0 
			SET @Result = NULL
		ELSE
		BEGIN
			-- civil_persian
			DECLARE @jdn INT
			DECLARE @ISO_8601 AS TINYINT
			DECLARE @Gregorian AS TINYINT
	        
			SET @ISO_8601 = 1
			SET @Gregorian = @ISO_8601
			SET @jdn = dbo.civil_jdn(@IYear,@IMonth,@IDay,@Gregorian)
	        
			-- jdn_persian
			DECLARE @depoch AS INT
			DECLARE @cycle AS INT
			DECLARE @cyear AS INT
			DECLARE @ycycle AS INT
			DECLARE @aux1 AS INT
			DECLARE @aux2 AS INT
			DECLARE @yday AS INT
	        
			SET @depoch = @jdn - dbo.persian_jdn(475, 1, 1)
			SET @cycle = dbo.Fix(@depoch / CAST(1029983 AS REAL))
			SET @cyear = @depoch % 1029983
			IF @cyear = 1029982
				SET @ycycle = 2820
			ELSE
			BEGIN
				SET @aux1 = dbo.Fix(@cyear / CAST(366 AS REAL))
				SET @aux2 = @cyear % 366
				SET @ycycle = FLOOR(((2134 * @aux1) + (2816 * @aux2) + 2815) / CAST(1028522 AS REAL)) + @aux1 + 1
			END
	        
			SET @OYear = @ycycle + (2820 * @cycle) + 474
			IF @OYear <= 0 
				SET @OYear = @OYear - 1
	        
			SET @yday = (@jdn - dbo.persian_jdn(@OYear, 1, 1)) + 1
			IF @yday <= 186 
				SET @OMonth = dbo.Ceil(@yday / CAST(31 AS REAL))
			ELSE
				SET @OMonth = dbo.Ceil((@yday - 6) / CAST(30 AS REAL))
	        
			SET @ODay = (@jdn - dbo.persian_jdn(@OYear, @OMonth, 1)) + 1
	        
	--        SET @Result =    RIGHT('0000'    + CAST(@OYear    AS VARCHAR(10)),4) + '/' + 
	--                        RIGHT('0'        + CAST(@OMonth    AS VARCHAR(10)),2) + '/' + 
	--                        RIGHT('0'        + CAST(@ODay    AS VARCHAR(10)),2)

			SET @Result =    RIGHT('00'    + CAST(@OYear    AS VARCHAR(10)),2) + '/' + 
							RIGHT('0'        + CAST(@OMonth    AS VARCHAR(10)),2) + '/' + 
							RIGHT('0'        + CAST(@ODay    AS VARCHAR(10)),2)
		END
    
	END 

	ELSE
	
	begin

		SET @Year = RIGHT(CAST(YEAR(@sDate) AS NVARCHAR(4)) , 2)
		--SET @Year = RIGHT(@Year ,2)
		
		SET @Month = CAST(MONTH(@sDate) AS NVARCHAR(2)) 
		SET @Day =  CAST(DAY(@sDate) AS NVARCHAR(2))
		
		IF LEN(@Month) = 1  SET @Month = '0' + @Month
		IF LEN(@Day) = 1    SET @Day = '0' + @Day
		
		SET @Result = @Year + '/' +  @Month + '/' + @Day 


	END

    RETURN @Result


END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentGroupByKolIdAndMoeinIdAndTafsiliId]'
GO

CREATE  PROCEDURE [dbo].[Get_All_DocumentGroupByKolIdAndMoeinIdAndTafsiliId](@AccountYear smallint, @Branch int, @DocumentId int) AS
SELECT     TOP 100 PERCENT tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.Bedehkar + tblAcc_DocumentDetail.Bestankar AS detl,
                      tblAcc_DocumentDetail.Bedehkar AS bd, tblAcc_DocumentDetail.Bestankar AS bs, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName AS MoeinTtl, 
                      tblAcc_Tafsili.TafsiliId, -- ' - ' + 
                       tblAcc_Tafsili.TafsiliName  + N' - ' + tblAcc_DocumentDetail.RowDes AS TafsiliName
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId  and tblAcc_DocumentHeader.Branch =tblAcc_DocumentDetail.Branch And     tblAcc_DocumentHeader.AccountYear =tblAcc_DocumentDetail.AccountYear  INNER JOIN
                      tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId LEFT OUTER JOIN
                      tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
WHERE (dbo.tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (dbo.tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentId = @DocumentId)
ORDER BY tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moeins_ByFK_KolID_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moeins_ByFK_KolID_Count] (
			
			
			@KolID int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Moein]
		
		WHERE
		
		
			[KolID] = @KolID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moein_Atfs_Count_ForAll]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moein_Atfs_Count_ForAll](@KolId int, @MoeinId int, @AtfId int)
				
		AS
		
		SELECT
		
			COUNT(*) AS ct
		
		FROM
		
			[tblAcc_Moein_Atf]
		
		WHERE
		
			[KolId] = @KolId AND [MoeinId] = @MoeinId AND [AtfId] = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TarazSoodZian_Rep]'
GO

CREATE PROCEDURE [dbo].[Get_TarazSoodZian_Rep]
    (
      @SystemDate NVARCHAR(10) ,
      @SystemDay NVARCHAR(10) ,
      @SystemTime NVARCHAR(5) ,
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT ,
      @MojodiPrice BIGINT 
    )
AS 


DECLARE @TotalSellAmount BIGINT
DECLARE @TotalSellReturnAmount BIGINT
DECLARE @TotalBuyAmount BIGINT
DECLARE @TotalBuyReturnAmount BIGINT
DECLARE @TotalSaleDiscount BIGINT
DECLARE @TotalBuyDiscount BIGINT

DECLARE @TotalCareeFee BIGINT
DECLARE @TotalPacking BIGINT

DECLARE @TotalLosses BIGINT
DECLARE @TotalHoghough BIGINT
DECLARE @TotalHazine BIGINT

		Select @TotalSellAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter 
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )

		Select @TotalSellReturnAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17)

		Select @TotalBuyAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16)

		Select @TotalBuyReturnAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18)

		Select @TotalSaleDiscount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2)

		Select @TotalBuyDiscount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29)

		Select @TotalLosses = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail 
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32)

		Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail 
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)

		Select @TotalHazine = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14)

		Select @TotalCareeFee = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4)

		Select @TotalPacking = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3)

SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
	   ISNULL(@TotalSellAmount , 0) AS TotalSellAmount ,
       ISNULL(@TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
       ISNULL(@TotalBuyAmount , 0) AS TotalBuyAmount ,
       ISNULL(@TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
       ISNULL(@TotalSaleDiscount , 0) AS TotalSaleDiscount ,
       ISNULL(@TotalBuyDiscount , 0) AS TotalBuyDiscount ,
       ISNULL(@TotalCareeFee , 0) AS TotalCareeFee , 
       ISNULL(@TotalPacking , 0) AS TotalPacking ,
       ISNULL(@TotalLosses , 0) AS TotalLosses ,
       ISNULL(@TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(@TotalHazine , 0) AS TotalHazine ,
       ISNULL(@MojodiPrice , 0) AS MojodiPrice

--===============================================

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Split_Acc]'
GO
SET ANSI_NULLS OFF
GO

CREATE Function [dbo].[Split_Acc](@nvcMainString nText)
RETURNS @ReturnTable TABLE(
	[AccountYear] smallint NOT NULL,
	[Branch] int NOT NULL,
	[DocumentId] int NOT NULL,
	[RowId] int NOT NULL,
	[KolId] int NOT NULL,
	[MoeinId] int NOT NULL,
	[TafsiliId] int NOT NULL,
	[RowDes] [nvarchar] (100) COLLATE Arabic_CI_AS NOT NULL,
	[Bedehkar] int NOT NULL,
	[Bestankar] int NOT NULL,
	[kind] tinyint NOT NULL,
	[SaveDate] int NOT NULL,
	[UserId] int NOT NULL ,
	CheckNo NVARCHAR(20) NULL ,
	Checkdate NVARCHAR(10) NULL 
)
AS
BEGIN
IF @nvcMainString IS NOT NULL
BEGIN
	DECLARE @AccountYear smallint
	DECLARE @Branch int
	DECLARE @DocumentId int
	DECLARE @RowId int
	DECLARE @KolId int
	DECLARE @MoeinId int
	DECLARE @TafsiliId int
	DECLARE @RowDes nvarchar(100)
	DECLARE @Bedehkar int
	DECLARE @Bestankar int
	DECLARE @kind tinyint
	DECLARE @SaveDate int
	DECLARE @UserId int
	DECLARE @CheckNo nvarchar(20)
	DECLARE @CheckDate nvarchar(10)

	DECLARE @TempTable Table (nvcMainString nText)
	DECLARE @intDelimiterPosField  int
	DECLARE @intDelimiterPosRecord int

	INSERT INTO @TempTable values (@nvcMainString)

	SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
	SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)

	WHILE @intDelimiterPosRecord <> 0
	BEGIN
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @AccountYear = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 4, @intDelimiterPosField - 4))) AS smallint) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @Branch = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @DocumentId = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @RowId = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @KolId = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @MoeinId = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @TafsiliId = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @RowDes = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS nvarchar(100)) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @Bedehkar = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @Bestankar = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @kind = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS tinyint) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @SaveDate = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @UserId = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS int) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)

		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @CheckNo = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS NVARCHAR(20)) FROM @TempTable)
		SET @CheckNo=(CASE WHEN @CheckNo=N'' THEN NULL ELSE @CheckNo END)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)

		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField) FROM @TempTable)
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)
		SET @CheckDate = (SELECT CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 3, @intDelimiterPosField - 3))) AS NVARCHAR(10)) FROM @TempTable)
		SET @CheckDate=(CASE WHEN @CheckDate=N'' THEN NULL ELSE @CheckDate END)		
		SET @intDelimiterPosField = ( select PatIndex('%/^/%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select PatIndex('%/$/%' , nvcMainString) from @TempTable)

		UPDATE @TempTable SET nvcMainString = (Select SUBSTRING(nvcMainString, @intDelimiterPosRecord, DataLength(nvcMainString)) FROM @TempTable)
		INSERT INTO @ReturnTable
		                      (AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId , CheckNo , CheckDate)
		VALUES     (@AccountYear, @Branch, @DocumentId, @RowId, @KolId, @MoeinId, @TafsiliId, @RowDes, @Bedehkar, @Bestankar, @kind, @SaveDate, @UserId , @CheckNo , @CheckDate)
	END --WHILE
END
RETURN
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[ShamsiInt]'
GO
SET ANSI_NULLS ON
GO
CREATE FUNCTION [dbo].[ShamsiInt] (@sDate  Datetime)  
RETURNS int AS  

BEGIN
	DECLARE @Date VARCHAR(10)
	DECLARE @Result CHAR(10)
	DECLARE @IYear INT 
	DECLARE @IMonth INT 
	DECLARE @IDay INT 
	DECLARE @OYear INT 
	DECLARE @OMonth INT 
	DECLARE @ODay INT 
	declare @SEPARATOR AS CHAR(1) 
	DECLARE @year CHAR(4)
	DECLARE @month CHAR(2)
	DECLARE @day CHAR(2)
	
	SET @year=CAST(datepart(yyyy,@sDate)AS CHAR(4))
	SET @month=CAST(datepart(mm,@sDate)AS CHAR(2))
	SET @day=CAST(datepart(dd,@sDate)AS CHAR(2))
	
	SET @month =REPLACE(SPACE(2 - LEN(@month)), ' ', '0') + @month
	SET @day =REPLACE(SPACE(2 - LEN(@day)), ' ', '0') + @day
	SET @Date=@year+'/'+@month +'/'+@day

	SET @SEPARATOR= '/'
	-- read date
	DECLARE @temp VARCHAR(10)
	DECLARE @i INT
	DECLARE @j INT
	
	SET @i = CHARINDEX(@SEPARATOR, @Date)
	IF @i > 1
	BEGIN
		SET @temp = LEFT(@Date, @i - 1)
		IF ISNUMERIC(@temp) = 1
			SET @IYear = CAST(@temp AS INT)
		ELSE
			SET @IYear = 0
	END
	ELSE
		SET @IYear = 0
		
	SET @j = CHARINDEX(@SEPARATOR, @Date, @i + 1)
	IF @j > 0
	BEGIN
		SET @temp = SUBSTRING(@Date,@i + 1,@j - @i - 1)
		IF ISNUMERIC(@temp) = 1
			SET @IMonth = CAST(@temp AS INT)
		ELSE
			SET @IMonth = 0
		
		IF @j < LEN(@Date)
		BEGIN
			SET @temp = RIGHT(@Date,LEN(@Date) - @j)
			IF ISNUMERIC(@temp) = 1
				SET @IDay = CAST(@temp AS INT)
			ELSE
				SET @IDay = 0
		END
		ELSE
			SET @IDay = 0
		
		IF @IMonth <= 0 SET @IMonth = 1
		IF @IMonth > 12 SET @IMonth = 12
		
		IF @IDay <= 0 SET @IDay = 1
		IF @IDay > 31 SET @IDay = 31
	END
	ELSE
	BEGIN
		SET @IMonth = 0
		SET @IDay = 0
	END
	
	IF @IYear = 0 AND @IMonth = 0 AND @IDay = 0 
		SET @Result = NULL
	ELSE
	BEGIN
		-- civil_persian
		DECLARE @jdn INT
		DECLARE @ISO_8601 AS TINYINT
		DECLARE @Gregorian AS TINYINT
		
		SET @ISO_8601 = 1
		SET @Gregorian = @ISO_8601
		SET @jdn = dbo.civil_jdn(@IYear,@IMonth,@IDay,@Gregorian)
		
		-- jdn_persian
		DECLARE @depoch AS INT
		DECLARE @cycle AS INT
		DECLARE @cyear AS INT
		DECLARE @ycycle AS INT
		DECLARE @aux1 AS INT
		DECLARE @aux2 AS INT
		DECLARE @yday AS INT
	    
		SET @depoch = @jdn - dbo.persian_jdn(475, 1, 1)
		SET @cycle = dbo.Fix(@depoch / CAST(1029983 AS REAL))
		SET @cyear = @depoch % 1029983
		IF @cyear = 1029982
			SET @ycycle = 2820
		ELSE
		BEGIN
			SET @aux1 = dbo.Fix(@cyear / CAST(366 AS REAL))
			SET @aux2 = @cyear % 366
			SET @ycycle = FLOOR(((2134 * @aux1) + (2816 * @aux2) + 2815) / CAST(1028522 AS REAL)) + @aux1 + 1
		END
	    
		SET @OYear = @ycycle + (2820 * @cycle) + 474
		IF @OYear <= 0 
			SET @OYear = @OYear - 1
	    
		SET @yday = (@jdn - dbo.persian_jdn(@OYear, 1, 1)) + 1
		IF @yday <= 186 
			SET @OMonth = dbo.Ceil(@yday / CAST(31 AS REAL))
		ELSE
			SET @OMonth = dbo.Ceil((@yday - 6) / CAST(30 AS REAL))
	    
		SET @ODay = (@jdn - dbo.persian_jdn(@OYear, @OMonth, 1)) + 1
		
		SET @Result =	RIGHT('0000'	+ CAST(@OYear	AS VARCHAR(10)),4) + '/' + 
						RIGHT('0'		+ CAST(@OMonth	AS VARCHAR(10)),2) + '/' + 
						RIGHT('0'		+ CAST(@ODay	AS VARCHAR(10)),2)
	END
	
	RETURN CAST(REPLACE(@Result,'/','') AS INT)
END
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_DocumentWithDetail]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_DocumentWithDetail](
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@DocumentDate int, 		
		@DocumentDes nvarchar(100), 		
		@State tinyint, 		
		@DocumentId2 int, 		
		@DocumentKind tinyint, 		
		@UserId INT,
		@ds1 NVARCHAR(4000),
		@ds2 NVARCHAR(4000),
		@ds3 NVARCHAR(4000),
		@result INT out
	) 
	
	AS
  
	
	BEGIN	TRAN
		UPDATE  [tblAcc_DocumentHeader]
			
SET 		[DocumentDate] = @DocumentDate,
			[DocumentDes] = @DocumentDes,
			[State] = @State,
			[DocumentId2] = @DocumentId2,
			[DocumentKind] = @DocumentKind,
			[SaveDate] = dbo.ShamsiInt(GetDate()),
			[UserId] = @UserId
WHERE 	[DocumentId] = @DocumentId
		AND [AccountYear] = @AccountYear
		AND [Branch] = @Branch

		IF @@ERROR>0
		BEGIN
			ROLLBACK TRAN
			SET @result=0
			RETURN
		END 

		DELETE  FROM tblAcc_DocumentDetail
		WHERE   AccountYear = @AccountYear
				AND Branch = @Branch
				AND DocumentId = @DocumentId
--DECLARE @Row INT
--SELECT @Row = MAX(RowId) FROM tblAcc_DocumentDetail
--WHERE 	[DocumentId] = @DocumentId
--	AND [AccountYear] = @AccountYear
--	AND [Branch] = @Branch
--SET @Row = ISNULL(@Row , 0)
--Add New Row to tblAcc_DocumentDetail

		IF @ds1<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate
					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId  , --New Row add + @Row
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds1)
			IF @@ERROR>0
			BEGIN
				ROLLBACK TRAN
				SET @result=0
				RETURN
			END 
		END
		IF @ds2<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						(RowId )  ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds2)
			IF @@ERROR>0
			BEGIN
				ROLLBACK TRAN
				SET @result=0
				RETURN
			END 
		END
		IF @ds3<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						(RowId )  ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds3)
			IF @@ERROR>0
			BEGIN
				ROLLBACK TRAN
				SET @result=0
				RETURN
			END 
		END

	COMMIT TRAN
	SET @result= 1
	RETURN
	


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Cancel_Cheque]'
GO
CREATE PROC [dbo].[Cancel_Cheque]
(
@Code AS INT,
@Result INT OUT
)
AS 

UPDATE [dbo].[tblAcc_PaymentSanad] SET [PaymentTypeId]=6,[Void]=1
WHERE [intSerialNo]=@Code
     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result =1
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Get_ShamsiDate_For_Current_Shift]'
GO
SET QUOTED_IDENTIFIER OFF
GO




ALTER   function [dbo].[Get_ShamsiDate_For_Current_Shift](@today dateTime)
Returns  nvarchar(8)

as
begin
declare @shamsi nvarchar(8)
Declare @SubtractDate bit
Set @SubtractDate = 0	

IF dbo.MiladiDate() = 0
begin
	
	SELECT @SubtractDate = 1
	FROM [dbo].[tShift] 
	Where [Code] = dbo.Get_Shift(@today) 
	And dbo.SetTimeFormat(StartTime) > dbo.SetTimeFormat(EndTime)
	And dbo.SetTimeFormat(StartTime) > dbo.SetTimeFormat(@today) 
	And dbo.SetTimeFormat(EndTime) > dbo.SetTimeFormat(@today) 	

	if @SubtractDate = 1
		Set @today = dateadd(Day , -1 ,@today)
	Select @shamsi = dbo.shamsi (@today)

End

ELSE 

BEGIN
DECLARE @Year NVARCHAR(2) 
DECLARE @Month NVARCHAR(2)
DECLARE @Day NVARCHAR(2)


	SELECT @SubtractDate = 1
	FROM [dbo].[tShift] 
	Where [Code] = dbo.Get_Shift(@today) 
	And dbo.SetTimeFormat(StartTime) > dbo.SetTimeFormat(EndTime)
	And dbo.SetTimeFormat(StartTime) > dbo.SetTimeFormat(@today) 
	And dbo.SetTimeFormat(EndTime) > dbo.SetTimeFormat(@today) 	

	if @SubtractDate = 1
		Set @today = dateadd(Day , -1 ,@today)

	SET @Year = RIGHT(CAST(YEAR(@today) AS NVARCHAR(4)) , 2)
	--SET @Year = RIGHT(@Year ,2)
	
	SET @Month = CAST(MONTH(@today) AS NVARCHAR(2)) 
	SET @Day =  CAST(DAY(@today) AS NVARCHAR(2))
	
	IF LEN(@Month) = 1  SET @Month = '0' + @Month
	IF LEN(@Day) = 1    SET @Day = '0' + @Day
	
	SET @shamsi = @Year + '/' +  @Month + '/' + @Day  



End

return @shamsi

END 


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_InvalidKolIdMoeinIdTafsiliId]'
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[Get_All_InvalidKolIdMoeinIdTafsiliId](@KolId int, @MoeinId int, @TafsiliId int) AS
SELECT CASE WHEN SUM(r) IS NOT NULL THEN SUM(r) ELSE 0 END AS rv
FROM (SELECT 10 AS r
	WHERE NOT EXISTS(SELECT KolId FROM tblAcc_Moein WHERE KolId = @KolId AND MoeinId = @MoeinId)
	UNION
	SELECT 1 AS r
	WHERE NOT EXISTS(SELECT TafsiliId FROM tblAcc_Tafsili WHERE TafsiliId = @TafsiliId UNION SELECT 0 AS TafsiliId)) t


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[CheckAtfExistance]'
GO

CREATE PROC [dbo].[CheckAtfExistance](
								@kolId int,
								@MoeinId int,
								@AtfId int
								)
as
BEGIN
	SELECT COUNT(* ) AS [ct]
	FROM dbo.tblAcc_Moein_Atf 
	WHERE AtfId=@AtfId 
			AND MoeinId=@MoeinId 
			AND KolId=@kolId
	
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_Cheque_Payment_Search]'
GO



CREATE  PROCEDURE [dbo].[Get_Cheque_Payment_Search] 
(
@PaymentTypeId int ,
@BankAccountTafsili INT ,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8),
@SortItem INT,
@SortType int,
@CheckNo nvarchar(20)
)
AS

BEGIN 
IF @SortType=0

		SELECT *,
		(SELECT [PaymentTypeName] FROM [dbo].[tblAcc_PayType]  WHERE [dbo].[tblAcc_PayType].[PaymentTypeId]=[dbo].[tblAcc_PaymentSanad].[PaymentTypeId])AS PaymentTypeName

		FROM [dbo].[tblAcc_PaymentSanad]

		WHERE [PaymentTypeId]=(CASE WHEN @PaymentTypeId=0 THEN [PaymentTypeId] ELSE @PaymentTypeId END)
				AND [CheckNo] is NOT null

		AND( (DateS>=@FromDate and DateS<@ToDate AND [PaymentTypeId]<>1) OR [PaymentTypeId]=1)
		ANd BankAccountTafsili=(CASE WHEN @BankAccountTafsili=0 THEN [BankAccountTafsili] ELSE @BankAccountTafsili END)
		And CheckNo=(Case When @CheckNo=N'' Then CheckNo Else @CheckNo End)

		ORDER BY (CASE WHEN @SortItem=0 THEN Replace(DateS,'/','') ELSE CASE WHEN @SortItem=1 THEN Replace(DateT,'/','') ELSE CASE WHEN @SortItem=2 THEN CheckNo ELSE CASE WHEN @SortItem=3 THEN BankAccountTafsili ELSE CASE WHEN @SortItem=4 THEN RecTafsili END END END END END) ASC
else


		SELECT *,
		(SELECT [PaymentTypeName] FROM [dbo].[tblAcc_PayType]  WHERE [dbo].[tblAcc_PayType].[PaymentTypeId]=[dbo].[tblAcc_PaymentSanad].[PaymentTypeId])AS PaymentTypeName

		FROM [dbo].[tblAcc_PaymentSanad]

		WHERE [PaymentTypeId]=(CASE WHEN @PaymentTypeId=0 THEN [PaymentTypeId] ELSE @PaymentTypeId END)
				AND [CheckNo] is NOT null

		AND( (DateS>=@FromDate and DateS<@ToDate AND [PaymentTypeId]<>1) OR [PaymentTypeId]=1)
				ANd BankAccountTafsili=(CASE WHEN @BankAccountTafsili=0 THEN [BankAccountTafsili] ELSE @BankAccountTafsili END)
		And CheckNo=(Case When @CheckNo=N'' Then CheckNo Else @CheckNo End)


		ORDER BY (CASE WHEN @SortItem=0 THEN Replace(DateS,'/','') ELSE CASE WHEN @SortItem=1 THEN Replace(DateT,'/','') ELSE CASE WHEN @SortItem=2 THEN CheckNo ELSE CASE WHEN @SortItem=3 THEN BankAccountTafsili ELSE CASE WHEN @SortItem=4 THEN RecTafsili END END END END END) DESC




END 


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_ByID_Max]'
GO

CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_ByID_Max] 

 	 AS
		
		
	       SELECT TOP 1 dbo.tblAcc_Tafsili.* 

		
		FROM 
		
		     [tblAcc_Tafsili]
		
		WHERE
		
		
			[Branch] = dbo.Get_Current_Branch()
                                
                          order by TafsiliID desc
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentGroupByKolIdAndMoeinId]'
GO



CREATE PROCEDURE [dbo].[Get_All_DocumentGroupByKolIdAndMoeinId](@AccountYear smallint, @Branch int, @DocumentId int) AS
SELECT     TOP 100 PERCENT tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, SUM(tblAcc_DocumentDetail.Bedehkar + tblAcc_DocumentDetail.Bestankar) AS detl, 
                      SUM(tblAcc_DocumentDetail.Bedehkar) AS bd, SUM(tblAcc_DocumentDetail.Bestankar) AS bs, MAX(tblAcc_Kol.KolName) AS des, 
                      MAX(LTRIM(RTRIM(tblAcc_Moein.MoeinName))) AS tblAcc_MoeinTtl
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId  and tblAcc_DocumentHeader.Branch =tblAcc_DocumentDetail.Branch And     tblAcc_DocumentHeader.AccountYear =tblAcc_DocumentDetail.AccountYear INNER JOIN
                      tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId
WHERE (dbo.tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (dbo.tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentId = @DocumentId)
GROUP BY tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId
ORDER BY tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_MoeinBetweenKolID]'
GO



CREATE PROCEDURE [dbo].[Get_All_MoeinBetweenKolID](@KolID1 int, @KolID2 int) AS
SELECT     *
FROM         tblAcc_Moein
WHERE     (KolID BETWEEN @KolID1 AND @KolID2)
ORDER BY KolID, MoeinId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moein_Atfs_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moein_Atfs_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, MoeinId int, AtfID int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, MoeinId, AtfID
		
	FROM [tblAcc_Moein_Atf] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, MoeinId, AtfID
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[GetCashPaymentSearch]'
GO
Create Proc [dbo].[GetCashPaymentSearch]
(
@TafsiliPayer int,
@TafsiliReceiver int,
@PriceFrom bigint,
@PriceTo bigint,
@DateTFrom nvarchar(10),
@DateTTo nvarchar(10),
@Resid nvarchar(255),
@SortItem int,
@SortType int,
@PaymentTypeID int
)
as
begin

Select *,(SELECT PaymentTypeName FROM [dbo].tblAcc_PayType WHERE [dbo].tblAcc_PayType.PaymentTypeId=[dbo].tblAcc_PaymentSanad.PaymentTypeId)AS PaymentTypeName,
(select TafsiliName from tblAcc_Tafsili where TafsiliId=PayTafsili)as PayTafsiliName,
(select TafsiliName from tblAcc_Tafsili where TafsiliId=RecTafsili)as DarTafsiliName
 from tblAcc_PaymentSanad Where 
PayTafsili=(case when @TafsiliPayer=0 then PayTafsili else @TafsiliPayer end)
and
RecTafsili=(case when @TafsiliReceiver=0 then RecTafsili else @TafsiliReceiver end)
and
price>=@PriceFrom
and
price<=@PriceTo
and 
DateT>=@DateTFrom
and
DateT<=@DateTTo
and 
Resid Like '%'+@Resid+'%'
and checkno is  null
and PaymentTypeId= @PaymentTypeID
order by case when @SortItem=0 and @SortType=0 then RecTafsili end asc,
			
			case when @SortItem=0 and @SortType=1 then RecTafsili end desc,


			case when @SortItem=1 and @SortType=0 then PayTafsili end asc,

			case when @SortItem=1 and @SortType=1 then PayTafsili end Desc,


			case when @SortItem=2 and @SortType=0 then DateT end asc,

			case when @SortItem=2 and @SortType=1 then DateT end desc,


			case when @SortItem=3 and @SortType=0 then Price end asc,

			case when @SortItem=3 and @SortType=1 then Price end desc,		

			case when @SortItem=4 and @SortType=0 then Resid end asc,

			case when @SortItem=4 and @SortType=1 then Resid end desc,
			
			case when @SortItem=5 and @SortType=0 then Sanad_Cash end asc,

			case when @SortItem=5 and @SortType=1 then Sanad_Cash end desc
			
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_Current_Branch_Number]'
GO

CREATE PROCEDURE [dbo].[Get_Current_Branch_Number] AS
	SELECT dbo.Get_Current_Branch() AS Branch



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[InsertSanad]'
GO
CREATE PROC [dbo].[InsertSanad]
(
@AccountYear SMALLINT,
@Branch INT,
@DocumentId INT,
@DocumentDate INT,
@DocumentDes nvarchar(255),
@State  TinyInt,
@DocumentId2 int,
@DocumentKind tinyint,
@UserId int,
@ds1 NVARCHAR(4000),
@ds2 NVARCHAR(4000),
@ds3 NVARCHAR(4000),
@ds4 NVARCHAR(4000)
)
as

begin
	INSERT INTO [tblAcc_DocumentHeader]
		
	(
		[AccountYear],
		[Branch],
		[DocumentId],
		[DocumentDate],
		[DocumentDes],
		[State],
		[DocumentId2],
		[DocumentKind],
		[SaveDate],
		[UserId]
	)		
		
	VALUES		
	(
		@AccountYear,
		@Branch,
		@DocumentId,
		@DocumentDate,
		@DocumentDes,
		@State,
		@DocumentId2,
		@DocumentKind,
		dbo.ShamsiInt(GetDate()),
		@UserId
	)

---------------------------------------------------------------------------------
UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/



---------------------------------------------------------------------------------------
DELETE tblAcc_DocumentDetail 
	WHERE AccountYear = @AccountYear 
		AND Branch = @Branch 
		AND DocumentId = @DocumentId
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds1)
IF @ds2<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds2)
IF @ds3<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds3)
IF @ds4<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds4)

-------------------------------------------------------------------------------

UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_AtfIdCountsInChilds]'
GO

CREATE PROCEDURE [dbo].[Get_AtfIdCountsInChilds](@AtfId int) AS
SELECT     SUM(c) AS ct
FROM         (SELECT     COUNT(AtfId) AS c
                        FROM         tblAcc_Moein_Atf
                        WHERE     (AtfId = @AtfId)
                        UNION
                        SELECT     COUNT(AtfId) AS c
                        FROM         tblAcc_Tafsili_Atf
                        WHERE     (AtfId = @AtfId)) t


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Moein_Atf]'
GO

CREATE PROCEDURE [dbo].[Delete_tblAcc_Moein_Atf]
(
  @KolID INT ,
  @MoeinId INT ,
  @AtfID INT=NULL
)
AS 
DELETE  [tblAcc_Moein_Atf]
WHERE   [KolID] = @KolID
        AND [MoeinId] = @MoeinId
        AND [AtfID] =ISNULL(@AtfID,[AtfID])


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_KartHesab_WithRemain]'
GO


CREATE PROCEDURE [dbo].[Get_All_KartHesab_WithRemain]
(@AccountYear smallint, @Branch int, @KolId int, @MoeinId int, @TafsiliId int, @d1 int, @d2 INT , @Remain BIT , @title nvarchar(255)) AS

IF @Remain = 1
	SELECT TOP 100 PERCENT * FROM (
	SELECT     DocumentId, dbo.ConvIntToDateFormat(MAX(dt)) AS sdate, SUM(Bedehkar) AS Bedehkar, SUM(Bestankar) AS Bestankar, MAX(RowDes) AS RowDes, 
						  KolId, MoeinId, TafsiliId, MAX(DocumentDate) AS DocumentDate, MAX(t1) AS t1, MAX(t2) AS t2, kind
	FROM         (SELECT     0 AS RowId, 0 AS DocumentId, 0 AS dt, SUM(tblAcc_DocumentDetail.Bedehkar) AS Bedehkar, SUM(tblAcc_DocumentDetail.Bestankar) AS Bestankar, 
                                              @title AS RowDes, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS DocumentDate, MAX(tblAcc_Kol.KolName) + ' - ' + MAX(tblAcc_Moein.MoeinName) AS t1, 
                                              MAX(tblAcc_Tafsili.TafsiliName) AS t2, 0 AS kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < @d1)
                        GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, RowId
                        UNION ALL
                        SELECT     RowId, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate AS dt, tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar, tblAcc_DocumentDetail.RowDes, 
                                              dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, tblAcc_DocumentHeader.DocumentDate, tblAcc_Kol.KolName + ' - ' + tblAcc_Moein.MoeinName AS t1, tblAcc_tafsili.TafsiliName AS t2, 
                                              dbo.tblAcc_DocumentDetail.kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) t
	WHERE ((KolId = @KolId) OR (@KolId = 0)) AND ((MoeinId = @MoeinId) OR (@MoeinId = 0)) AND ((TafsiliId = @TafsiliId) OR (@TafsiliId = 0))
	GROUP BY DocumentId, RowId, KolId, MoeinId, TafsiliId, Kind) dt
	ORDER BY KolId, MoeinId, TafsiliId, DocumentDate, DocumentId, kind

ELSE

	SELECT TOP 100 PERCENT * FROM (
	SELECT     DocumentId, dbo.ConvIntToDateFormat(MAX(dt)) AS sdate, SUM(Bedehkar) AS Bedehkar, SUM(Bestankar) AS Bestankar, MAX(RowDes) AS RowDes, 
						  KolId, MoeinId, TafsiliId, MAX(DocumentDate) AS DocumentDate, MAX(t1) AS t1, MAX(t2) AS t2, kind
	FROM         (     SELECT     RowId, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate AS dt, tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar, tblAcc_DocumentDetail.RowDes, 
                                              dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, tblAcc_DocumentHeader.DocumentDate, tblAcc_Kol.KolName + ' - ' + tblAcc_Moein.MoeinName AS t1, tblAcc_tafsili.TafsiliName AS t2, 
                                              dbo.tblAcc_DocumentDetail.kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) t
	WHERE ((KolId = @KolId) OR (@KolId = 0)) AND ((MoeinId = @MoeinId) OR (@MoeinId = 0)) AND ((TafsiliId = @TafsiliId) OR (@TafsiliId = 0))
	GROUP BY DocumentId, RowId, KolId, MoeinId, TafsiliId, Kind) dt
	ORDER BY KolId, MoeinId, TafsiliId, DocumentDate, DocumentId, kind


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentGroupByKolId]'
GO



CREATE PROCEDURE [dbo].[Get_All_DocumentGroupByKolId](@AccountYear smallint, @Branch int, @DocumentId int) AS
SELECT     TOP 100 PERCENT tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd, SUM(tblAcc_DocumentDetail.Bestankar) AS bs, 
                      MAX(tblAcc_Kol.KolName) AS des
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId  and tblAcc_DocumentHeader.Branch =tblAcc_DocumentDetail.Branch And     tblAcc_DocumentHeader.AccountYear =tblAcc_DocumentDetail.AccountYear INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId
WHERE (dbo.tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (dbo.tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentId = @DocumentId)
GROUP BY tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId
ORDER BY tblAcc_DocumentDetail.DocumentId, tblAcc_DocumentDetail.Kind, tblAcc_DocumentDetail.KolId
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentsRows]'
GO

CREATE PROCEDURE [dbo].[Get_All_DocumentsRows](@AccountYear smallint, @Branch int, @DocumentId int) AS
SELECT     dbo.tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(dbo.tblAcc_DocumentHeader.DocumentDate) AS dt, dbo.tblAcc_DocumentHeader.DocumentDes, dbo.tblAcc_DocumentDetail.KolId, 
                      dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, dbo.tblAcc_DocumentDetail.RowDes, dbo.tblAcc_DocumentDetail.Bedehkar, 
                      dbo.tblAcc_DocumentDetail.Bestankar, dbo.tBranch.nvcBranchName
FROM         dbo.tblAcc_DocumentHeader INNER JOIN
                      dbo.tblAcc_DocumentDetail ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND 
                      dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND 
                      dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId INNER JOIN
                      dbo.tBranch ON dbo.tblAcc_DocumentHeader.Branch = dbo.tBranch.Branch
WHERE     (dbo.tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (dbo.tblAcc_DocumentHeader.Branch = @Branch) AND (dbo.tblAcc_DocumentHeader.DocumentId = @DocumentId)

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_TafsiliDetails]'
GO


CREATE PROCEDURE [dbo].[Update_tblAcc_TafsiliDetails] (
				
		@Branch int, 		
		@TafsiliId int, 		
		@TafsiliName nvarchar(50), 		
		@Active BIT,
		@AtfId INT
	) 
	
	AS
BEGIN

UPDATE tblAcc_Tafsili

	SET 
		[TafsiliName] = @TafsiliName,
		[Active] = @Active
		
		WHERE Branch = @Branch AND TafsiliId = @TafsiliId
		
	DELETE FROM tblAcc_Tafsili_Atf 
		WHERE Branch = @Branch AND TafsiliId = @TafsiliId

	INSERT dbo.tblAcc_Tafsili_Atf
		(
			Branch,
			TafsiliId, 
			AtfId
		 )
	VALUES
		(
		 @Branch,
		 @TafsiliId,
		 @AtfId
		 )

END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atf_MoeinByAtf]'
GO




CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atf_MoeinByAtf](@Branch int, @AtfId int) AS
SELECT DISTINCT tblAcc_Tafsili.TafsiliId, tblAcc_Tafsili.TafsiliName
FROM         tblAcc_Moein_Atf INNER JOIN
                      tblAcc_Tafsili INNER JOIN
                      tblAcc_Tafsili_Atf ON tblAcc_Tafsili.Branch = tblAcc_Tafsili_Atf.Branch AND tblAcc_Tafsili.TafsiliId = tblAcc_Tafsili_Atf.TafsiliId ON 
                      tblAcc_Moein_Atf.AtfID = tblAcc_Tafsili_Atf.AtfId
WHERE     (tblAcc_Tafsili.Branch = @Branch) AND [tblAcc_Tafsili_Atf].[AtfId]=@AtfId  AND (tblAcc_Tafsili.Active = 1)AND [dbo].[tblAcc_Tafsili].[TafsiliId]<>0
ORDER BY tblAcc_Tafsili.TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentHeaders_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentHeaders_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AccountYear smallint, Branch int, DocumentId int, DocumentDate int, DocumentDes nvarchar(50), State tinyint, DocumentId2 int, DocumentKind tinyint, SaveDate int, UserId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AccountYear, Branch, DocumentId, DocumentDate, DocumentDes, State, DocumentId2, DocumentKind, SaveDate, UserId
		
	FROM [tblAcc_DocumentHeader] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AccountYear, Branch, DocumentId, DocumentDate, DocumentDes, State, DocumentId2, DocumentKind, SaveDate, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Do_SaveInDetails]'
GO
CREATE PROCEDURE [dbo].[Do_SaveInDetails](@AccountYear smallint,
										 @Branch int,
										 @DocumentId int,
										 @ds1 NVARCHAR(4000),
										 @ds2 NVARCHAR(4000),
										 @ds3 NVARCHAR(4000),
										 @ds4 NVARCHAR(4000)
										)
AS
BEGIN
	BEGIN TRAN
	DELETE tblAcc_DocumentDetail 
	WHERE AccountYear = @AccountYear 
		AND Branch = @Branch 
		AND DocumentId = @DocumentId
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds1)
IF @ds2<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds2)
IF @ds3<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds3)
IF @ds4<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds4)

	IF @@ERROR<>0 ROLLBACK TRAN
	ELSE COMMIT TRAN
END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tBranch]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tBranch] (
		
				
		@Branch int
		
		) AS
		
		DELETE [tBranch]
		
		WHERE
		
		
			[Branch] = @Branch




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_Sanad]'
GO
Create PROC [dbo].[Insert_Sanad]
(
@AccountYear SMALLINT,
@Branch INT,
@DocumentId INT,
@DocumentDate INT,
@DocumentDes nvarchar(255),
@State  TinyInt,
@DocumentId2 int,
@DocumentKind tinyint,
@UserId int,
@ds1 NVARCHAR(4000),
@ds2 NVARCHAR(4000),
@ds3 NVARCHAR(4000),
@ds4 NVARCHAR(4000)
)
as

begin
	INSERT INTO [tblAcc_DocumentHeader]
		
	(
		[AccountYear],
		[Branch],
		[DocumentId],
		[DocumentDate],
		[DocumentDes],
		[State],
		[DocumentId2],
		[DocumentKind],
		[SaveDate],
		[UserId]
	)		
		
	VALUES		
	(
		@AccountYear,
		@Branch,
		@DocumentId,
		@DocumentDate,
		@DocumentDes,
		@State,
		@DocumentId2,
		@DocumentKind,
		dbo.ShamsiInt(GetDate()),
		@UserId
	)

---------------------------------------------------------------------------------
UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/



---------------------------------------------------------------------------------------
DELETE tblAcc_DocumentDetail 
	WHERE AccountYear = @AccountYear 
		AND Branch = @Branch 
		AND DocumentId = @DocumentId
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds1)
IF @ds2<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds2)
IF @ds3<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds3)
IF @ds4<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds4)

-------------------------------------------------------------------------------

UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_PayCheck]'
GO
CREATE PROC [dbo].[Get_All_PayCheck]
(
@PayType INT,
@FromDate NVARCHAR(8),
@ToDate NVARCHAR(8)
)
AS
BEGIN
SELECT [intSerialNo],
		[DateT],
		[Sanad_Pardakhti],
		[CheckNo],
		[DateS],
		[PayTafsili],
		[RecTafsili],
		[BankAccountTafsili],
		[Resid],
		[Price]
 FROM [dbo].[tblAcc_PaymentSanad]
WHERE [PaymentTypeId]=@PayType --AND ([PaymentTypeId]<>1 or [PaymentTypeId]=1) --AND  [DateT]>=@FromDate AND [DateT]<=LEFT(@ToDate,2)+ N'/12/30'
end

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Atf]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_Atf] (
		
				
		@AtfID int
		
		) AS
		
		DELETE [tblAcc_Atf]
		
		WHERE
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_SanadPayment]'
GO
CREATE PROC [dbo].[Insert_SanadPayment]
(
	@AccountYear smallint, 		
	@Branch int, 		
	@DocumentId int, 		
	@DocumentDate int, 		
	@DocumentDes nvarchar(100), 		
	@State tinyint, 		
	@DocumentId2 int, 		
	@DocumentKind tinyint, 		
	@UserId INT,
	@ds1 NVARCHAR(4000),
	@ds2 NVARCHAR(4000),
	@ds3 NVARCHAR(4000),
	@ItemNo INT,
	@SerialNo int
)
as
begin
DECLARE @number int
SET @number=
		(SELECT COUNT(*) FROM  
		[tblAcc_DocumentHeader]
		WHERE
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId)
IF @number=0 
	begin
		

INSERT INTO [tblAcc_DocumentHeader]
			
		(
			[AccountYear],
			[Branch],
			[DocumentId],
			[DocumentDate],
			[DocumentDes],
			[State],
			[DocumentId2],
			[DocumentKind],
			[SaveDate],
			[UserId] 
		)		
			
		VALUES		
		(
			@AccountYear,
			@Branch,
			@DocumentId,
			@DocumentDate,
			@DocumentDes,
			@State,
			@DocumentId2,
			@DocumentKind,
			dbo.ShamsiInt(GetDate()),
			@UserId
		)

		DELETE  FROM tblAcc_DocumentDetail
		WHERE   AccountYear = @AccountYear
				AND Branch = @Branch
				AND DocumentId = @DocumentId
		IF @ds1<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate
					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds1)
		END
		IF @ds2<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds2)
		END
		IF @ds3<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds3)
			END



-----------------------------------Do_ValidateDocumentDetail
UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)
--------------------------------------------------------------


------------------------------------Update_tblAcc_RecieveSanad_SanadNo
IF @ItemNo = 2 
UPDATE tblAcc_PaymentSanad
	SET [Sanad_Pardakhti] = @DocumentId
		WHERE [intSerialNo]=@SerialNo
ELSE IF @ItemNo = 3 
UPDATE tblAcc_PaymentSanad
	SET [Sanad_Vosouli] = @DocumentId
	WHERE [intSerialNo]=@SerialNo
ELSE IF @ItemNo = 4 
UPDATE tblAcc_PaymentSanad
	SET Sanad_Bargashti = @DocumentId
		WHERE [intSerialNo]=@SerialNo
ELSE IF @ItemNo = 5 
UPDATE tblAcc_PaymentSanad
	SET Sanad_BargashtiMoshtari = @DocumentId
		WHERE [intSerialNo]=@SerialNo
ELSE IF @ItemNo = 7 OR @ItemNo=8 OR @ItemNo=10 OR @ItemNo=11 OR @ItemNo=12
UPDATE tblAcc_PaymentSanad
	SET [Sanad_Cash] = @DocumentId
		WHERE [intSerialNo]=@SerialNo

---------------------------------------------------------------------

	END	
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[tblAcc_Bank]'
GO
ALTER TABLE [dbo].[tblAcc_Bank] ALTER COLUMN [tintBank] [int] NOT NULL
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Bank] on [dbo].[tblAcc_Bank]'
GO
ALTER TABLE [dbo].[tblAcc_Bank] ADD CONSTRAINT [PK_tblAcc_Bank] PRIMARY KEY CLUSTERED  ([tintBank])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_RecieveType]'
GO
CREATE TABLE [dbo].[tblAcc_RecieveType]
(
[RecieveTypeId] [int] NOT NULL,
[ReceiveTypeName] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_RecieveType] on [dbo].[tblAcc_RecieveType]'
GO
ALTER TABLE [dbo].[tblAcc_RecieveType] ADD CONSTRAINT [PK_tblAcc_RecieveType] PRIMARY KEY CLUSTERED  ([RecieveTypeId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_RecieveSanad]'
GO
CREATE TABLE [dbo].[tblAcc_RecieveSanad]
(
[intSerialNo] [int] NOT NULL IDENTITY(1, 1),
[CheckNo] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[DateS] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Price] [bigint] NOT NULL,
[Descs] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[RecieveTypeId] [int] NOT NULL,
[DateT] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[BankNo] [int] NULL,
[PayKol] [int] NULL,
[PayMoein] [int] NULL,
[PayTafsili] [int] NULL,
[PayTafsiliName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Vagozari_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Vosouli_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Kharj_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Bargashti_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[BargashtiMoshtari_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Sanad_Daryafti] [int] NULL,
[Sanad_Vagozari] [int] NULL,
[Sanad_Vosouli] [int] NULL,
[Sanad_Kharj] [int] NULL,
[Sanad_Bargashti] [int] NULL,
[Sanad_BargashtiMoshtari] [int] NULL,
[Sanad_Cash] [int] NULL,
[Resid] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Void] [bit] NOT NULL CONSTRAINT [DF_tblAcc_RecieveSanad_Void] DEFAULT ((0)),
[KolTaraf] [int] NULL,
[MoeinTaraf] [int] NULL,
[TafsiliTaraf] [int] NULL,
[TafsiliNameTaraf] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Cash_Date] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_RecieveSanad] on [dbo].[tblAcc_RecieveSanad]'
GO
ALTER TABLE [dbo].[tblAcc_RecieveSanad] ADD CONSTRAINT [PK_tblAcc_RecieveSanad] PRIMARY KEY CLUSTERED  ([intSerialNo])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating index [IX_tblAcc_RecieveSanad] on [dbo].[tblAcc_RecieveSanad]'
GO
CREATE NONCLUSTERED INDEX [IX_tblAcc_RecieveSanad] ON [dbo].[tblAcc_RecieveSanad] ([CheckNo])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Cheque_Received_ByDate]'
GO

--==========================
-- براي ريپورت هست
Create PROCEDURE [dbo].[Get_All_Cheque_Received_ByDate] 
(
@SystemDate NVARCHAR(10) ,
@SystemDay AS NVARCHAR(20) ,
@SystemTime NVARCHAR(5) ,
@BankNo INT ,
@RecieveTypeId TINYINT ,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) ,
@OrderView INT ,
@AscDesc INT ,
@BankDesc NVARCHAR(30) ,
@ChequeType NVARCHAR(20) ,
@OrderDesc NVARCHAR(20) ,
@SortDesc NVARCHAR(20)
)
AS

		
IF @OrderView = 0 OR @OrderView = 3 
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		,@FromDate AS FromDate ,@ToDate AS ToDate , @BankDesc  AS BankDesc, @ChequeType AS ChequeType, @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate  )--OR tblAcc_RecieveSanad.DateS IS NULL

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
			tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		,@FromDate AS FromDate ,@ToDate AS ToDate , @BankDesc  AS BankDesc, @ChequeType AS ChequeType,@OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate  )--OR tblAcc_RecieveSanad.DateS IS NULL

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			END DESC 
		END 
END 
ELSE IF @OrderView = 1 OR @OrderView = 2 OR @OrderView = 4
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
			tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		,@FromDate AS FromDate ,@ToDate AS ToDate , @BankDesc  AS BankDesc, @ChequeType AS ChequeType, @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate  )--OR tblAcc_RecieveSanad.DateS IS NULL

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 4 THEN TafsiliNameTaraf
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
			tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		,@FromDate AS FromDate ,@ToDate AS ToDate , @BankDesc  AS BankDesc, @ChequeType AS ChequeType, @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate) -- OR tblAcc_RecieveSanad.DateS IS NULL 

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 4 THEN TafsiliNameTaraf
			END DESC 
		END 
END 


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Atf]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------		
	
CREATE PROCEDURE [dbo].[Insert_tblAcc_Atf] (
				
		@AtfID int, 		
		@AtfName nvarchar(50), 		
		@Active bit
	) 
	
	AS
		
	INSERT INTO [tblAcc_Atf]
		
	(
		[AtfID],
		[AtfName],
		[Active]
	)		
		
	VALUES		
	(
		@AtfID,
		@AtfName,
		@Active
	)
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_ChequePrintTemplate]'
GO
CREATE TABLE [dbo].[tblAcc_ChequePrintTemplate]
(
[PrintTemplateID] [int] NOT NULL,
[Name] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Path] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[Active] [bit] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_DefineChequePrintTemplate] on [dbo].[tblAcc_ChequePrintTemplate]'
GO
ALTER TABLE [dbo].[tblAcc_ChequePrintTemplate] ADD CONSTRAINT [PK_tblAcc_DefineChequePrintTemplate] PRIMARY KEY CLUSTERED  ([PrintTemplateID])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_ChequePrintTemplates]'
GO

CREATE PROC  [dbo].[Get_ChequePrintTemplates]
as
begin

SELECT * FROM tblAcc_ChequePrintTemplate
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Cheque_Received]'
GO

--براي صفحه نمايش بدون شرط تاريخ
CREATE    PROCEDURE dbo.Get_All_Cheque_Received
(
@SystemDate NVARCHAR(10) ,
@SystemDay AS NVARCHAR(20) ,
@SystemTime NVARCHAR(5) ,
@BankNo INT ,
@RecieveTypeId TINYINT ,
@OrderView INT ,
@AscDesc INT ,
@BankDesc NVARCHAR(30) ,
@ChequeType NVARCHAR(20) ,
@OrderDesc NVARCHAR(20) ,
@SortDesc NVARCHAR(20)
)
AS

		
IF @OrderView = 0 OR @OrderView = 3 
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @BankDesc  AS BankDesc, @ChequeType AS ChequeType,@OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			--AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
			tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @BankDesc  AS BankDesc, @ChequeType AS ChequeType, @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			--AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			END DESC 
		END 
END 
ELSE IF @OrderView = 1 OR @OrderView = 2 OR @OrderView = 4 
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
			tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @BankDesc  AS BankDesc, @ChequeType AS ChequeType,@OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			--AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 4 THEN TafsiliNameTaraf
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
			tblAcc_RecieveSanad.* , ReceiveTypeName , ISNULL(tblAcc_Bank.nvcBankName ,N'') AS nvcBankName 
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @BankDesc  AS BankDesc, @ChequeType AS ChequeType,@OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_RecieveSanad
		INNER JOIN dbo.tblAcc_RecieveType ON dbo.tblAcc_RecieveSanad.RecieveTypeId = dbo.tblAcc_RecieveType.RecieveTypeId
		LEFT OUTER JOIN dbo.tblAcc_Bank ON dbo.tblAcc_Bank.tintBank = dbo.tblAcc_RecieveSanad.BankNo

		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
			--AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 4 THEN TafsiliNameTaraf
			END DESC 
		END 
END 




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_UGroups]'
GO
CREATE TABLE [dbo].[tblAcc_UGroups]
(
[UGroupId] [tinyint] NOT NULL,
[UGroupName] [nvarchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_UGroups_UGroupName] DEFAULT ('')
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_UGroups] on [dbo].[tblAcc_UGroups]'
GO
ALTER TABLE [dbo].[tblAcc_UGroups] ADD CONSTRAINT [PK_tblAcc_UGroups] PRIMARY KEY CLUSTERED  ([UGroupId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_UGroups]'
GO



-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_UGroups] (
		
				
		@UGroupId tinyint
		
		) AS
		
		DELETE [tblAcc_UGroups]
		
		WHERE
		
		
			[UGroupId] = @UGroupId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[GetAllTafsiliById]'
GO
CREATE PROC [dbo].[GetAllTafsiliById]
(
@Id INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_Tafsili] WHERE 
[TafsiliId]=@Id
END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_RecieveSanad]'
GO


CREATE PROCEDURE [dbo].[Insert_tblAcc_RecieveSanad]
(
	@CheckNo NVARCHAR(20),
	@DateS NVARCHAR(10),
	@Price BIGINT ,
	@Descs NVARCHAR(255),
	@RecieveTypeId TINYINT,
	@DateT NVARCHAR(10),
	@BankNo INT,
	@PayKol INT ,
	@PayMoein INT ,
	@PayTafsili INT ,
	@PayTafsiliName NVARCHAR(50) ,
	@Resid NVARCHAR(255),
	@Result INT OUT 

) 
AS

IF @CheckNo=N''
SET @CheckNo=NULL

IF @DateS=N''
SET @DateS=NULL
IF @BankNo=0
SET @BankNo=NULL

	
INSERT INTO dbo.tblAcc_RecieveSanad (
	CheckNo,
	DateS,
	Price,
	Descs,
	RecieveTypeId,
	DateT,
	BankNo,
	PayKol,
	PayMoein,
	PayTafsili,
	PayTafsiliName,
	Resid
) VALUES ( 
	@CheckNo ,
	@DateS ,
	@Price ,
	@Descs ,
	@RecieveTypeId ,
	@DateT ,
	@BankNo ,
	@PayKol ,
	@PayMoein ,
	@PayTafsili ,
	@PayTafsiliName,
	@Resid
		)
     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result =@@IDENTITY
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Do_ValidateDocumentDetail]'
GO


CREATE   PROCEDURE [dbo].[Do_ValidateDocumentDetail](
		 	
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int

) AS

UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Atfs_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Atfs_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Atf]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad_SanadNo]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad_SanadNo]
(
	@intSeriaNo INT ,
	@sanadNo INT ,
	@ItemNo INT ,
	@nvcDate NVARCHAR(10) ,
	@KolTaraf INT ,
	@MoeinTaraf INT ,
	@TafsiliTaraf INT ,
	@TafsiliNameTaraf nvarchar(50)
)
AS
IF @ItemNo = 1 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Daryafti = @sanadNo,
		RecieveTypeId = @ItemNo ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 
ELSE IF @ItemNo = 2 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Kharj = @sanadNo ,
		RecieveTypeId = @ItemNo ,
		Kharj_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 
ELSE IF @ItemNo = 3 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Vagozari = @sanadNo,
		RecieveTypeId = @ItemNo ,
		Vagozari_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 
ELSE IF @ItemNo = 4 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Vosouli = @sanadNo,
		RecieveTypeId = @ItemNo ,
		Vosouli_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 
ELSE IF @ItemNo = 5 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Bargashti = @sanadNo,
		RecieveTypeId = @ItemNo ,
		Bargashti_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 
ELSE IF @ItemNo = 6 
UPDATE tblAcc_RecieveSanad
	SET Sanad_BargashtiMoshtari = @sanadNo,
		RecieveTypeId = @ItemNo ,
		BargashtiMoshtari_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 

ELSE IF    @ItemNo = 7 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Cash = @sanadNo,
		RecieveTypeId = @ItemNo ,
		Cash_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 

ELSE IF    @ItemNo = 8 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Cash = @sanadNo,
		RecieveTypeId = @ItemNo ,
		Cash_Date = @nvcDate ,
		KolTaraf = @KolTaraf,
		MoeinTaraf = @MoeinTaraf,
		TafsiliTaraf = @TafsiliTaraf,
		TafsiliNameTaraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_Sanad]'
GO
Create PROC [dbo].[Update_Sanad]
(
@AccountYear SMALLINT,
@Branch INT,
@DocumentId INT,
@DocumentDate INT,
@DocumentDes nvarchar(255),
@State  TinyInt,
@DocumentId2 int,
@DocumentKind tinyint,
@UserId int,
@ds1 NVARCHAR(4000),
@ds2 NVARCHAR(4000),
@ds3 NVARCHAR(4000),
@ds4 NVARCHAR(4000)
)
as

begin
	UPDATE [tblAcc_DocumentHeader]
		
		SET
		
		
				[AccountYear] = @AccountYear,
				[Branch] = @Branch,
				[DocumentId] = @DocumentId,
				[DocumentDate] = @DocumentDate,
				[DocumentDes] = @DocumentDes,
				[State] = @State,
				[DocumentId2] = @DocumentId2,
				[DocumentKind] = @DocumentKind,
				[SaveDate] = dbo.ShamsiInt(GetDate()),
				[UserId] = @UserId

		
		WHERE
		
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId

---------------------------------------------------------------------------------
UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/



---------------------------------------------------------------------------------------
DELETE tblAcc_DocumentDetail 
	WHERE AccountYear = @AccountYear 
		AND Branch = @Branch 
		AND DocumentId = @DocumentId
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds1)
IF @ds2<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds2)
IF @ds3<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds3)
IF @ds4<>''
	INSERT INTO tblAcc_DocumentDetail
	            (AccountYear,
				 Branch,
				 DocumentId,
				 RowId,
				 KolId,
				 MoeinId,
				 TafsiliId,
				 RowDes,
				 Bedehkar,
				 Bestankar,
				 kind,
				 SaveDate,
				 UserId)
	SELECT      AccountYear,
				Branch, 
				DocumentId, 
				RowId, 
				KolId, 
				MoeinId, 
				TafsiliId, 
				RowDes, 
				Bedehkar, 
				Bestankar, 
				kind, 
				SaveDate, 
				UserId
	FROM dbo.Split_Acc(@ds4)

-------------------------------------------------------------------------------

UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)

/*
UPDATE    tblAcc_DocumentDetail
SET              Bedehkar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 2))

UPDATE    tblAcc_DocumentDetail
SET              Bestankar = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0) AND (EXISTS
                          (SELECT     *
                             FROM         tblAcc_Moein
                             WHERE     tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId AND tblAcc_Moein.Kind = 1))
*/

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_AllTafsiliById]'
GO
CREATE PROC [dbo].[Get_AllTafsiliById]
(
@Id INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_Tafsili] WHERE 
[TafsiliId]=@Id
END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Tafsili_Atf]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Update_tblAcc_Tafsili_Atf] (
				
				
		@Branch int, 		
		@TafsiliId int, 		
		@AtfId int

		
		) AS
		
		UPDATE [tblAcc_Tafsili_Atf]
		
		SET
		
		
				[Branch] = @Branch,
				[TafsiliId] = @TafsiliId,
				[AtfId] = @AtfId

		
		WHERE
		
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId AND 
			[AtfId] = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_UGroupss_Count]'
GO



-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_UGroupss_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_UGroups]
		
		
	



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atfs]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atfs]
				
		AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[AtfId]
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Do_Commit]'
GO

CREATE PROCEDURE [dbo].[Do_Commit](@AccountYear smallint, @Branch int, @DocumentDate int, @NewDocumentId int) AS
UPDATE tblAcc_DocumentHeader
SET DocumentId = -DocumentId, DocumentId2 = DocumentId
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (State = 2) AND (DocumentDate <= @DocumentDate)

UPDATE    tblAcc_DocumentDetail
SET DocumentId = -DocumentId
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId IN (SELECT - DocumentId FROM tblAcc_DocumentHeader))

DECLARE cr CURSOR
KEYSET
FOR SELECT     tblAcc_DocumentHeader.DocumentId
FROM         tblAcc_DocumentHeader
WHERE (AccountYear = @AccountYear) AND (Branch = @Branch) AND (tblAcc_DocumentHeader.State = 2) AND (tblAcc_DocumentHeader.DocumentDate <= @DocumentDate)
ORDER BY tblAcc_DocumentHeader.DocumentDate, tblAcc_DocumentHeader.DocumentId DESC

DECLARE @DocumentId int

OPEN cr

FETCH NEXT FROM cr INTO @DocumentId
WHILE (@@fetch_status <> -1)
BEGIN
	IF (@@fetch_status <> -2)
	BEGIN
		UPDATE    tblAcc_DocumentHeader
		SET State = 3, DocumentId = @NewDocumentId
		WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (State = 2) AND (DocumentId = @DocumentId)

		UPDATE    tblAcc_DocumentDetail
		SET DocumentId = @NewDocumentId
		WHERE (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId)
	END
	FETCH NEXT FROM cr INTO @DocumentId
	SET @NewDocumentId = @NewDocumentId + 1
END

CLOSE cr
DEALLOCATE cr

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliByName]'
GO

CREATE PROC	 [dbo].[Get_TafsiliByName](  @Branch int, 
								 @KolId int,
								 @MoeinId int,
								 @Search nvarchar(100),
								 @State INt
								)
AS	
BEGIN
IF @State=1
BEGIN
	SELECT tblAcc_Tafsili.TafsiliId,
					 tblAcc_Tafsili.TafsiliName
	FROM dbo.tblAcc_Moein_Atf
			JOIN tblAcc_Tafsili_Atf ON dbo.tblAcc_Moein_Atf.AtfId = dbo.tblAcc_Tafsili_Atf.AtfId
			JOIN tblAcc_Tafsili ON dbo.tblAcc_Tafsili_Atf.Branch = dbo.tblAcc_Tafsili.Branch 
							AND dbo.tblAcc_Tafsili_Atf.TafsiliId = dbo.tblAcc_Tafsili.TafsiliId
	WHERE (tblAcc_Tafsili.Branch = @Branch) 
			AND (tblAcc_Moein_Atf.KolID = @KolId) 
			AND (tblAcc_Moein_Atf.MoeinId = @MoeinId) 
			AND (tblAcc_Tafsili.Active = 1)
			AND CHARindex(@Search,tblAcc_Tafsili.TafsiliName)>0
END
ELSE IF @State=2
BEGIN
		SELECT 
				[Branch],
				[TafsiliId],
				[TafsiliName],
				[Active]		
		FROM 
		[tblAcc_Tafsili]
		WHERE 
		[Branch] = @Branch
		AND CHARindex(@Search,TafsiliName)>0
END
END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Tafsili_Atf]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------		
	
CREATE PROCEDURE [dbo].[Insert_tblAcc_Tafsili_Atf] (
				
		@Branch int, 		
		@TafsiliId int, 		
		@AtfId int
	) 
	
	AS
		
	INSERT INTO [tblAcc_Tafsili_Atf]
		
	(
		[Branch],
		[TafsiliId],
		[AtfId]
	)		
		
	VALUES		
	(
		@Branch,
		@TafsiliId,
		@AtfId
	)
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_Max_Tafsili]'
GO
CREATE PROC [dbo].[Get_Max_Tafsili]
as
begin
SELECT ISNULL(MAX(TafsiliId),0)+1 FROM tblAcc_Tafsili 
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Atfs]'
GO
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Atfs]
				
		AS
		
		SELECT 
		
		
				[AtfID],
				[AtfName],
				[Active]
		
		FROM 
		
		[tblAcc_Atf]
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atfs_For_TafsiliId]'
GO




CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atfs_For_TafsiliId](@Branch int, @TafsiliId int)
				
		AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[AtfId]
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		
		WHERE 
		
		Branch = @Branch AND TafsiliId = @TafsiliId
		
		ORDER BY 
		
		AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_UGroupss]'
GO

-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_UGroupss]
				
		AS
		
		SELECT 
		
		
				[UGroupId],
				[UGroupName]
		
		FROM 
		
		[tblAcc_UGroups]
		

		ORDER BY UGroupId


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_ReceivedCheckByTypeDirect]'
GO
CREATE PROC [dbo].[Get_ReceivedCheckByTypeDirect](
@ReceivedType INT,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) 
)
AS
BEGIN
SELECT *,(SELECT [ReceiveTypeName] FROM [dbo].[tblAcc_RecieveType] WHERE [dbo].[tblAcc_RecieveType].[RecieveTypeId]=[dbo].[tblAcc_RecieveSanad].[RecieveTypeId])AS ReceiveTypeName
FROM [dbo].[tblAcc_RecieveSanad]
where [RecieveTypeId]=@ReceivedType
AND [DateT]>=@FromDate AND [DateT]<=@ToDate


end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atfs_Count_ForAll]'
GO




CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atfs_Count_ForAll](@Branch int, @TafsiliId int, @AtfId int)
				
		AS
		
		SELECT
		
			COUNT(*) AS ct
		
		FROM
		
			[tblAcc_Tafsili_Atf]
		
		WHERE
		
			Branch = @Branch AND TafsiliId = @TafsiliId AND AtfId = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DaftarKol2]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_DaftarKol2](@AccountYear smallint, @Branch int, @KolId int, @d1 int= 0, @d2 int = 0) AS
SELECT     0 AS DocumentId, '' AS sdate, '' AS DocumentDes, tblAcc_Kol.KolId, MAX(tblAcc_Kol.KolName) AS KolName,
                      SUM(tblAcc_DocumentDetail.Bedehkar) AS sd, SUM(tblAcc_DocumentDetail.Bestankar) AS ss, 0 AS kind
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId=tblAcc_Kol.KolId
WHERE (tblAcc_DocumentHeader.State > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND(tblAcc_DocumentHeader.DocumentDate < @d1) AND ((@KolId = 0) OR (tblAcc_DocumentDetail.KolId = @KolId))
GROUP BY tblAcc_Kol.KolId
UNION ALL
SELECT     tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdate, tblAcc_DocumentHeader.DocumentDes, tblAcc_Kol.KolId, tblAcc_Kol.KolName,
                      SUM(tblAcc_DocumentDetail.Bedehkar) AS SBDA, SUM(tblAcc_DocumentDetail.Bestankar) AS SBSA, tblAcc_DocumentDetail.kind
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId=tblAcc_Kol.KolId
WHERE (State > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND(tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND ((@KolId = 0) OR (tblAcc_DocumentDetail.KolId = @KolId))
GROUP BY tblAcc_Kol.KolId, tblAcc_DocumentHeader.DocumentDate, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDes, tblAcc_Kol.KolName, tblAcc_DocumentDetail.kind
ORDER BY tblAcc_Kol.KolId, sdate, DocumentId, kind



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Atf]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Update_tblAcc_Atf] (
				
				
		@AtfID int, 		
		@AtfName nvarchar(50), 		
		@Active bit

		
		) AS
		
		UPDATE [tblAcc_Atf]
		
		SET
		
		
				[AtfID] = @AtfID,
				[AtfName] = @AtfName,
				[Active] = @Active

		
		WHERE
		
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_Branch_TafsiliId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_Branch_TafsiliId_Count] (
			
			
			@Branch int,
			@TafsiliId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsilis_Without_Branch]'
GO
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsilis_Without_Branch]
				
		AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[TafsiliName],
				[Active]
		
		FROM 
		
		[tblAcc_Tafsili]
		
		
		ORDER BY 
		
		[TafsiliId]

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Get_All_tAccountYears]'
GO

ALTER PROCEDURE [dbo].[Get_All_tAccountYears]
			
AS
	
	SELECT *
	
	FROM 
	
	[tAccountYears]
	
	ORDER BY AccountYear
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Atf_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Atf_ByID] (
		 	
				
		@AtfID int

		) AS
		
		SELECT 
		
		
				[AtfID],
				[AtfName],
				[Active]
		
		FROM 
		
		[tblAcc_Atf]
		
		WHERE
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atfs_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atfs_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Tafsili_Atf]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Split_Salary]'
GO
SET ANSI_NULLS OFF
GO

CREATE Function [dbo].[Split_Salary]

(
    @nvcMainString nText
)

RETURNS  @ReturnTable TABLE(
	Ppno INT  ,
	Tafsili INT  ,
	DastmozdRooz INT  ,
	KarkardRooz INT  ,
	KarkardMah INT  ,
	FeeEzafe INT  ,
	SaatEzafe INT  ,
	KarkardEzafe INT ,
	BimeShakhs INT   ,
	MaliatShakhs INT   ,
	Kosourat INT   ,
	NetKarkardMah INT   ,
	BimeKarfarma INT  ,   
	MaliatKarfarma INT   
)	
As

BEGIN

	IF @nvcMainString IS NOT  NULL
	BEGIN
		DECLARE @intDelimiterPosField  INT
		DECLARE @intDelimiterPosRecord INT

		DECLARE @Ppno INT  
		DECLARE @Tafsili INT  
		DECLARE @DastmozdRooz INT  
		DECLARE @KarkardRooz INT  
		DECLARE @KarkardMah INT  
		DECLARE @FeeEzafe INT  
		DECLARE @SaatEzafe INT  
		DECLARE @KarkardEzafe INT 
		DECLARE @BimeShakhs INT   
		DECLARE @MaliatShakhs INT   
		DECLARE @Kosourat INT   
		DECLARE @NetKarkardMah INT   
		DECLARE @BimeKarfarma INT   
		DECLARE @MaliatKarfarma INT   

		DECLARE @TempTable Table (nvcMainString nText)

		insert into @TempTable values (@nvcMainString)

		SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
		SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)

		WHILE @intDelimiterPosRecord <> 0
			BEGIN
		--**********
			SET @Ppno = ( select CAST(LTrim(RTrim(SUBSTRING(nvcMainString, 1, @intDelimiterPosField - 1))) AS INT )  from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @Tafsili = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @DastmozdRooz = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @KarkardRooz = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @KarkardMah = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @FeeEzafe = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @SaatEzafe = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @KarkardEzafe = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @BimeShakhs = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @MaliatShakhs = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @Kosourat = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @NetKarkardMah = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @BimeKarfarma = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable)
			SET @MaliatKarfarma = ( select  CAST(LTrim(RTrim(SUBSTRING(nvcMainString,1 , @intDelimiterPosField - 1))) AS INT) from @TempTable )
			Update @TempTable SET nvcMainString = ( Select SUBSTRING(nvcMainString, @intDelimiterPosField + 1, DataLength(nvcMainString) - @intDelimiterPosField  ) from @TempTable )

			SET @intDelimiterPosRecord = ( select patindex('%/%' , nvcMainString)  from @TempTable)
		    Update @TempTable SET nvcMainString = ( select SUBSTRING(nvcMainString, @intDelimiterPosRecord + 1, DataLength(nvcMainString) - @intDelimiterPosRecord  ) from @TempTable )

			INSERT INTO @ReturnTable( Ppno , Tafsili , DastmozdRooz , KarkardRooz , KarkardMah , FeeEzafe , 
					SaatEzafe , KarkardEzafe ,BimeShakhs ,MaliatShakhs , Kosourat ,NetKarkardMah ,BimeKarfarma , MaliatKarfarma)
			VALUES (  @Ppno ,@Tafsili ,@DastmozdRooz ,@KarkardRooz ,@KarkardMah ,@FeeEzafe ,@SaatEzafe ,
					  @KarkardEzafe ,@BimeShakhs ,@MaliatShakhs ,@Kosourat ,@NetKarkardMah ,@BimeKarfarma , @MaliatKarfarma)
		            
			SET @intDelimiterPosField = ( select patindex('%;%' , nvcMainString) from @TempTable )
			SET @intDelimiterPosRecord = ( Select patindex('%/%' , nvcMainString)  from @TempTable )

			End
	END 

RETURN 


End

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tSalaryD]'
GO
SET ANSI_NULLS ON
GO
CREATE TABLE [dbo].[tSalaryD]
(
[SalaryId] [int] NOT NULL,
[Branch] [int] NOT NULL,
[Ppno] [int] NOT NULL,
[Tafsili] [int] NOT NULL,
[DastmozdRooz] [int] NOT NULL,
[KarkardRooz] [int] NOT NULL,
[KarkardMah] [int] NOT NULL,
[FeeEzafe] [int] NOT NULL,
[SaatEzafe] [int] NOT NULL,
[KarkardEzafe] [int] NOT NULL,
[BimeShakhs] [int] NOT NULL,
[MaliatShakhs] [int] NOT NULL,
[Kosourat] [int] NOT NULL,
[NetKarkardMah] [int] NOT NULL,
[BimeKarfarma] [int] NOT NULL,
[MaliatKarfarma] [int] NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating index [IX_tSalaryD] on [dbo].[tSalaryD]'
GO
CREATE NONCLUSTERED INDEX [IX_tSalaryD] ON [dbo].[tSalaryD] ([SalaryId], [Ppno])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tSalaryM]'
GO
CREATE TABLE [dbo].[tSalaryM]
(
[SalaryId] [int] NOT NULL IDENTITY(1, 1),
[Branch] [int] NOT NULL,
[AccountYear] [smallint] NOT NULL,
[Month] [int] NOT NULL,
[UserId] [int] NOT NULL,
[SanadNo] [int] NULL,
[SanadDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[DocumentDesc] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[nvcAddDate] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tSalaryM] on [dbo].[tSalaryM]'
GO
ALTER TABLE [dbo].[tSalaryM] ADD CONSTRAINT [PK_tSalaryM] PRIMARY KEY CLUSTERED  ([SalaryId], [Branch])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating index [IX_tSalaryM] on [dbo].[tSalaryM]'
GO
CREATE NONCLUSTERED INDEX [IX_tSalaryM] ON [dbo].[tSalaryM] ([AccountYear], [Month])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tSalaryMD]'
GO

CREATE PROCEDURE [dbo].[Update_tSalaryMD]
(
  @Branch INT ,
  @UserId INT ,
  @DocumentDesc NVARCHAR(100) ,
  @st NVARCHAR(4000) ,
  @SalaryId INT ,
  @SanadNo INT ,
  @SanadDate NVARCHAR(10) ,
  @Result INT OUT 
)
AS 

IF @SanadNo = 0 SET @SanadNo = NULL 
IF @SanadDate = N'' SET @SanadDate = NULL 
BEGIN TRAN

UPDATE dbo.tSalaryM
	SET 
          UserId = @UserId ,
          DocumentDesc = @DocumentDesc ,
          SanadNo = @SanadNo ,
          SanadDate = @SanadDate
	WHERE SalaryId = @SalaryId AND Branch = @Branch
	
IF @@Error <> 0 GOTO EventHandler
DELETE FROM dbo.tSalaryD
	WHERE SalaryId = @SalaryId AND Branch = @Branch
IF @st <> N''
BEGIN 	
	INSERT INTO dbo.tSalaryD
			( SalaryId ,
			  Branch ,
			  Ppno ,
			  Tafsili ,
			  DastmozdRooz ,
			  KarkardRooz ,
			  KarkardMah ,
			  FeeEzafe ,
			  SaatEzafe ,
			  KarkardEzafe ,
			  BimeShakhs ,
			  MaliatShakhs ,
			  Kosourat ,
			  NetKarkardMah ,
			  BimeKarfarma ,
			  MaliatKarfarma
			)
	SELECT    @SalaryId ,
			  @Branch ,
			  Ppno ,
			  Tafsili ,
			  DastmozdRooz ,
			  KarkardRooz ,
			  KarkardMah ,
			  FeeEzafe ,
			  SaatEzafe ,
			  KarkardEzafe ,
			  BimeShakhs ,
			  MaliatShakhs ,
			  Kosourat ,
			  NetKarkardMah ,
			  BimeKarfarma ,
			  MaliatKarfarma
			FROM dbo.Split_Salary(@st)

END 

COMMIT TRAN
SET @Result = @SalaryId
RETURN @Result

EventHandler:

ROLLBACK TRAN
SET @Result = -1
RETURN @Result
        
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_UGroups]'
GO



-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------		
	
CREATE PROCEDURE [dbo].[Insert_tblAcc_UGroups] (
				
		@UGroupId tinyint, 		
		@UGroupName nvarchar(40)
	) 
	
	AS
		
	INSERT INTO [tblAcc_UGroups]
		
	(
		[UGroupId],
		[UGroupName]
	)		
		
	VALUES		
	(
		@UGroupId,
		@UGroupName
	)
		
	



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_General_PaymentSanad]'
GO
-------------------------------------
CREATE proc [dbo].[Update_General_PaymentSanad]
(
@SerialNo int,
@DateT NVARCHAR(10),
@DateS NVARCHAR(10),
@Price BIGINT ,
@Resid NVARCHAR(255),
@Descs nvarchar(255),
@RecTafsili int,
@Taraf  NVARCHAR(50) ,
@PayTafsili int,
@PayTafsiliName NVARCHAR(50) ,
@Vosouli_Date nvarchar(10),
@Bargashti_Date nvarchar(10),
@BargashtiMoshtari_Date nvarchar(10),
@Branch int,
@SanadNo int,
@UserID int,
@AccountYear smallint,
@DocumentDate int,
@RowDesc nvarchar(255),
@TafsiliBedehkar INT,
@TafsiliBestankar INT
)
as 
begin
UPDATE [tblAcc_PaymentSanad]
   SET [DateS] = (case when @DateS=N'' then DateS else @DateS end)
      ,[Price] = (case when @Price=-1 then [Price] else @Price end)
      ,[Descs] = (case when @Descs=N'' then [Descs] else @Descs end)

      ,[DateT] = (case when @DateT=N'' then DateT else @DateT end)

      ,[RecTafsili] = (case when @RecTafsili=-1 then RecTafsili else @RecTafsili end)
      ,[Taraf] = (case when @Taraf=N'' then Taraf else @Taraf end)
      ,[PayTafsili] =  (case when @PayTafsili=-1 then PayTafsili else @PayTafsili end)
      ,[PayTafsiliName] =  (case when @PayTafsiliName=N'' then PayTafsiliName else @PayTafsiliName end)
      
      ,[Vosouli_Date] =(case when @Vosouli_Date=N'' then Vosouli_Date else @Vosouli_Date end)
      ,[BargashtiMoshtari_Date] = (case when @BargashtiMoshtari_Date=N'' then BargashtiMoshtari_Date else @BargashtiMoshtari_Date end)
      ,Bargashti_Date = (case when @Bargashti_Date=N'' then Bargashti_Date else @Bargashti_Date end)
      
      ,[Resid] =(case when @Resid=N'' then Resid else @Resid end)
 WHERE intSerialNo=@SerialNo
 
 
 UPDATE [tblAcc_DocumentHeader]
   SET [AccountYear] = @AccountYear
      ,[Branch] = @Branch
      ,[DocumentDate] = @DocumentDate
      ,[SaveDate] = dbo.ShamsiInt(GetDate())
      ,[UserId] = @UserId
 WHERE DocumentId=@SanadNo  and Branch=@Branch and AccountYear=@AccountYear
 
 UPDATE [tblAcc_DocumentDetail]
   SET [AccountYear] = @AccountYear
      ,[Branch] = @Branch

      ,[TafsiliId] = @TafsiliBedehkar
      ,[RowDes] = @RowDesc
      ,[Bedehkar] =(case when @Price=-1 then [Bedehkar] else @Price end)
      ,[Bestankar] = 0

      ,[SaveDate] = dbo.ShamsiInt(GetDate())
      ,[UserId] = @UserId
      
      ,[CheckDate] = (case when @DateS=N'' then null else cast('13'+replace(@DateS,'/','') as int) end)
      
 WHERE DocumentId=@SanadNo and RowId=1  and Branch=@Branch and AccountYear=@AccountYear
 
 UPDATE [tblAcc_DocumentDetail]
   SET [AccountYear] = @AccountYear
      ,[Branch] = @Branch

      ,[TafsiliId] = @TafsiliBestankar
      ,[RowDes] = @RowDesc
      ,[Bedehkar] = 0
      ,[Bestankar] = (case when @Price=-1 then [Bestankar] else @Price end)

      ,[SaveDate] = dbo.ShamsiInt(GetDate())
      ,[UserId] = @UserId
      
      ,[CheckDate] = (case when @DateS=N'' then null else cast('13'+replace(@DateS,'/','') as int) end)
      
 WHERE DocumentId=@SanadNo and RowId=2  and Branch=@Branch and AccountYear=@AccountYear
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_Branch_TafsiliId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_Branch_TafsiliId] (
			
			
			@Branch int,
			@TafsiliId int
				
		) AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[AtfId]
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tSalaryM]'
GO

CREATE PROCEDURE [dbo].[Get_tSalaryM]
(
  @Branch INT  ,
  @AccountYear SMALLINT ,
  @Month INT  
)
AS 

SELECT * FROM tsalaryM
WHERE Branch = @Branch 
	AND AccountYear = @AccountYear
	AND [Month] = @Month
	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DaftarKol1]'
GO

CREATE   PROCEDURE [dbo].[Get_All_DaftarKol1](@AccountYear smallint, @Branch int, @KolId int, @d1 int= 0, @d2 int = 0) AS
SELECT     0 AS DocumentId, ' ' AS sdate, ' ' AS RowDes, tblAcc_DocumentDetail.KolId, MAX(tblAcc_Kol.KolName) AS KolName, 0 AS MoeinId, 0 AS TafsiliId,
                      ' ' AS OnvanHesab, SUM(tblAcc_DocumentDetail.Bedehkar) AS Bedehkar, SUM(tblAcc_DocumentDetail.Bestankar) AS Bestankar
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId=tblAcc_Kol.KolId
WHERE (tblAcc_DocumentHeader.State > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND(tblAcc_DocumentHeader.DocumentDate < @d1) AND ((@KolId = 0) OR (tblAcc_DocumentDetail.KolId = @KolId))
GROUP BY tblAcc_DocumentDetail.KolId
UNION
SELECT     tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdate, tblAcc_DocumentDetail.RowDes, tblAcc_DocumentDetail.KolId, tblAcc_Kol.KolName, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId,
                      ISNULL(tblAcc_Tafsili.TafsiliName , '') AS OnvanHesab , 
                      --CASE tblAcc_DocumentDetail.TafsiliId WHEN 0 THEN tblAcc_Moein.MoeinName ELSE tblAcc_Tafsili.TafsiliName END AS OnvanHesab, 
                      tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                      tblAcc_Kol ON tblAcc_DocumentDetail.KolId=tblAcc_Kol.KolId INNER JOIN
                      tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId LEFT OUTER JOIN
                      tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
WHERE (State > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND(tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND ((@KolId = 0) OR (tblAcc_DocumentDetail.KolId = @KolId))
ORDER BY tblAcc_DocumentDetail.KolId, sdate, DocumentId





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Atf_ByID_Count]'
GO




CREATE PROCEDURE [dbo].[Get_tblAcc_Atf_ByID_Count] (
		 	
				
		@AtfID int

		) AS
		
		SELECT 
		
		
				COUNT([AtfID]) AS ct
		
		FROM 
		
		[tblAcc_Atf]
		
		WHERE
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Tafsili_Atf]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_Tafsili_Atf] (
		
				
		@Branch int, 		
		@TafsiliId int, 		
		@AtfId int
		
		) AS
		
		DELETE [tblAcc_Tafsili_Atf]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId AND 
			[AtfId] = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tSalaryD]'
GO

CREATE PROCEDURE [dbo].[Get_tSalaryD]
(
  @Branch INT  ,
  @SalaryId INT  
)
AS 

SELECT * FROM tsalaryD
WHERE Branch = @Branch 
	AND SalaryId = @SalaryId
	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_AtfMoeinByKolByMoein]'
GO
CREATE PROC [dbo].[Get_All_AtfMoeinByKolByMoein]
(
@KolId INT,
@MoeinId INT
)
as

begin
SELECT [AtfId]
      ,[AtfName]
      ,[Active]
	,(SELECT COUNT(*) FROM tblAcc_Moein_Atf WHERE KolId=@KolId AND MoeinId=@MoeinId AND [dbo].[tblAcc_Moein_Atf] .[AtfId]=[dbo].[tblAcc_Atf].[AtfId])AS Relation
  FROM [tblAcc_Atf]
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atf_Moein]'
GO




CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atf_Moein](@Branch int, @KolId int, @MoeinId int) AS
SELECT DISTINCT tblAcc_Tafsili.TafsiliId, tblAcc_Tafsili.TafsiliName
FROM         tblAcc_Moein_Atf INNER JOIN
                      tblAcc_Tafsili INNER JOIN
                      tblAcc_Tafsili_Atf ON tblAcc_Tafsili.Branch = tblAcc_Tafsili_Atf.Branch AND tblAcc_Tafsili.TafsiliId = tblAcc_Tafsili_Atf.TafsiliId ON 
                      tblAcc_Moein_Atf.AtfID = tblAcc_Tafsili_Atf.AtfId
WHERE     (tblAcc_Tafsili.Branch = @Branch) AND (tblAcc_Moein_Atf.KolID = @KolId) AND (tblAcc_Moein_Atf.MoeinId = @MoeinId) AND (tblAcc_Tafsili.Active = 1)
ORDER BY tblAcc_Tafsili.TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atf_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atf_ByID] (
		 	
				
		@Branch int, 		
		@TafsiliId int, 		
		@AtfId int

		) AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[AtfId]
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId AND 
			[AtfId] = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tSalary_SanadNo]'
GO


Create PROCEDURE [dbo].Update_tSalary_SanadNo
(
	@Branch INT ,
	@SalaryId INT ,
	@sanadNo INT ,
	@SanadDate NVARCHAR(10)
)
AS

	UPDATE dbo.tSalaryM
	SET SanadNo = @sanadNo ,
		SanadDate = @SanadDate
	WHERE Branch = @Branch AND SalaryId = @SalaryId

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_UGroups_ByID]'
GO



-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_UGroups_ByID] (
		 	
				
		@UGroupId tinyint

		) AS
		
		SELECT 
		
		
				[UGroupId],
				[UGroupName]
		
		FROM 
		
		[tblAcc_UGroups]
		
		WHERE
		
		
			[UGroupId] = @UGroupId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_AtfId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_AtfId] (
			
			
			@AtfId int
				
		) AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[AtfId]
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		
		WHERE
		
		
			[AtfId] = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tSalaryMD_Previous]'
GO

CREATE PROCEDURE [dbo].[Get_tSalaryMD_Previous]
(
  @Branch INT  ,
  @AccountYear SMALLINT 
)
AS 

DECLARE @SalaryId INT 
SELECT @SalaryId = MAX(SalaryId) FROM dbo.tSalaryM WHERE  AccountYear = @AccountYear AND Branch = @Branch
SET @SalaryId = ISNULL(@SalaryId , 0)
IF @SalaryId > 0 
SELECT * FROM dbo.tSalaryD
WHERE Branch = @Branch AND SalaryId = @SalaryId
	 
	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Atfs_ByPK_AtfID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Atfs_ByPK_AtfID] (
			
			
			@AtfID int
				
		) AS
		
		SELECT 
		
		
				[AtfID],
				[AtfName],
				[Active]
		
		FROM 
		
		[tblAcc_Atf]
		
		WHERE
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_AsnadGhateiDocumentIds]'
GO



CREATE PROCEDURE [dbo].[Get_All_AsnadGhateiDocumentIds](@AccountYear smallint, @Branch int, @st tinyint) AS
IF (@st = 1)
BEGIN
	SELECT     TOP 100 PERCENT dbo.tblAcc_DocumentHeader.DocumentId, dbo.tblAcc_DocumentHeader.DocumentId2 AS NewDocumentId, dbo.ConvIntToDateFormat(dbo.tblAcc_DocumentHeader.DocumentDate) AS sDocumentDate, t.ct, dbo.tblAcc_DocumentHeader.DocumentDes, 
	                      t.sdBedehkar
	FROM         dbo.tblAcc_DocumentHeader INNER JOIN
	                          (SELECT     AccountYear, Branch, DocumentId, COUNT(DocumentId) AS ct, CASE WHEN SUM(Bedehkar) IS NOT NULL THEN SUM(Bedehkar) ELSE 0 END AS sdBedehkar
	                             FROM tblAcc_DocumentDetail
	                             GROUP BY AccountYear, Branch, DocumentId) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
	WHERE (dbo.tblAcc_DocumentHeader.State = 3) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)
	ORDER BY dbo.tblAcc_DocumentHeader.DocumentDate, dbo.tblAcc_DocumentHeader.DocumentId
END
ELSE
BEGIN
	SELECT     TOP 100 PERCENT dbo.tblAcc_DocumentHeader.DocumentId, dbo.tblAcc_DocumentHeader.DocumentId2 AS NewDocumentId, dbo.ConvIntToDateFormat(dbo.tblAcc_DocumentHeader.DocumentDate) AS sDocumentDate, t.ct, 
				dbo.tblAcc_DocumentHeader.DocumentDes, t.sdBedehkar
	FROM         dbo.tblAcc_DocumentHeader INNER JOIN
	                          (SELECT     AccountYear, Branch, DocumentId, COUNT(DocumentId) AS ct, CASE WHEN SUM(Bedehkar) IS NOT NULL THEN SUM(Bedehkar) ELSE 0 END AS sdBedehkar
	                             FROM tblAcc_DocumentDetail
	                             GROUP BY AccountYear, Branch, DocumentId) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
	WHERE (dbo.tblAcc_DocumentHeader.State = 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)
	ORDER BY dbo.tblAcc_DocumentHeader.DocumentDate, dbo.tblAcc_DocumentHeader.DocumentId
END
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_AtfId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_AtfId_Count] (
			
			
			@AtfId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Tafsili_Atf]
		
		WHERE
		
		
			[AtfId] = @AtfId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_AtfbyKolMoein]'
GO




CREATE PROC [dbo].[Get_AtfbyKolMoein](  @KolId INT,
								@MoeinId INT
								)

AS	
BEGIN
	SELECT DISTINCT [dbo].[tblAcc_Atf].AtfName,
					[dbo].[tblAcc_Atf].AtfId
	FROM [dbo].[tblAcc_Atf] 
--			JOIN [dbo].[tblAcc_Tafsili_Atf] ON dbo.tblAcc_Atf.AtfId = dbo.tblAcc_Tafsili_Atf.AtfId
			JOIN [dbo].[tblAcc_Moein_Atf] ON dbo.tblAcc_Moein_Atf.AtfId =dbo.tblAcc_Atf.AtfId

	WHERE dbo.tblAcc_Moein_Atf.KolId=@KolId--52
			AND dbo.tblAcc_Moein_Atf.MoeinId=@MoeinId--1

END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_PaymentCheckChangeCheckType]'
GO

CREATE proc [dbo].[Update_PaymentCheckChangeCheckType]
(
	@PaymentTypeId tinyint,
	@Date NVARCHAR(10),
	@intSerialNo int
)
as
update tblAcc_PaymentSanad
	set [Vosouli_Date]=(case when @PaymentTypeId=3 then @Date else Vosouli_Date end),
		[Bargashti_Date]=(case when @PaymentTypeId=4 then @Date else Bargashti_Date end),
		[BargashtiMoshtari_Date]=(case when @PaymentTypeId=5 then @Date else BargashtiMoshtari_Date end),
	PaymentTypeId=@PaymentTypeId
	where intSerialNo=@intSerialNo



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Group]'
GO
CREATE TABLE [dbo].[tblAcc_Group]
(
[GroupId] [int] NOT NULL,
[GroupName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Group_GroupName] DEFAULT (''),
[Active] [bit] NOT NULL CONSTRAINT [DF_tblAcc_Group_Active] DEFAULT ((1))
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Group] on [dbo].[tblAcc_Group]'
GO
ALTER TABLE [dbo].[tblAcc_Group] ADD CONSTRAINT [PK_tblAcc_Group] PRIMARY KEY CLUSTERED  ([GroupId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Groups_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Groups_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		GroupID int, GroupName nvarchar(50), Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT GroupID, GroupName, Active
		
	FROM [tblAcc_Group] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT GroupID, GroupName, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_AtfMoein]'
GO
CREATE PROC [dbo].[Update_AtfMoein]
(
@KolId INT,
@MoeinId INT,
@AtfId INT,
@Relation BIT
)
as
begin
DECLARE @count INT
SET @count=(SELECT  COUNT(*) FROM [dbo].[tblAcc_Moein_Atf] WHERE [KolId]=@KolId AND [MoeinId]=@MoeinId AND [AtfId]=@AtfId)

IF (@count=0 AND @Relation=1)
INSERT INTO [dbo].[tblAcc_Moein_Atf]
        ( [KolId], [MoeinId], [AtfId] )
VALUES  ( @KolId, -- KolId - int
          @MoeinId, -- MoeinId - int
          @AtfId  -- AtfId - int
          )
ELSE
DELETE FROM [dbo].[tblAcc_Moein_Atf] 
WHERE [KolId]=@KolId AND [MoeinId]=@MoeinId AND [AtfId]=@AtfId
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_UGroups_ByID_Count]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_UGroups_ByID_Count](@UGroupId tinyint) AS
SELECT COUNT(*) AS ct
FROM tblAcc_UGroups
WHERE UGroupId = @UGroupId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_DocumentId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_DocumentId] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int
				
		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_AtfTafsili]'
GO
Create PROC [dbo].[Update_AtfTafsili]
(
@Branch INT,
@TafsiliId INT,
@AtfId INT,
@Relation BIT
)
as
begin
DECLARE @count INT
SET @count=(SELECT  COUNT(*) FROM [dbo].[tblAcc_Tafsili_Atf] WHERE [TafsiliId]=@TafsiliId AND [AtfId]=@AtfId)

IF (@count=0 AND @Relation=1)
INSERT INTO [dbo].[tblAcc_Tafsili_Atf]
        ( [Branch], [TafsiliId], [AtfId] )
VALUES  ( @Branch, -- Branch - int
          @TafsiliId, -- TafsiliId - int
          @AtfId  -- AtfId - int
          )
ELSE
DELETE FROM [dbo].[tblAcc_Tafsili_Atf]
WHERE [Branch]=@Branch AND [TafsiliId]=@TafsiliId AND [AtfId]=@AtfId
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad]'
GO


CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad]
(
	@CheckNo NVARCHAR(20) ,
	@DateS NVARCHAR(10) ,
	@Price BIGINT  ,
	@Descs NVARCHAR(200) ,
	@DateT NVARCHAR(10) ,
	@intseialNo INT ,
	@Branch INT ,
	@AccountYear SMALLINT ,
	@SanadNo INT ,
	@Result INT OUT 

) 
AS

UPDATE  dbo.tblAcc_RecieveSanad 
SET 	
	CheckNo = @CheckNo ,
	DateS = @DateS,
	Price = @Price ,
	Descs = @Descs ,
	DateT = @DateT 
WHERE intserialNo = @intseialNo

UPDATE dbo.tblAcc_DocumentDetail
	SET Bedehkar = @Price ,
		CheckDate = @DateS ,
		CheckNo = @CheckNo 
		WHERE Branch = @Branch AND AccountYear = @AccountYear AND DocumentId = @SanadNo AND kind = 0
UPDATE dbo.tblAcc_DocumentDetail
	SET Bestankar = @Price ,
		CheckDate = @DateS ,
		CheckNo = @CheckNo 
		WHERE Branch = @Branch AND AccountYear = @AccountYear AND DocumentId = @SanadNo AND kind = 1
		

     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result = 1
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentIdByStateGreaterThan1]'
GO



CREATE PROCEDURE [dbo].[Get_All_DocumentIdByStateGreaterThan1](@AccountYear smallint, @Branch int) AS
SELECT     DocumentId
FROM         tblAcc_DocumentHeader
WHERE     (tblAcc_DocumentHeader.State > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)
ORDER BY DocumentId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Kols_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Kols_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, GroupID int, KolName nvarchar(50), Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, GroupID, KolName, Active
		
	FROM [tblAcc_Kol] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, GroupID, KolName, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_RecieveType]'
GO


CREATE  PROCEDURE [dbo].[Get_tblAcc_RecieveType]
AS
		
	SELECT * FROM dbo.tblAcc_RecieveType
	--WHERE [RecieveTypeId]<=6


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_EmptyForEmptyComboList]'
GO

CREATE PROC [dbo].[Get_EmptyForEmptyComboList]

AS
BEGIN
SELECT [intSerialNo]AS ID
      ,[CheckNo]AS Number
	FROM [tblAcc_PaymentSanad]

WHERE [intSerialNo]=0
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_DocumentId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_DocumentId_Count] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_PayType]'
GO

CREATE PROCEDURE [dbo].[Get_tblAcc_PayType]
AS
		
	SELECT * FROM dbo.tblAcc_PayType WHERE [PaymentTypeId]<=6


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_UGroups]'
GO



-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Update_tblAcc_UGroups] (
				
				
		@UGroupId tinyint, 		
		@UGroupName nvarchar(40)

		
		) AS
		
		UPDATE [tblAcc_UGroups]
		
		SET
		
		
				[UGroupId] = @UGroupId,
				[UGroupName] = @UGroupName

		
		WHERE
		
		
		
			[UGroupId] = @UGroupId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
--PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad_vagozari]'
--GO
--CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad_vagozari]
--(
--	@RecieveTypeId TINYINT,
--	@BankKol INT ,
--	@BankMoein INT ,
--	@BankTafsili INT ,
--	@BankTafsiliName NVARCHAR(50) ,
--	@Darjaryan_Date NVARCHAR(10) ,
--	@intSerialNo INT ,
--	@Descs NVARCHAR(255),
--	@Resid NVARCHAR(255),
--	@Result INT OUT 

--) 
--AS
		
--UPDATE  dbo.tblAcc_RecieveSanad 
--SET 	
--	RecieveTypeId = @RecieveTypeId ,
--	BankKol = @BankKol,
--	BankMoein = @BankMoein ,
--	BankTafsili = @BankTafsili ,
--	BankTafsiliName = @BankTafsiliName ,
--	Darjaryan_Date = @Darjaryan_Date,
--	Resid=@Resid,
--	Descs=@Descs
--WHERE intserialNo = @intSerialNo

--     IF @@ERROR <>0
--        GoTo EventHandler

--    SET @Result =@intSerialNo
--RETURN @Result

--EventHandler:
--    SET @Result = -1
--	RETURN @Result

--GO
--IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
--GO
--IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
--GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_AtfID_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_AtfID_Paged] (
			
			
			@AtfID int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, MoeinId int, AtfID int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, MoeinId, AtfID
		
	FROM [tblAcc_Moein_Atf] 
	
	WHERE
		
		
			[AtfID] = @AtfID	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, MoeinId, AtfID
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_MaxDocumentIdByStateGreaterThan1A]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_MaxDocumentIdByStateGreaterThan1A](
		 	
		@AccountYear smallint, 
		@Branch int,
		@di int

		) AS
		
SELECT     DocumentId, DocumentDate
FROM         tblAcc_DocumentHeader
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId =
                          (SELECT     MAX(DocumentId) AS mv
                             FROM         tblAcc_DocumentHeader
                             WHERE     (state > 1) AND (DocumentId <@di)))



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_KolId_MoeinId]'
GO


-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_KolId_MoeinId] (
			
			
			@KolId int,
			@MoeinId int
				
		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[KolId] = @KolId AND 
			[MoeinId] = @MoeinId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_MoeinId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_MoeinId_Paged] (
			
			
			@KolID int,
			@MoeinId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, MoeinId int, AtfID int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, MoeinId, AtfID
		
	FROM [tblAcc_Moein_Atf] 
	
	WHERE
		
		
			[KolID] = @KolID AND 
			[MoeinId] = @MoeinId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, MoeinId, AtfID
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_AtfTafsiliByTafsili]'
GO

CREATE PROC [dbo].[Get_All_AtfTafsiliByTafsili]
(
@TafsiliId INT
)
as

begin
SELECT [AtfId]
      ,[AtfName]
      ,[Active]
	,(SELECT COUNT(*) FROM [dbo].[tblAcc_Tafsili_Atf] WHERE [TafsiliId]=@TafsiliId  AND [dbo].[tblAcc_Tafsili_Atf] .[AtfId]=[dbo].[tblAcc_Atf].[AtfId])AS Relation
  FROM [tblAcc_Atf]
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad_Vosouli]'
GO
CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad_Vosouli]
(
	@RecieveTypeId TINYINT,
	@Vosouli_Date NVARCHAR(10) ,
	@intSerialNo INT ,
	@Descs NVARCHAR(255),
	@Resid NVARCHAR(255),
	@Result INT OUT 

) 
AS
		
UPDATE  dbo.tblAcc_RecieveSanad 
SET 	
	RecieveTypeId = @RecieveTypeId ,
	Vosouli_Date = @Vosouli_Date,
	Resid=@Resid,
	Descs=@Descs
WHERE intserialNo = @intSerialNo

     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result =@intSerialNo
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccounting]'
GO
CREATE PROCEDURE [dbo].[SetAccessAccounting]
AS
BEGIN

DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccDefineAccounting'
           ,N'تعاريف حسابداري'
           ,'AccDefineAccounting'
           ,1
           ,102)


-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTafsili'
           ,N'تفضيلي ها'
           ,'AccfrmTafsili'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmGroup'
           ,N'گروه ها'
           ,'AccfrmTafsili'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmKol'
           ,N'حساب هاي كل'
           ,'AccfrmKol'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAtfTafsili'
           ,N'عطف هاي تفضيلي'
           ,'AccfrmAtfTafsili'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmMoein'
           ,N'حساب هاي معين'
           ,'AccfrmMoein'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAtfMoein'
           ,N'معين و عطف ها'
           ,'AccfrmAtfMoein'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAccCoding'
           ,N'كد هاي حسابداري'
           ,'AccfrmAccCoding'
           ,1
           ,@id)



-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAtf'
           ,N'عطف ها'
           ,'AccfrmAtf'
           ,1
           ,@id)

INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+8

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Fn_SoodZian]'
GO
SET ANSI_NULLS OFF
GO

CREATE Function [dbo].Fn_SoodZian

(
  @DateBefore INT  ,
  @DateAfter INT  ,
  @AccountYear SMALLINT ,
  @Branch INT 
)

RETURNS  @ReturnTable TABLE(
 TotalSellAmount BIGINT ,
 TotalSellReturnAmount BIGINT ,
 TotalFirstPrice BIGINT ,
 TotalBuyAmount BIGINT ,
 TotalBuyReturnAmount BIGINT ,
 TotalSaleDiscount BIGINT ,
 TotalBuyDiscount BIGINT ,

 TotalCareeFee BIGINT ,
 TotalPacking BIGINT ,

 TotalLosses BIGINT ,
 TotalHoghough BIGINT ,
 TotalHazine BIGINT ,
 TotalHazineMali BIGINT ,
 TotalHazineTozie BIGINT 
)	
As

BEGIN


	DECLARE @TotalSellAmount BIGINT
	DECLARE @TotalSellReturnAmount BIGINT
	DECLARE @TotalFirstPrice BIGINT
	DECLARE @TotalBuyAmount BIGINT
	DECLARE @TotalBuyReturnAmount BIGINT
	DECLARE @TotalSaleDiscount BIGINT
	DECLARE @TotalBuyDiscount BIGINT

	DECLARE @TotalCareeFee BIGINT
	DECLARE @TotalPacking BIGINT

	DECLARE @TotalLosses BIGINT
	DECLARE @TotalHoghough BIGINT
	DECLARE @TotalHazine BIGINT
	DECLARE @TotalHazineMali BIGINT
	DECLARE @TotalHazineTozie BIGINT
	

		Select @TotalSellAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 1  )

		Select @TotalSellReturnAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter 
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 17)

		Select @TotalFirstPrice = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch --AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter 
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 35)

		Select @TotalBuyAmount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter 
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 16)

		Select @TotalBuyReturnAmount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter 
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 18)

		Select @TotalSaleDiscount = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 2)

		Select @TotalBuyDiscount = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 29)

		Select @TotalLosses = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32)

		Select @TotalHoghough = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 13)  --all moein code will be calculate

		Select @TotalHazine = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate

		Select @TotalHazineMali = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 36  )
		--AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 14) --all moein code will be calculate

		Select @TotalHazineTozie = Sum(Bedehkar - Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 37  )
		AND MoeinId <> (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 32) --Losses  moein code calculated in totallosses

		Select @TotalCareeFee = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 4)

		Select @TotalPacking = Sum(-Bedehkar + Bestankar) 
		From tblAcc_DocumentDetail
		Where AccountYear = @AccountYear And Branch = @Branch  AND SaveDate >= @DateBefore AND SaveDate <= @DateAfter
		AND KolID = (Select Kol From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3  )
		AND MoeinId = (Select Moein From dbo.TblAcc_Sale WHERE TblAcc_Sale.Code = 3)

		INSERT INTO @ReturnTable(  TotalSellAmount  , TotalSellReturnAmount  , TotalFirstPrice  , TotalBuyAmount  ,
			TotalBuyReturnAmount  , TotalSaleDiscount  , TotalBuyDiscount  , TotalCareeFee  , TotalPacking  ,
			 TotalLosses  , TotalHoghough  , TotalHazine , TotalHazineMali , TotalHazineTozie )
		VALUES (  @TotalSellAmount  , @TotalSellReturnAmount  , @TotalFirstPrice  , @TotalBuyAmount  ,
			@TotalBuyReturnAmount  , @TotalSaleDiscount  , @TotalBuyDiscount  , @TotalCareeFee  , @TotalPacking  ,
			 @TotalLosses  , @TotalHoghough  , @TotalHazine , @TotalHazineMali , @TotalHazineTozie)
		            


RETURN 


End

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Users]'
GO
SET ANSI_NULLS ON
GO
CREATE TABLE [dbo].[tblAcc_Users]
(
[UserId] [smallint] NOT NULL,
[UserLogin] [nvarchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Users_UserLogin] DEFAULT (''),
[UserPassword] [nvarchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Users_UserPassword] DEFAULT (''),
[UserName] [nvarchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL CONSTRAINT [DF_tblAcc_Users_UserName] DEFAULT (''),
[UGroupId] [tinyint] NOT NULL CONSTRAINT [DF_tblAcc_Users_UGroupId] DEFAULT ((0)),
[Active] [bit] NOT NULL CONSTRAINT [DF_tblAcc_Users_Active] DEFAULT ((0))
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Users] on [dbo].[tblAcc_Users]'
GO
ALTER TABLE [dbo].[tblAcc_Users] ADD CONSTRAINT [PK_tblAcc_Users] PRIMARY KEY CLUSTERED  ([UserId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Users]'
GO


-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------		
	
CREATE PROCEDURE [dbo].[Insert_tblAcc_Users] (
				
		@UserId smallint, 		
		@UserLogin nvarchar(25), 		
		@UserPassword nvarchar(25), 		
		@UserName nvarchar(60), 		
		@UGroupId tinyint, 		
		@Active bit
	) 
	
	AS
		
	INSERT INTO [tblAcc_Users]
		
	(
		[UserId],
		[UserLogin],
		[UserPassword],
		[UserName],
		[UGroupId],
		[Active]
	)		
		
	VALUES		
	(
		@UserId,
		@UserLogin,
		@UserPassword,
		@UserName,
		@UGroupId,
		@Active
	)
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_TafsiliId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_TafsiliId] (
			
			
			@AccountYear smallint,
			@Branch int,
			@TafsiliId int
				
		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_Paged] (
			
			
			@KolID int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, MoeinId int, AtfID int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, MoeinId, AtfID
		
	FROM [tblAcc_Moein_Atf] 
	
	WHERE
		
		
			[KolID] = @KolID	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, MoeinId, AtfID
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TarazSoodZian]'
GO


CREATE   PROCEDURE [dbo].[Get_TarazSoodZian]
    (
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT 
    )
AS 


SELECT ISNULL(TotalSellAmount , 0) AS TotalSellAmount ,
       ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
       ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
       ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
       ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
       ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
       ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
       ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
       ISNULL(TotalPacking , 0) AS TotalPacking ,
       ISNULL(TotalLosses , 0) AS TotalLosses ,
       ISNULL(TotalHoghough , 0) AS TotalHoghough ,
       ISNULL(TotalHazine , 0) AS TotalHazine ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
       
	FROM DBO.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch )
--===============================================

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_MinMaxDocumentId]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_MinMaxDocumentId] (
		 	
		@AccountYear smallint, 
		@Branch int

		) AS
		
		SELECT 
		
		
				CASE WHEN MIN([DocumentId]) IS NULL THEN 1 ELSE MIN([DocumentId]) + 1 END AS DocumentId1, 
				CASE WHEN MAX([DocumentId]) IS NULL THEN 1 ELSE MAX([DocumentId]) + 1 END AS DocumentId2
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_TafsiliId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_TafsiliId_Count] (
			
			
			@AccountYear smallint,
			@Branch int,
			@TafsiliId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingDefineAccountin]'
GO
Create PROCEDURE [dbo].[SetAccessAccountingDefineAccountin]
AS
BEGIN

DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccDefineAccounting'
           ,N'تعاريف حسابداري'
           ,'AccDefineAccounting'
           ,1
           ,102)


-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTafsili'
           ,N'تفضيلي ها'
           ,'AccfrmTafsili'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmGroup'
           ,N'گروه ها'
           ,'AccfrmTafsili'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmKol'
           ,N'حساب هاي كل'
           ,'AccfrmKol'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAtfTafsili'
           ,N'عطف هاي تفضيلي'
           ,'AccfrmAtfTafsili'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmMoein'
           ,N'حساب هاي معين'
           ,'AccfrmMoein'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAtfMoein'
           ,N'معين و عطف ها'
           ,'AccfrmAtfMoein'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAccCoding'
           ,N'كد هاي حسابداري'
           ,'AccfrmAccCoding'
           ,1
           ,@id)



-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAtf'
           ,N'عطف ها'
           ,'AccfrmAtf'
           ,1
           ,@id)

INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+8

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Rep_TarazSoodZian]'
GO


CREATE   PROCEDURE [dbo].[Rep_TarazSoodZian]
    (
      @SystemDate NVARCHAR(20) ,
      @SystemDay NVARCHAR(20) ,
      @SystemTime NVARCHAR(20) ,
      @DateBefore INT  ,
      @DateAfter INT  ,
      @AccountYear SMALLINT ,
      @Branch INT ,
      @MojodiPrice BIGINT 
    )
AS 

    DECLARE @TimeTitle NVARCHAR(10)      
    SET @TimeTitle = N' ساعت : '   

SELECT @SystemDay + ' ' + @SystemDate + ' ' + @TimeTitle + @SystemTime AS SysDay  ,
		SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,1 ,4) + N'/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,5,2) + N'/' + SUBSTRING(CAST(@DateBefore AS NVARCHAR(8)) ,7,2) AS FromDate ,
		SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,1 ,4) + N'/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,5,2) + N'/' + SUBSTRING(CAST(@DateAfter AS NVARCHAR(8)) ,7,2) AS ToDate ,
		@MojodiPrice AS MojodiPrice ,
		ISNULL(TotalSellAmount , 0) AS TotalSellAmount ,
		ISNULL(TotalSellReturnAmount , 0) AS TotalSellReturnAmount ,
		ISNULL(TotalFirstPrice , 0) AS TotalFirstPrice ,
		ISNULL(TotalBuyAmount , 0) AS TotalBuyAmount ,
		ISNULL(TotalBuyReturnAmount , 0) AS TotalBuyReturnAmount ,
		ISNULL(TotalSaleDiscount , 0) AS TotalSaleDiscount ,
		ISNULL(TotalBuyDiscount , 0) AS TotalBuyDiscount ,
		ISNULL(TotalCareeFee , 0) AS TotalCareeFee , 
		ISNULL(TotalPacking , 0) AS TotalPacking ,
		ISNULL(TotalLosses , 0) AS TotalLosses ,
		ISNULL(TotalHoghough , 0) AS TotalHoghough ,
		ISNULL(TotalHazine , 0) AS TotalHazine  ,
       ISNULL(TotalHazineMali , 0) AS TotalHazineMali ,
       ISNULL(TotalHazineTozie , 0) AS TotalHazineTozie 
	FROM dbo.Fn_SoodZian(@DateBefore ,@DateAfter ,@AccountYear ,@Branch )
--===============================================

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Users_UserPassword]'
GO


CREATE PROCEDURE [dbo].[Update_tblAcc_Users_UserPassword](@UserId smallint, @UserPassword nvarchar(25))
AS
	
UPDATE tblAcc_Users
SET UserPassword = @UserPassword
WHERE (UserId = @UserId)


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad_Bargashti]'
GO
CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad_Bargashti]
(
	@RecieveTypeId TINYINT,
	@Bargashti_Date nvarchar(10),
	@intSerialNo INT ,
	@Descs NVARCHAR(255),
	@Resid NVARCHAR(255),
	@Result INT OUT 

) 
AS
		
UPDATE  dbo.tblAcc_RecieveSanad 
SET 	
	RecieveTypeId = @RecieveTypeId ,
	[Bargashti_Date]=@Bargashti_Date,
	Resid =@Resid,
	Descs=@Descs
WHERE intserialNo = @intSerialNo

     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result =@intSerialNo
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessBaseAccountingDefin]'
GO
Create PROCEDURE [dbo].[SetAccessBaseAccountingDefin]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccBaseAccountingDefine'
           ,N'تعاريف پايه حسابداري'
           ,'AccBaseAccountingDefine'
           ,1
           ,102)

-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmBank'
           ,N'بانك ها'
           ,'AccfrmBank'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmBankAccount'
           ,N'حساب هاي بانكي'
           ,'AccfrmBankAccount'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckBook'
           ,N'دسته چك ها'
           ,'AccfrmCheckBook'
           ,1
           ,@id)

INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+4

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_PaymentSanad_SanadNo]'
GO


Create PROCEDURE [dbo].[Update_tblAcc_PaymentSanad_SanadNo]
(
	@intSeriaNo INT ,
	@sanadNo INT ,
	@ItemNo INT ,
	@nvcDate NVARCHAR(10) ,
	@KolTaraf INT ,
	@MoeinTaraf INT ,
	@TafsiliTaraf INT ,
	@TafsiliNameTaraf nvarchar(50)
)
AS


IF @ItemNo = 2 
UPDATE tblAcc_PaymentSanad
	SET Sanad_Pardakhti = @sanadNo ,
		PaymentTypeId = @ItemNo ,
		DateT = @nvcDate ,
		RecKol = @KolTaraf,
		RecMoein = @MoeinTaraf,
		RecTafsili = @TafsiliTaraf,
		Taraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 

IF @ItemNo = 3 
UPDATE tblAcc_PaymentSanad
	SET Sanad_Vosouli = @sanadNo ,
		PaymentTypeId = @ItemNo ,
		Vosouli_Date = @nvcDate 
		WHERE intSerialNo = @intSeriaNo 

ELSE IF @ItemNo = 5 
UPDATE tblAcc_PaymentSanad
	SET Sanad_Bargashti = @sanadNo ,
		PaymentTypeId = @ItemNo ,
		Bargashti_Date = @nvcDate ,
		RecKol = @KolTaraf,
		RecMoein = @MoeinTaraf,
		RecTafsili = @TafsiliTaraf,
		Taraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 

ELSE IF @ItemNo = 6 
UPDATE tblAcc_PaymentSanad
	SET Void = 1 ,
		PaymentTypeId = @ItemNo ,
		Void_Date = @nvcDate ,
		RecKol = NULL ,
		RecMoein = NULL ,
		RecTafsili = NULL ,
		Taraf = NULL 
		WHERE intSerialNo = @intSeriaNo 

ELSE IF @ItemNo = 7 OR @ItemNo = 13
UPDATE tblAcc_PaymentSanad
	SET Sanad_Cash = @sanadNo ,
		PaymentTypeId = @ItemNo ,
		Cash_Date = @nvcDate ,
		RecKol = @KolTaraf,
		RecMoein = @MoeinTaraf,
		RecTafsili = @TafsiliTaraf,
		Taraf = @TafsiliNameTaraf
		WHERE intSerialNo = @intSeriaNo 

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_MinMaxDocumentIdStateEq3]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_MinMaxDocumentIdStateEq3] (
		 	
		@AccountYear smallint, 
		@Branch int

		) AS
		
		SELECT 
		
		
				CASE WHEN MIN([DocumentId]) IS NULL THEN 1 ELSE MIN([DocumentId]) + 1 END AS DocumentId1, 
				CASE WHEN MAX([DocumentId]) IS NULL THEN 1 ELSE MAX([DocumentId]) + 1 END AS DocumentId2
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			State = 3



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_AtfId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_AtfId_Paged] (
			
			
			@AtfId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, TafsiliId int, AtfId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, TafsiliId, AtfId
		
	FROM [tblAcc_Tafsili_Atf] 
	
	WHERE
		
		
			[AtfId] = @AtfId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, TafsiliId, AtfId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_PaymentSanad]'
GO

CREATE PROCEDURE [dbo].[Insert_tblAcc_PaymentSanad]
(
	@CheckNo NVARCHAR(20),
	@DateS NVARCHAR(10),
	@Price BIGINT ,
	@Descs NVARCHAR(255),
	@PaymentTypeId TINYINT,
	@DateT NVARCHAR(10),
	@PayKol INT ,
	@PayMoein INT ,
	@PayTafsili INT ,
	@PayTafsiliName nvarchar(50),
	@Resid NVARCHAR(255),
	@Result INT OUT 

) 
AS
IF @CheckNo=N''
SET @CheckNo=NULL

IF @DateS=N''
SET @DateS=NULL


INSERT INTO [tblAcc_PaymentSanad]
           ([CheckNo]
           ,[DateS]
           ,[Price]
           ,[Descs]
           ,[PaymentTypeId]
           ,[DateT]
           ,[PayKol]
           ,[PayMoein]
           ,[PayTafsili]
           ,[PayTafsiliName]
		   ,[Resid]
			)
     VALUES
           (@CheckNo
           ,@DateS
           ,@Price
           ,@Descs
           ,@PaymentTypeId
           ,@DateT
           ,@PayKol
           ,@PayMoein
           ,@PayTafsili
           ,@PayTafsiliName
		   ,@Resid
           )

		
     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result =@@IDENTITY
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_Sanad]'
GO
CREATE PROC [dbo].[Delete_Sanad]
(
@DocumentId INT,
@Branch INT,
@AccountYear SMALLINT
)
AS
BEGIN
DELETE FROM [dbo].[tblAcc_DocumentDetail] 
WHERE [DocumentId]=@DocumentId AND [Branch]=@Branch AND [AccountYear]=@AccountYear
DELETE FROM [dbo].[tblAcc_DocumentHeader] 
WHERE [DocumentId]=@DocumentId AND [Branch]=@Branch AND [AccountYear]=@AccountYear
END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingReceived]'
GO
Create PROCEDURE [dbo].[SetAccessAccountingReceived]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccReceived'
           ,N'دريافت ها'
           ,'AccReceived'
           ,1
           ,102)

-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmReceivedCash'
           ,N'دريافت وجه نقد'
           ,'AccfrmReceivedCash'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmReceivedCashFromAccount'
           ,N'دريافت وجه نقد از طريق واريز به حساب'
           ,'AccfrmReceivedCashFromAccount'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmReceivedCheck'
           ,N'دريافت چك'
           ,'AccfrmReceivedCheck'
           ,1
           ,@id)

INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+4

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_PaymentCheck_Pardakhti]'
GO


CREATE PROCEDURE [dbo].[Update_PaymentCheck_Pardakhti]
(
	@PaymentTypeId TINYINT,
	@DateT NVARCHAR(10),
	@DateS NVARCHAR(10),
	@Price BIGINT ,
	@PayKol INT ,
	@PayMoein INT ,
	@PayTafsili INT ,
	@PayTafsiliName nvarchar(50),
	@intSerialNo int,
	@Descs NVARCHAR(255),
	@Resid NVARCHAR(255) ,
	@Result INT OUT 
) 
AS

BEGIN TRAN
UPDATE [tblAcc_PaymentSanad]
   SET [DateS] = @DateS
      ,[Price] = @Price
      ,[Descs] = @Descs
      ,[PaymentTypeId] = @PaymentTypeId
      ,[DateT] = @DateT
      ,[PayKol] = @PayKol
      ,[PayMoein] = @PayMoein
      ,[PayTafsili] = @PayTafsili
      ,[PayTafsiliName] = @PayTafsiliName
      ,[Resid] = @Resid
 WHERE intSerialNo=@intSerialNo



     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result = @intSerialNo
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Users]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Update_tblAcc_Users] (
				
				
		@UserId smallint, 		
		@UserLogin nvarchar(25), 		
		@UserPassword nvarchar(25), 		
		@UserName nvarchar(60), 		
		@UGroupId tinyint, 		
		@Active bit

		
		) AS
		
		UPDATE [tblAcc_Users]
		
		SET
		
		
				[UserId] = @UserId,
				[UserLogin] = @UserLogin,
				[UserPassword] = @UserPassword,
				[UserName] = @UserName,
				[UGroupId] = @UGroupId,
				[Active] = @Active

		
		WHERE
		
		
		
			[UserId] = @UserId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Atfs_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Atfs_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, TafsiliId int, AtfId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, TafsiliId, AtfId
		
	FROM [tblAcc_Tafsili_Atf] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, TafsiliId, AtfId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_CheckBook]'
GO
CREATE TABLE [dbo].[tblAcc_CheckBook]
(
[CheckBookId] [int] NOT NULL IDENTITY(1, 1),
[Branch] [int] NULL,
[AccountTafsiliID] [int] NULL,
[StartSerial] [int] NULL,
[EndSerial] [int] NULL,
[PageNumber] [int] NULL,
[Seri] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
[PrintTemplateID] [int] NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_CheckBook] on [dbo].[tblAcc_CheckBook]'
GO
ALTER TABLE [dbo].[tblAcc_CheckBook] ADD CONSTRAINT [PK_tblAcc_CheckBook] PRIMARY KEY CLUSTERED  ([CheckBookId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_PyamentCheck_Print]'
GO


CREATE Proc [dbo].[Get_PyamentCheck_Print]
(
@CheckNo nvarchar(20)
)
As
Begin
Select DateT, 
	   CheckNo,
	   DateS,
	   Price,
	   Descs,
	   PaymentTypeId,
	   Taraf,
	   Resid,
	   BankAccountTafsili,
	   (Select TafsiliName From tblAcc_Tafsili Where TafsiliId=tblAcc_PaymentSanad.BankAccountTafsili)As BankAccountName,
	   (Select Name From tblAcc_ChequePrintTemplate Where PrintTemplateID=(Select PrintTemplateID From tblAcc_CheckBook Where AccountTafsiliID=tblAcc_PaymentSanad.BankAccountTafsili))As PrintTemplateName,
	   (Select [Path] From tblAcc_ChequePrintTemplate Where PrintTemplateID=(Select PrintTemplateID From tblAcc_CheckBook Where AccountTafsiliID=tblAcc_PaymentSanad.BankAccountTafsili))As PrintTemplatePath,
	   dbo.NumberToHarf(Price,0) As PriceRial,
	   dbo.NumberToHarf(Price,1) As PriceTooman
	   
From tblAcc_PaymentSanad
Where CheckNo=@CheckNo

End
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_MinMaxDocumentDateStateEq3]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_MinMaxDocumentDateStateEq3] (
		 	
		@AccountYear smallint, 
		@Branch int

		) AS
		
		SELECT 
		
		
				CASE WHEN MIN([DocumentDate]) IS NOT NULL THEN MIN([DocumentDate]) ELSE 0 END AS DocumentDate1, 
				CASE WHEN MAX([DocumentDate]) IS NOT NULL THEN MAX([DocumentDate]) ELSE 0 END AS DocumentDate2
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			State = 3



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingPayment]'
GO
CREATE PROCEDURE [dbo].[SetAccessAccountingPayment]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccPayment'
           ,N'پرداخت ها'
           ,'AccPayment'
           ,1
           ,102)

-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmPaymentCash'
           ,N'پرداخت وجه نقد'
           ,'AccfrmPaymentCash'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmPaymentCheck'
           ,N'پرداخت چك'
           ,'AccfrmPaymentCheck'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmPayAccountToAccountCash'
           ,N'وجه نقد از حساب به حساب'
           ,'AccfrmPayAccountToAccountCash'
           ,1
           ,@id)



SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmPayAccountToSandoghCash'
           ,N'وجه نقد از حساب به صندوق'
           ,'AccfrmPayAccountToSandoghCash'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmPaySandoghToSandoghCash'
           ,N'وجه نقد از صندوق به صندوق'
           ,'AccfrmPaySandoghToSandoghCash'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmPaySandoghToAccountCash'
           ,N'وجه نقد از صندوق به حساب'
           ,'AccfrmPaySandoghToAccountCash'
           ,1
           ,@id)
INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+7

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_PyamentCheckResid_Print]'
GO



CREATE Proc [dbo].[Get_PyamentCheckResid_Print]
(
@CheckNo nvarchar(20)
)
As
Begin
Select DateT,
	   CheckNo,
	   DateS,
	   Price,
	   Descs,
	   PaymentTypeId,
	   Taraf,
	   Resid,
	   BankAccountTafsili,
	   (Select TafsiliName From tblAcc_Tafsili Where TafsiliId=tblAcc_PaymentSanad.BankAccountTafsili)As BankAccountName,
	   (Select Name From tblAcc_ChequePrintTemplate Where PrintTemplateID=(Select PrintTemplateID From tblAcc_CheckBook Where AccountTafsiliID=tblAcc_PaymentSanad.BankAccountTafsili))As PrintTemplateName,
	   (Select [Path] From tblAcc_ChequePrintTemplate Where PrintTemplateID=(Select PrintTemplateID From tblAcc_CheckBook Where AccountTafsiliID=tblAcc_PaymentSanad.BankAccountTafsili))As PrintTemplatePath,
	   dbo.NumberToHarf(Price,0) As PriceRial,
	   dbo.NumberToHarf(Price,1) As PriceTooman
	   
From tblAcc_PaymentSanad
Where CheckNo=@CheckNo

End
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Userss_Count_ByUserName]'
GO


CREATE PROCEDURE [dbo].[Get_tblAcc_Userss_Count_ByUserName](@UserLogin nvarchar(25))
				
		AS
		
		SELECT
		
			COUNT(*) AS ct
		
		FROM
		
			[tblAcc_Users]
		
		WHERE (UserLogin = @UserLogin)


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_Permissions]'
GO
CREATE TABLE [dbo].[tblAcc_Permissions]
(
[UGroupId] [tinyint] NOT NULL,
[FormId] [tinyint] NOT NULL,
[Show] [bit] NOT NULL CONSTRAINT [DF_tblAcc_Permissions_Show] DEFAULT ((0)),
[Access] [bit] NOT NULL CONSTRAINT [DF_tblAcc_Permissions_Access] DEFAULT ((0))
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_Permissions] on [dbo].[tblAcc_Permissions]'
GO
ALTER TABLE [dbo].[tblAcc_Permissions] ADD CONSTRAINT [PK_tblAcc_Permissions] PRIMARY KEY CLUSTERED  ([UGroupId], [FormId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_PermissionsByUGroupId]'
GO

CREATE PROCEDURE [dbo].[Delete_tblAcc_PermissionsByUGroupId] (
		
				
		@UGroupId tinyint 		
		
		) AS
		
		DELETE [tblAcc_Permissions]
		
		WHERE
		
		
			[UGroupId] = @UGroupId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingCheckReceived]'
GO
create PROCEDURE [dbo].[SetAccessAccountingCheckReceived]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccCheckReceived'
           ,N'چك هاي دريافتني'
           ,'AccCheckReceived'
           ,1
           ,102)

-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAllCheckReceived'
           ,N'نمايش چك ها'
           ,'AccfrmAllCheckReceived'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckReceivedOperationKharj'
           ,N'خرج چك'
           ,'AccfrmCheckReceivedOperationKharj'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckReceivedOperationVagozari'
           ,N'واگذاري چك به بانك'
           ,'AccfrmCheckReceivedOperationVagozari'
           ,1
           ,@id)



SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckReceivedOperationVosouli'
           ,N'وصولي چك'
           ,'AccfrmCheckReceivedOperationVosouli'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckReceivedOperationBargashti'
           ,N'برگشتي چك'
           ,'AccfrmCheckReceivedOperationBargashti'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckReceivedOperationBargashtiMoshtari'
           ,N'برگشت به مشتري چك'
           ,'AccfrmCheckReceivedOperationBargashtiMoshtari'
           ,1
           ,@id)
INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+7

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_KolShenaseh]'
GO
CREATE TABLE [dbo].[tblAcc_KolShenaseh]
(
[ShenaseId] [int] NOT NULL,
[ShenaseName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_KolShenaseh] on [dbo].[tblAcc_KolShenaseh]'
GO
ALTER TABLE [dbo].[tblAcc_KolShenaseh] ADD CONSTRAINT [PK_tblAcc_KolShenaseh] PRIMARY KEY CLUSTERED  ([ShenaseId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_KolType]'
GO


Create Proc [dbo].[Get_KolType]
As
begin
Select * From tblAcc_KolShenaseh

End

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Insert_tblAcc_Sale]'
GO


ALTER PROCEDURE [dbo].[Insert_tblAcc_Sale]( 
					@Code INT,
					@Description nvarchar(50),
					@Kol int,
					@Moein int,
					@Tafsili int,
					@Active BIT ,
					@MoeinDesc NVARCHAR(50)
		
				)
 
 AS

begin Tran
insert into dbo.tblAcc_Sale (
			Code ,
			[Description] ,
			Kol,
			Moein,
			Tafsili ,
			Active ,
			MoeinDesc
)
values(
			@Code ,
			@Description ,
			@Kol,
			@Moein,
			@Tafsili ,
			@Active ,
			@MoeinDesc
)
if @@Error <> 0 
		GOTO EventHandler	
		


commit Tran

RETURN

EventHandler: 

	ROLLBACK TRAN

	RETURN -1





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeaders_ByPK_AccountYear_Branch_DocumentId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeaders_ByPK_AccountYear_Branch_DocumentId] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int
				
		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[DocumentDate],
				[DocumentDes],
				[State],
				[DocumentId2],
				[DocumentKind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_KolId_MoeinId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_KolId_MoeinId_Count] (
			
			
			@KolId int,
			@MoeinId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[KolId] = @KolId AND 
			[MoeinId] = @MoeinId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingCheckPayment]'
GO
create PROCEDURE [dbo].[SetAccessAccountingCheckPayment]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccCheckPayment'
           ,N'چك هاي پرداختني'
           ,'AccCheckPayment'
           ,1
           ,102)

-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAllCheckPayment'
           ,N'نمايش چك ها'
           ,'AccfrmAllCheckPayment'
           ,1
           ,@id)


-------------------------

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckPaymentOperationVosouli'
           ,N'وصولي چك'
           ,'AccfrmCheckPaymentOperationVosouli'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckPaymentOperationBargashtiMoshtari'
           ,N'برگشت از مشتري چك'
           ,'AccfrmCheckPaymentOperationBargashtiMoshtari'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckPaymentEbtal'
           ,N'ابطال چك'
           ,'AccfrmCheckPaymentEbtal'
           ,1
           ,@id)
INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+5

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Kols]'
GO


CREATE PROCEDURE [dbo].[Get_All_tblAcc_Kols]
				
		AS
		
		SELECT 
		
				
				[KolID],
				[GroupID],
				(SELECT [GroupName] FROM [dbo].[tblAcc_Group] WHERE [GroupId]=tblAcc_Kol.[GroupId])AS GroupName,
				[KolName],
				[Active],
				ShenaseID As Shenase
		
		FROM 
		
		[tblAcc_Kol]

		ORDER BY KolId
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Userss_MaxUserIdPlus1]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_Userss_MaxUserIdPlus1]
				
		AS
		
		SELECT
		
			CASE WHEN MAX(UserId) IS NULL THEN 1 ELSE MAX(UserId) + 1 END AS mv
		
		FROM
		
			[tblAcc_Users]




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Permissionss_ByUGroupId]'
GO

CREATE PROCEDURE [dbo].[Get_All_tblAcc_Permissionss_ByUGroupId](@UGroupId tinyint)
				
		AS
		
		SELECT 
		
		
				[UGroupId],
				[FormId],
				[Show],
				[Access]
		
		FROM 
		
		[tblAcc_Permissions]
		
		WHERE (UGroupId = @UGroupId)



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingAsnad]'
GO
create PROCEDURE [dbo].[SetAccessAccountingAsnad]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccfrmAsnad'
           ,N'اسناد حسابداري'
           ,'AccfrmAsnad'
           ,1
           ,102)


INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Kol]'
GO


CREATE PROCEDURE [dbo].[Insert_tblAcc_Kol] (
				
		@KolID int, 		
		@GroupID int, 		
		@KolName nvarchar(50), 		
		@Active bit,
		@ShenaseID int
	) 
	
	AS
		
	INSERT INTO [tblAcc_Kol]
		
	(
		[KolID],
		[GroupID],
		[KolName],
		[Active],
		[ShenaseID]
	)		
		
	VALUES		
	(
		@KolID,
		@GroupID,
		@KolName,
		@Active,
		@ShenaseID
	)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeaders_ByPK_Branch_DocumentId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeaders_ByPK_Branch_DocumentId] (
			
			
			@Branch int,
			@DocumentId int
				
		) AS
		
		SELECT 
		
		
				[Branch],
				[DocumentId],
				[DocumentDate],
				[DocumentDes],
				[State],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_ValidationDocument]'
GO



CREATE PROCEDURE [dbo].[Get_ValidationDocument](@AccountYear smallint, @Branch int, @DocumentId int) AS
	SELECT     CASE WHEN d = s THEN 0 ELSE 1 END AS ct
	FROM         (SELECT     CASE WHEN SUM(Bedehkar) IS NULL THEN 0 ELSE SUM(Bedehkar) END AS d, CASE WHEN SUM(Bestankar) IS NULL 
	                                              THEN 0 ELSE SUM(Bestankar) END AS s
	                        FROM         tblAcc_DocumentDetail
	                        WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId)) t
	UNION ALL
	SELECT     COUNT(*) AS ct
	FROM         tblAcc_DocumentDetail
	WHERE     (NOT EXISTS
	                          (SELECT     *
	                             FROM         [dbo].[tblAcc_Moein]
	                             WHERE     tblAcc_Moein.KolId = tblAcc_DocumentDetail.KolId AND tblAcc_Moein.MoeinId = tblAcc_DocumentDetail.MoeinId)) AND (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId)
	UNION ALL
	SELECT     COUNT(*) AS ct
	FROM         tblAcc_DocumentDetail
	WHERE     (NOT EXISTS
	                          (SELECT     *
	                             FROM         [dbo].[tblAcc_Tafsili]
	                             WHERE     tblAcc_Tafsili.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_Tafsili.TafsiliId = tblAcc_DocumentDetail.TafsiliId)) AND (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId)



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[SetAccessAccountingReport]'
GO
CREATE PROCEDURE [dbo].[SetAccessAccountingReport]
AS
BEGIN
DECLARE @id INT
SET @id=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])
INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id
           ,'AccReport'
           ,N'گزارشات'
           ,'AccReport'
           ,1
           ,102)

-------------------------
DECLARE @id1 INT
SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmKartHesabReport'
           ,N'كارت حساب'
           ,'AccfrmKartHesabReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTafsiliReport'
           ,N'حساب هاي تفضيلي'
           ,'AccfrmTafsiliReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmMoeinReport'
           ,N'حساب هاي معين'
           ,'AccfrmMoeinReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmKolReport'
           ,N'حساب هاي كل'
           ,'AccfrmKolReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmKolReport'
           ,N'حساب هاي كل'
           ,'AccfrmKolReport'
           ,1
           ,@id)



SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTarazTafsiliReport'
           ,N'تراز حساب هاي تفضيلي'
           ,'AccfrmTarazTafsiliReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTarazMoeinReport'
           ,N'تراز حساب هاي معين'
           ,'AccfrmTarazMoeinReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTarazKolReport'
           ,N'تراز حساب هاي كل'
           ,'AccfrmTarazKolReport'
           ,1
           ,@id)



SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmDaftarKolReport'
           ,N'دفتر كل'
           ,'AccfrmDaftarKolReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmDaftarMoeinReport'
           ,N'دفتر معين'
           ,'AccfrmDaftarMoeinReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmJaoftadegiSanadNoReport'
           ,N'جا افتادگي در شماره اسناد'
           ,'AccfrmJaoftadegiSanadNoReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmJaoftadegiTarikhReport'
           ,N'جا افتادگي در تاريخ اسناد'
           ,'AccfrmJaoftadegiTarikhReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmDaftarKolRizReport'
           ,N'ريز دفتر كل'
           ,'AccfrmDaftarKolRizReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmDaftarRuznameReport'
           ,N'دفتر روزنامه'
           ,'AccfrmDaftarRuznameReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmDaftarRuznameRizReport'
           ,N'ريز دفتر روزنامه'
           ,'AccfrmDaftarRuznameRizReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAsnadSummaryReport'
           ,N'گزارش خلاصه اسناد'
           ,'AccfrmAsnadSummaryReport'
           ,1
           ,@id)


SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmTafsiliNoGardeshReport'
           ,N'تفضيلي هاي بدون گردش'
           ,'AccfrmTafsiliNoGardeshReport'
           ,1
           ,@id)



SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmAsnadNoTarazReport'
           ,N'اسناد تراز نشده'
           ,'AccfrmAsnadNoTarazReport'
           ,1
           ,@id)



SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckPaymentSarresidReport'
           ,N'چك هاي پرداختني سررسيد'
           ,'AccfrmCheckPaymentSarresidReport'
           ,1
           ,@id)

SET @id1=(SELECT ISNULL(MAX([intObjectCode]),0)+1 FROM [dbo].[tObjects])

INSERT INTO [tObjects]
           ([intObjectCode]
           ,[ObjectId]
           ,[ObjectName]
           ,[objectLatinName]
           ,[intObjectType]
           ,[ObjectParent])
     VALUES
           (@id1
           ,'AccfrmCheckReceivedSarresidReport'
           ,N'چك هاي دريافتني سررسيد'
           ,'AccfrmCheckReceivedSarresidReport'
           ,1
           ,@id)

INSERT INTO [tAccess_Object]
           ([intAccessLevel]
           ,[intObjectCode])
SELECT 1,[intObjectCode] FROM [dbo].[tObjects]
WHERE [intObjectCode]>=@id AND [intObjectCode]<=@id+20

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Kol]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_Kol] (
	@KolID int, 		
	@GroupID int, 		
	@KolName nvarchar(50), 		
	@Active bit,
	@ShenaseID int



) AS

UPDATE [tblAcc_Kol]

SET
		[KolID] = @KolID,
		[GroupID] = @GroupID,
		[KolName] = @KolName,
		[Active] = @Active,
		[ShenaseID]=@ShenaseID

WHERE
	[KolID] = @KolID


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Users_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Users_ByID] (
		 	
				
		@UserId smallint

		) AS
		
		SELECT 
		
		
				[UserId],
				[UserLogin],
				[UserPassword],
				[UserName],
				[UGroupId],
				[Active]
		
		FROM 
		
		[tblAcc_Users]
		
		WHERE
		
		
			[UserId] = @UserId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_PaymentCheckByType]'
GO
CREATE PROC [dbo].[Get_PaymentCheckByType](
@PaymentType INT,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) 
)
AS
BEGIN
SELECT *,
	(SELECT [PaymentTypeName] FROM [dbo].[tblAcc_PayType] WHERE [dbo].[tblAcc_PayType].[PaymentTypeId]=[dbo].[tblAcc_PaymentSanad].PaymentTypeId)AS PaymentTypeName
FROM [dbo].[tblAcc_PaymentSanad]

where ((@PaymentType=3 AND PaymentTypeId=2) OR
	  (@PaymentType=3 AND PaymentTypeId=4) OR
	  (@PaymentType=4 AND PaymentTypeId=2) OR
	  (@PaymentType=5 AND PaymentTypeId=2) OR
	  (@PaymentType=5 AND PaymentTypeId=4) OR
	  (@PaymentType=6 AND PaymentTypeId<6))
AND [DateT]>=@FromDate AND [DateT]<=@ToDate

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_Branch_TafsiliId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_Atfs_ByFK_Branch_TafsiliId_Paged] (
			
			
			@Branch int,
			@TafsiliId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, TafsiliId int, AtfId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, TafsiliId, AtfId
		
	FROM [tblAcc_Tafsili_Atf] 
	
	WHERE
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, TafsiliId, AtfId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_CheckBook]'
GO


CREATE PROC [dbo].[Insert_CheckBook]
(
@Branch INT,
@BankAccountId INT,
@BankAccountName NVARCHAR(50),
@StartSerial INT,
@PageNumber INT,
@Seri NVARCHAR(50),
@PrintTemplateID INT ,
@Result INT OUT 
)
AS

DECLARE @identity INT
INSERT INTO [tblAcc_CheckBook]
           ([Branch]
           ,[AccountTafsiliID]
           ,[StartSerial]
           ,[EndSerial]
           ,[PageNumber]
           ,[Seri]
           ,[PrintTemplateID])
     VALUES
           (@Branch
           ,@BankAccountId
           ,@StartSerial
           ,@StartSerial+@PageNumber
           ,@PageNumber
           ,@Seri
           ,@PrintTemplateID)

     IF @@ERROR <>0
        GoTo EventHandler

	SET @identity=@@IDENTITY


	DECLARE @I INT
	SET @I=0

	WHILE @I<@PageNumber
	BEGIN
	
	INSERT INTO [tblAcc_PaymentSanad]
           ([CheckNo]
           ,[DateS]
           ,[Price]
           ,[Descs]
           ,[BankAccountTafsili]
           ,[PaymentTypeId]
           ,[Void]
           ,CheckBookId)
     VALUES
           (@StartSerial+@I--, nvarchar(20),>
           ,N''--<DateS, nvarchar(10),>
           ,0--<Price, bigint,>
           ,N'چك خام'--<Descs, nvarchar(255),>
           ,@BankAccountId--<BankAccountTafsili, int,>
           ,1--<PaymentTypeId, int,>
           ,0--<Void, bit,>)
           ,@identity)
	SET @I=@I+1
 
	END 

     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result = @identity
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_DocumentHeader]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_DocumentHeader] (
		
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int
		
		) AS
		
		DELETE [tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_PermissionsShow]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_PermissionsShow] (
				
				
		@UGroupId tinyint, 		
		@FormId tinyint, 		
		@Show bit 		

		
		) AS
		
		UPDATE [tblAcc_Permissions]
		
		SET
		
		
				[Show] = @Show

		
		WHERE
		
		
		
			[UGroupId] = @UGroupId AND 
			[FormId] = @FormId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_CheckBook]'
GO

CREATE  PROCEDURE [dbo].[Delete_tblAcc_CheckBook] (
	@Branch int, 		
	@CheckBookId INT ,
	@Result INT OUT 
	)
 AS
BEGIN TRAN

DECLARE @Count INT 
SELECT @Count = COUNT(*) FROM tblacc_paymentSanad WHERE CheckBookId = @CheckBookId AND PaymentTypeId <> 1 
SET @Count = ISNULL(@Count ,0)

SET @Result = -1
IF @Count = 0
BEGIN
	DELETE FROM tblacc_paymentSanad 
		WHERE	checkBookId = @CheckBookId
	DELETE FROM [dbo].[tblAcc_CheckBook]
		WHERE	[Branch] = @Branch AND checkBookId = @CheckBookId

SET @Result = 1
END 

COMMIT TRAN
RETURN @Result


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Userss_Count_ByUserIdPassword]'
GO


CREATE PROCEDURE [dbo].[Get_tblAcc_Userss_Count_ByUserIdPassword](@UserId smallint, @UserPassword nvarchar(25))
				
		AS
		
		SELECT
		
			COUNT(*) AS ct
		
		FROM
		
			[tblAcc_Users]
		
		WHERE (UserId = @UserId) AND (UserPassword = @UserPassword) AND (Active = 1)


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetail_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetail_ByID] (
		 	
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@RowId int

		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId AND 
			[RowId] = @RowId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_CheckBook]'
GO

CREATE PROC [dbo].[Get_All_CheckBook]
AS
BEGIN
SELECT CheckBookId
      ,[Branch]
      ,[AccountTafsiliID]
	  ,(SELECT [TafsiliName] FROM [dbo].[tblAcc_Tafsili] WHERE [TafsiliId]=[AccountTafsiliID] )AS AccountName
      ,[StartSerial]
      ,[EndSerial]
      ,[PageNumber]
      ,[Seri]
      ,[PrintTemplateID]
	  ,(SELECT [Name] FROM [dbo].[tblAcc_ChequePrintTemplate] WHERE [PrintTemplateID]=[dbo].[tblAcc_CheckBook].[PrintTemplateID])AS PrintTemplateName
	  
  FROM [tblAcc_CheckBook]
ORDER BY CheckBookId
end

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_CountDocumentIdByStateEq1]'
GO


CREATE PROCEDURE [dbo].[Get_CountDocumentIdByStateEq1](@AccountYear smallint, @Branch int) AS
SELECT     COUNT(tblAcc_DocumentHeader.DocumentId) AS ct
FROM         tblAcc_DocumentHeader
WHERE     (tblAcc_DocumentHeader.State = 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Permissionss_Paged]'
GO

-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Permissionss_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		UGroupId tinyint, FormId tinyint, Show bit, Access bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT UGroupId, FormId, Show, Access
		
	FROM [tblAcc_Permissions] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT UGroupId, FormId, Show, Access
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tSalaryMD]'
GO

CREATE PROCEDURE [dbo].[Insert_tSalaryMD]
(
  @Branch INT  ,
  @AccountYear SMALLINT ,
  @Month INT  ,
  @UserId INT ,
  @DocumentDesc NVARCHAR(100) ,
  @st NVARCHAR(4000) ,
  @nvcAddDate NVARCHAR(8) ,
  @SalaryId INT OUT 
)
AS 
BEGIN TRAN
INSERT INTO dbo.tSalaryM
        ( 
          Branch ,
          AccountYear ,
          [Month] ,
          UserId ,
          DocumentDesc ,
          nvcAddDate
        )
VALUES  ( 
		  @Branch  ,
		  @AccountYear ,
		  @Month  ,
		  @UserId ,
		  @DocumentDesc ,
          @nvcAddDate  -- nvcAddDate - nvarchar(10)
        )
IF @@Error <> 0 GOTO EventHandler

SET @SalaryId = @@identity
IF @st <> N''
BEGIN 
	INSERT INTO dbo.tSalaryD
			( SalaryId ,
			  Branch ,
			  Ppno ,
			  Tafsili ,
			  DastmozdRooz ,
			  KarkardRooz ,
			  KarkardMah ,
			  FeeEzafe ,
			  SaatEzafe ,
			  KarkardEzafe ,
			  BimeShakhs ,
			  MaliatShakhs ,
			  Kosourat ,
			  NetKarkardMah ,
			  BimeKarfarma ,
			  MaliatKarfarma
			)
	SELECT    @SalaryId ,
			  @Branch ,
			  Ppno ,
			  Tafsili ,
			  DastmozdRooz ,
			  KarkardRooz ,
			  KarkardMah ,
			  FeeEzafe ,
			  SaatEzafe ,
			  KarkardEzafe ,
			  BimeShakhs ,
			  MaliatShakhs ,
			  Kosourat ,
			  NetKarkardMah ,
			  BimeKarfarma ,
			  MaliatKarfarma
			FROM dbo.Split_Salary(@st)
END 

COMMIT TRAN
RETURN @salaryId

EventHandler:

ROLLBACK TRAN
SET @salaryId = -1
RETURN @salaryId
        
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Userss]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Userss]
				
		AS
		
		SELECT 
		
		
				[UserId],
				[UserLogin],
				[UserPassword],
				[UserName],
				[UGroupId],
				[Active]
		
		FROM 
		
		[tblAcc_Users]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_PermissionsAccess]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_PermissionsAccess] (
				
				
		@UGroupId tinyint, 		
		@FormId tinyint, 		
		@Access bit

		
		) AS
		
		UPDATE [tblAcc_Permissions]
		
		SET
		
		
				[Access] = @Access

		
		WHERE
		
		
		
			[UGroupId] = @UGroupId AND 
			[FormId] = @FormId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_AccountYears]'
GO
CREATE TABLE [dbo].[tblAcc_AccountYears]
(
[AccountYear] [smallint] NOT NULL,
[UserId] [smallint] NOT NULL CONSTRAINT [DF_tblAcc_AccountYears_UserId] DEFAULT ((0))
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_AccountYears] on [dbo].[tblAcc_AccountYears]'
GO
ALTER TABLE [dbo].[tblAcc_AccountYears] ADD CONSTRAINT [PK_tblAcc_AccountYears] PRIMARY KEY CLUSTERED  ([AccountYear])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_AccountYearss_Paged]'
GO


-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_AccountYearss_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AccountYear smallint, UserId smallint
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AccountYear, UserId
		
	FROM [tblAcc_AccountYears] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AccountYear, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Cheque_Payment]'
GO


--براي صفحه نمايش هست 
CREATE  PROCEDURE [dbo].[Get_All_Cheque_Payment]
(
@SystemDate NVARCHAR(10) ,
@SystemDay AS NVARCHAR(20) ,
@SystemTime NVARCHAR(5) ,
@PaymentTypeId TINYINT ,
@BankTafsili INT ,
@OrderView INT ,
@AscDesc INT ,
@ChequeType NVARCHAR(20) ,
@AccountDesc NVARCHAR(50) ,
@OrderDesc NVARCHAR(20) ,
@SortDesc NVARCHAR(20)
)
AS

		
IF @OrderView = 0 OR @OrderView = 3 OR @OrderView = 4
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId]<=6 AND [CheckNo] is NOT null			
		--AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			WHEN 4 THEN BankAccountTafsili
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId]<=6 AND [CheckNo] is NOT null			
			--AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 0 THEN intSerialNo
			WHEN 3 THEN CheckNo
			WHEN 4 THEN BankAccountTafsili
			END DESC 
		END 
END 
ELSE IF @OrderView = 1 OR @OrderView = 2 OR @OrderView = 5
BEGIN 
	IF @AscDesc = 0  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId]<=6 AND [CheckNo] is NOT null			
			--AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 5 THEN Taraf
			END ASC 
		END 
	ELSE IF @AscDesc = 1  
		BEGIN 
		SELECT @SystemDay + N' ' + @SystemDate + N' ساعت ' + @SystemTime AS [SysDay] ,
		tblAcc_PaymentSanad.* , PaymentTypeName  
		--,@FromDate AS FromDate ,@ToDate AS ToDate 
		, @ChequeType AS ChequeType,@AccountDesc AS AccountDesc , @OrderDesc AS OrderDesc , @SortDesc AS SortDesc
		FROM dbo.tblAcc_PaymentSanad
		INNER JOIN dbo.tblAcc_PayType ON dbo.tblAcc_PaymentSanad.PaymentTypeId = dbo.tblAcc_PayType.PaymentTypeId

		WHERE 
			 (tblAcc_PaymentSanad.PaymentTypeId = @PaymentTypeId OR @PaymentTypeId = 0)
			AND (tblAcc_PaymentSanad.BankAccountTafsili = @BankTafsili OR @BankTafsili = 0)
			AND [tblAcc_PaymentSanad].[PaymentTypeId]<=6 AND [CheckNo] is NOT null			
			--AND (tblAcc_PaymentSanad.DateS >= @FromDate AND  tblAcc_PaymentSanad.DateS <= @ToDate)

		ORDER BY CASE @OrderView
			WHEN 1 THEN DateS
			WHEN 2 THEN DateT
			WHEN 5 THEN Taraf
			END DESC 
		END 
END 


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentDetails_ForAll]'
GO


CREATE  PROCEDURE [dbo].[Get_All_tblAcc_DocumentDetails_ForAll](@AccountYear smallint, @Branch int, @DocumentId int)
				
		AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE 
		
			[AccountYear] = @AccountYear AND
			[Branch] = @Branch AND
			[DocumentId] = @DocumentId
		
		ORDER BY 
		
			[RowId] --[kind], [KolId], [MoeinId], [TafsiliId], [Bedehkar], [Bestankar]





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetail_MaxRow]'
GO




CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetail_MaxRow] (
		 	
				
		@AccountYear smallint,
		@Branch int,
		@DocumentId int

		) AS
		
		SELECT 
		
		
				CASE WHEN MAX([RowId]) IS NULL THEN 1 ELSE MAX([RowId]) + 1 END AS ms
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND
			[Branch] = @Branch AND
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Userss_Count_ByUserNamePassword]'
GO



CREATE PROCEDURE [dbo].[Get_All_tblAcc_Userss_Count_ByUserNamePassword](@UserLogin nvarchar(25), @UserPassword nvarchar(25))
				
		AS
		
		SELECT
		
			COUNT(*) AS ct
		
		FROM
		
			[tblAcc_Users]
		
		WHERE (UserLogin = @UserLogin) AND (UserPassword = @UserPassword) AND (Active = 1)

		SELECT
		
			*
		
		FROM
		
			[tblAcc_Users]
		
		WHERE (UserLogin = @UserLogin) AND (UserPassword = @UserPassword) AND (Active = 1)



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_Accounts]'
GO

Create PROC [dbo].[Get_Accounts]( @AtfId INT )
AS	
    BEGIN
        SELECT DISTINCT
                dbo.tblAcc_Tafsili.TafsiliName ,
                dbo.tblAcc_Tafsili.TafsiliId
		FROM    dbo.tblAcc_Tafsili
				JOIN dbo.tblAcc_Tafsili_Atf ON dbo.tblAcc_Tafsili.TafsiliId = dbo.tblAcc_Tafsili_Atf.TafsiliId 
		WHERE dbo.tblAcc_Tafsili_Atf.AtfId=@AtfId AND [dbo].[tblAcc_Tafsili].[TafsiliId]<>0 --

    END
---------------------------------------------------------
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Userss_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Userss_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Users]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Permissions]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Update_tblAcc_Permissions] (
				
				
		@UGroupId tinyint, 		
		@FormId tinyint, 		
		@Show bit, 		
		@Access bit

		
		) AS
		
		UPDATE [tblAcc_Permissions]
		
		SET
		
		
				[UGroupId] = @UGroupId,
				[FormId] = @FormId,
				[Show] = @Show,
				[Access] = @Access

		
		WHERE
		
		
		
			[UGroupId] = @UGroupId AND 
			[FormId] = @FormId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Userss_Count_ByUserName]'
GO



CREATE PROCEDURE [dbo].[Get_All_tblAcc_Userss_Count_ByUserName](@UserLogin nvarchar(25))
				
		AS
		
		SELECT
		
			COUNT(*) AS ct
		
		FROM
		
			[tblAcc_Users]
		
		WHERE (UserLogin = @UserLogin)




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_AccountYear_Branch_DocumentId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_AccountYear_Branch_DocumentId] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int
				
		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_Taf]'
GO
CREATE PROC [dbo].[Get_All_Taf](@Tafsili1 INT)
AS
SELECT *  FROM [dbo].[tblAcc_Tafsili]	
WHERE [TafsiliId]=@Tafsili1
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moeins_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moeins_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, MoeinId int, MoeinName nvarchar(50), Kind tinyint, Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, MoeinId, MoeinName, Kind, Active
		
	FROM [tblAcc_Moein] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, MoeinId, MoeinName, Kind, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Test_Get]'
GO
CREATE PROC [dbo].[Test_Get]
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_Bank]
SELECT *  FROM [dbo].[tGood]


end

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Kol]'
GO



CREATE PROCEDURE [dbo].[Delete_tblAcc_Kol] (
@KolID int
) AS
DELETE [tblAcc_Kol]
WHERE
	[KolID] = @KolID



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Users]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_Users] (
		
				
		@UserId smallint
		
		) AS
		
		DELETE [tblAcc_Users]
		
		WHERE
		
		
			[UserId] = @UserId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Permissionss]'
GO

-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Permissionss]
				
		AS
		
		SELECT 
		
		
				[UGroupId],
				[FormId],
				[Show],
				[Access]
		
		FROM 
		
		[tblAcc_Permissions]


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_CountSanadInPaymentReceived]'
GO
CREATE PROC [dbo].[Get_CountSanadInPaymentReceived]
(
@SanadNo INT
)
AS
BEGIN
DECLARE @PaymentSanad INT
DECLARE @ReceivedSanad INT
SET @PaymentSanad=0
SET @ReceivedSanad=0

SET @PaymentSanad =(SELECT COUNT(*) FROM [dbo].[tblAcc_PaymentSanad]
WHERE 
Sanad_Cash=@SanadNo 
OR 
Sanad_Pardakhti=@SanadNo 
OR 
Sanad_Vosouli=@SanadNo 
OR
Sanad_BargashtiMoshtari=@SanadNo
OR
Sanad_Bargashti=@SanadNo)

SET @ReceivedSanad=(SELECT COUNT(*) FROM [dbo].[tblAcc_RecieveSanad] 
where
Sanad_Daryafti=@SanadNo
OR
Sanad_Vagozari=@SanadNo
OR
Sanad_Vosouli=@SanadNo
OR
Sanad_Kharj=@SanadNo
OR
Sanad_Bargashti=@SanadNo
OR
Sanad_BargashtiMoshtari=@SanadNo
OR
Sanad_Cash=@SanadNo)
SELECT (@PaymentSanad+@ReceivedSanad)

END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_AccountYear_Branch_DocumentId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_AccountYear_Branch_DocumentId_Count] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[AllTafsili]'
GO


CREATE PROCEDURE [dbo].[AllTafsili]
(@Id INT)
AS
BEGIN
	
SELECT     *
		FROM         tblAcc_Tafsili
WHERE [TafsiliId]=@Id
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_CrudeCheque]'
GO

CREATE  PROC [dbo].[Get_All_CrudeCheque]
AS
BEGIN
SELECT [ChequeID]
      ,[Branch]
      ,[AccountTafsiliID]
	  ,(SELECT [TafsiliName] FROM [dbo].[tblAcc_Tafsili] WHERE [TafsiliId]=[AccountTafsiliID] )AS AccountName
      ,[StartSerial]
      ,[EndSerial]
      ,[PageNumber]
      ,[Seri]
      ,[PrintTemplateID]
	  ,(SELECT [Name] FROM [dbo].[tblAcc_ChequePrintTemplate] WHERE [PrintTemplateID]=[dbo].[tblAcc_CrudeCheque].[PrintTemplateID])AS PrintTemplateName
  FROM [tblAcc_CrudeCheque]
ORDER BY chequeID
end

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliByAtf]'
GO
CREATE PROC [dbo].[Get_TafsiliByAtf]( @AtfId INT )
AS	
    BEGIN
        SELECT DISTINCT
                dbo.tblAcc_Tafsili.TafsiliName ,
                dbo.tblAcc_Tafsili.TafsiliId,CAST(dbo.tblAcc_Tafsili.TafsiliId AS NVARCHAR(50))+'--'+  dbo.tblAcc_Tafsili.TafsiliName AS FullName
		FROM    dbo.tblAcc_Tafsili
				JOIN dbo.tblAcc_Tafsili_Atf ON dbo.tblAcc_Tafsili.TafsiliId = dbo.tblAcc_Tafsili_Atf.TafsiliId 
		WHERE dbo.tblAcc_Tafsili_Atf.AtfId=(case when @AtfId=-1 then dbo.tblAcc_Tafsili_Atf.AtfId ELSE @AtfId  end)
UNION ALL 		
        SELECT DISTINCT
                dbo.tblAcc_Tafsili.TafsiliName ,
                dbo.tblAcc_Tafsili.TafsiliId,CAST(dbo.tblAcc_Tafsili.TafsiliId AS NVARCHAR(50))+'--'+  dbo.tblAcc_Tafsili.TafsiliName AS FullName
		FROM    dbo.tblAcc_Tafsili
				JOIN dbo.tblAcc_Tafsili_Atf ON dbo.tblAcc_Tafsili.TafsiliId = dbo.tblAcc_Tafsili_Atf.TafsiliId 
		WHERE dbo.tblAcc_Tafsili_Atf.AtfId=(case when @AtfId=2 then 3 when @AtfId=3 then 2  end)

ORDER BY dbo.tblAcc_Tafsili.TafsiliId

    END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Permissionss_Count]'
GO

-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Permissionss_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Permissions]


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_ByID]'
GO
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_ByID] (
		 	
				
		@Branch int, 		
		@TafsiliId int

		) AS
		
		SELECT 
		
		*
		
		FROM 
		
		[dbo].[tCust]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[Tafsili] = @TafsiliId


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_MaxDocumentIdAndDocumentDateByStateEq3]'
GO


CREATE PROCEDURE [dbo].[Get_MaxDocumentIdAndDocumentDateByStateEq3](@AccountYear smallint, @Branch int) AS
SELECT     MAX(tblAcc_DocumentHeader.DocumentId) AS maxDocumentId, MAX(tblAcc_DocumentHeader.DocumentDate) AS maxDocumentDate
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND 
                      tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE     (tblAcc_DocumentHeader.State = 3) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1]'
GO


CREATE VIEW [dbo].[vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1]
AS
SELECT     dbo.tblAcc_DocumentHeader.AccountYear, dbo.tblAcc_DocumentHeader.Branch, dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, 
                      SUM(dbo.tblAcc_DocumentDetail.Bedehkar) AS sd, SUM(dbo.tblAcc_DocumentDetail.Bestankar) AS ss
FROM         dbo.tblAcc_DocumentHeader INNER JOIN
                      dbo.tblAcc_DocumentDetail ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND 
                      dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
WHERE     (dbo.tblAcc_DocumentHeader.State > 1)
GROUP BY dbo.tblAcc_DocumentHeader.AccountYear, dbo.tblAcc_DocumentHeader.Branch, dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazMoein]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazMoein](@AccountYear smallint, @Branch int, @KolId1 int, @KolId2 int, @MoeinId1 int, @MoeinId2 int, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
IF(@DocumentId2 >0)
BEGIN
	SELECT     *, CASE WHEN sd > ss THEN sd - ss ELSE 0 END AS rd, CASE WHEN ss > sd THEN ss - sd ELSE 0 END AS rs
	FROM         (SELECT     tblAcc_Moein.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName, CASE WHEN SUM(tblAcc_DocumentDetail.Bedehkar) 
	                                              IS NOT NULL THEN SUM(tblAcc_DocumentDetail.Bedehkar) ELSE 0 END AS sd, CASE WHEN SUM(tblAcc_DocumentDetail.Bestankar) IS NOT NULL 
	                                              THEN SUM(tblAcc_DocumentDetail.Bestankar) ELSE 0 END AS ss
	                        FROM         tblAcc_Kol INNER JOIN
	                                              tblAcc_Moein ON tblAcc_Kol.KolId = tblAcc_Moein.KolId INNER JOIN
	                                              tblAcc_DocumentDetail ON tblAcc_Moein.KolId = tblAcc_DocumentDetail.KolId AND tblAcc_Moein.MoeinId = tblAcc_DocumentDetail.MoeinId INNER JOIN
	                                              tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND 
	                                              tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentId <= @DocumentId2)
	                        GROUP BY tblAcc_Moein.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName) t
	ORDER BY KolId, MoeinId
END
ELSE
BEGIN
	IF(@d2 > 0)
	BEGIN
		SELECT     *, CASE WHEN sd > ss THEN sd - ss ELSE 0 END AS rd, CASE WHEN ss > sd THEN ss - sd ELSE 0 END AS rs
		FROM         (SELECT     tblAcc_Moein.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName, CASE WHEN SUM(tblAcc_DocumentDetail.Bedehkar) 
		                                              IS NOT NULL THEN SUM(tblAcc_DocumentDetail.Bedehkar) ELSE 0 END AS sd, CASE WHEN SUM(tblAcc_DocumentDetail.Bestankar) IS NOT NULL 
		                                              THEN SUM(tblAcc_DocumentDetail.Bestankar) ELSE 0 END AS ss
		                        FROM         tblAcc_Kol INNER JOIN
		                                              tblAcc_Moein ON tblAcc_Kol.KolId = tblAcc_Moein.KolId INNER JOIN
		                                              tblAcc_DocumentDetail ON tblAcc_Moein.KolId = tblAcc_DocumentDetail.KolId AND tblAcc_Moein.MoeinId = tblAcc_DocumentDetail.MoeinId INNER JOIN
		                                              tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND 
		                                              tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
		                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate <= @d2)
		                        GROUP BY tblAcc_Moein.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName) t
		ORDER BY KolId, MoeinId
	END
	ELSE
	BEGIN
		SELECT     tblAcc_Kol.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName, vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.sd,
		                      vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.ss,
		                      CASE WHEN vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.sd > vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.ss THEN vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.sd - vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.ss ELSE 0 END AS rd,
		                      CASE WHEN vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.ss > vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.sd THEN vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.ss - vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.sd ELSE 0 END AS rs
		FROM tblAcc_Kol INNER JOIN tblAcc_Moein ON tblAcc_Kol.KolId = tblAcc_Moein.KolId INNER JOIN vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1 ON tblAcc_Moein.KolId = vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.KolId AND tblAcc_Moein.MoeinId = vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.MoeinId
		WHERE     (vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.AccountYear = @AccountYear) AND (vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.Branch = @Branch) AND ((vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.sd <> 0) OR (vw_Get_All_DocumentGroupByKolIdAndMoeinIdByStateGreaterThan1.ss <> 0)) AND (((tblAcc_Moein.KolId = @KolId1) AND (tblAcc_Moein.MoeinId >= @MoeinId1)) OR ((tblAcc_Moein.KolId > @KolId1) AND (tblAcc_Moein.KolId < @KolId2)) OR ((tblAcc_Moein.KolId = @KolId2) AND (tblAcc_Moein.MoeinId <= @MoeinId2)))
		ORDER BY tblAcc_Kol.KolId, tblAcc_Moein.MoeinId
	END
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Permissions]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_Permissions] (
		
				
		@UGroupId tinyint, 		
		@FormId tinyint
		
		) AS
		
		DELETE [tblAcc_Permissions]
		
		WHERE
		
		
			[UGroupId] = @UGroupId AND 
			[FormId] = @FormId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
--PRINT N'Creating [dbo].[Get_All_Cheque]'
--GO

--CREATE  PROCEDURE [dbo].[Get_All_Cheque] 
--(
--@BankNo INT ,
--@RecieveTypeId TINYINT ,
--@BankTafsili INT ,
--@FromDate NVARCHAR(8) ,
--@ToDate NVARCHAR(8) 
--)
--AS
--BEGIN
--		SELECT * FROM [dbo].[tblAcc_RecieveSanad]
	
--		WHERE (tblAcc_RecieveSanad.BankNo = @BankNo OR @BankNo = 0)
--			AND (tblAcc_RecieveSanad.RecieveTypeId = @RecieveTypeId OR @RecieveTypeId = 0)
--			AND (tblAcc_RecieveSanad.BankTafsili = @BankTafsili OR @BankTafsili = 0)
--			AND (tblAcc_RecieveSanad.DateS >= @FromDate AND  tblAcc_RecieveSanad.DateS <= @ToDate OR tblAcc_RecieveSanad.DateS IS NULL )
--			AND [CheckNo] is NOT null	
--		ORDER BY  DateS desc
--END 




--GO
--IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
--GO
--IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
--GO
PRINT N'Creating [dbo].[Get_MinMaxDocumentDateByStateGreaterThan1]'
GO



CREATE PROCEDURE [dbo].[Get_MinMaxDocumentDateByStateGreaterThan1](@AccountYear smallint, @Branch int) AS
SELECT     MIN(tblAcc_DocumentHeader.DocumentDate) AS minDocumentDate, MAX(tblAcc_DocumentHeader.DocumentDate) AS maxDocumentDate
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND 
                      tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE     (tblAcc_DocumentHeader.State > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_CheckPaymentSarresid]'
GO



CREATE  PROC [dbo].[Get_All_CheckPaymentSarresid]
(
@DateS NVARCHAR(10)
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_PaymentSanad]
WHERE [DateS]<=@DateS AND [CheckNo] IS NOT NULL  
AND [PaymentTypeId]=2
ORDER BY DateS
END	

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[GetChequeNumbersByAccountTafsiliID]'
GO

CREATE PROC [dbo].[GetChequeNumbersByAccountTafsiliID]
(
@AccountTafsiliId INT
)
AS
BEGIN
SELECT [intSerialNo]AS ID
      ,[CheckNo]AS Number
	FROM [tblAcc_PaymentSanad]

WHERE [PaymentTypeId]=1 AND [PayTafsili]= @AccountTafsiliId
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_PayAccountToSandoghCash]'
GO
CREATE PROC [dbo].[Get_All_PayAccountToSandoghCash]
(
@PayType int,
@FromDate NVARCHAR(8),
@ToDate NVARCHAR(8)
)
AS
BEGIN
SELECT [intSerialNo],
		[DateT],
		[Sanad_Cash],
		[PayTafsili],
		[RecTafsili],
		[Resid],
		[Price]
 FROM [dbo].[tblAcc_PaymentSanad]
WHERE [PaymentTypeId]=10 AND [DateT]>=@FromDate AND [DateT]<=@ToDate
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1]'
GO


CREATE VIEW [dbo].[vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1]
AS
SELECT     dbo.tblAcc_DocumentHeader.AccountYear, dbo.tblAcc_DocumentHeader.Branch, dbo.tblAcc_DocumentDetail.KolId, SUM(dbo.tblAcc_DocumentDetail.Bedehkar) AS sd, 
                      SUM(dbo.tblAcc_DocumentDetail.Bestankar) AS ss
FROM         dbo.tblAcc_DocumentHeader INNER JOIN
                      dbo.tblAcc_DocumentDetail ON dbo.tblAcc_DocumentHeader.AccountYear = dbo.tblAcc_DocumentDetail.AccountYear AND 
                      dbo.tblAcc_DocumentHeader.Branch = dbo.tblAcc_DocumentDetail.Branch AND dbo.tblAcc_DocumentHeader.DocumentId = dbo.tblAcc_DocumentDetail.DocumentId
WHERE     (dbo.tblAcc_DocumentHeader.State > 1)
GROUP BY dbo.tblAcc_DocumentHeader.AccountYear, dbo.tblAcc_DocumentHeader.Branch, dbo.tblAcc_DocumentDetail.KolId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentGroupByKolIdByStateGreaterThan1]'
GO

CREATE PROCEDURE [dbo].[Get_All_DocumentGroupByKolIdByStateGreaterThan1](@AccountYear smallint, @Branch int) AS
SELECT     tblAcc_Kol.KolId, tblAcc_Kol.KolName, vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1.sd, 
                      vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1.ss, CASE WHEN sd >= ss THEN sd - ss ELSE 0 END AS rd, 
                      CASE WHEN ss >= sd THEN ss - sd ELSE 0 END AS rs
FROM         tblAcc_Kol INNER JOIN
                      vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1 ON 
                      tblAcc_Kol.KolId = vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1.KolId
WHERE     (vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1.AccountYear = @AccountYear) AND 
                      (vw_Get_All_DocumentGroupByKolIdByStateGreaterThan1.Branch = @Branch)
ORDER BY tblAcc_Kol.KolId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Permissions]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------		
	
CREATE PROCEDURE [dbo].[Insert_tblAcc_Permissions] (
				
		@UGroupId tinyint, 		
		@FormId tinyint, 		
		@Show bit, 		
		@Access bit
	) 
	
	AS
		
	INSERT INTO [tblAcc_Permissions]
		
	(
		[UGroupId],
		[FormId],
		[Show],
		[Access]
	)		
		
	VALUES		
	(
		@UGroupId,
		@FormId,
		@Show,
		@Access
	)
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Check_KolExist]'
GO



CREATE PROC [dbo].[Check_KolExist](@kolId INT,
							@result INT OUTPUT)
AS
BEGIN
	SET @result=0
	IF EXISTS(SELECT * FROM dbo.tblAcc_Kol WHERE KolId=@kolId)
	SET @result=1

END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_CheckReceivedSarresid]'
GO




CREATE  PROC [dbo].[Get_All_CheckReceivedSarresid]
(
@DateS NVARCHAR(10)
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_RecieveSanad]
WHERE [DateS]<=@DateS AND [CheckNo] IS NOT NULL 
AND [RecieveTypeId]=1 
ORDER BY DateS
END	

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliIdCountsInChilds]'
GO
CREATE  PROC [dbo].[Get_TafsiliIdCountsInChilds](@branch INT,
	@Tafsili int)
AS
BEGIN
	SELECT  *  
	FROM    dbo.tblAcc_DocumentDetail
	WHERE   TafsiliId = @Tafsili
			AND Branch=@branch
	
	

END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_GroupIdCountsInChilds]'
GO

CREATE PROCEDURE [dbo].[Get_GroupIdCountsInChilds](@GroupId int) AS
SELECT     COUNT(GroupId) AS ct
FROM         tblAcc_Kol
WHERE     (GroupId = @GroupId)


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_PaymentCashSanad]'
GO

CREATE PROC [dbo].[Update_tblAcc_PaymentCashSanad]
(
@intSerialNo INT,
@SanadNo INT
)
AS
BEGIN
UPDATE [dbo].[tblAcc_PaymentSanad]
SET [Sanad_Cash]=@SanadNo WHERE [intSerialNo]=@intSerialNo
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_CrudeCheque]'
GO

CREATE  PROC [dbo].[Insert_CrudeCheque]
(
@Branch INT,
@AccountTafsiliID INT,
@StartSerial INT,
@PageNumber INT,
@Seri NVARCHAR(50),
@PrintTemplateID INT,
@BankKol INT,
@BankMoein int,
@BankTafsiliName NVARCHAR(50),
            
@Result INT OUT

)
AS
BEGIN
BEGIN TRAN
DECLARE @identity INT
INSERT INTO [tblAcc_CrudeCheque]
           ([Branch]
           ,[AccountTafsiliID]
           ,[StartSerial]
           ,[EndSerial]
           ,[PageNumber]
           ,[Seri]
           ,[PrintTemplateID])
     VALUES
           (@Branch
           ,@AccountTafsiliID
           ,@StartSerial
           ,@StartSerial+@PageNumber
           ,@PageNumber
           ,@Seri
           ,@PrintTemplateID)

SET @identity=@@IDENTITY

 IF @@ERROR <> 0 
	GOTO EventHandler

	DECLARE @I INT
	SET @I=0

	WHILE @I<=@PageNumber
	BEGIN
	
INSERT INTO [tblAcc_PaymentSanad]
           ([CheckNo]
           ,[DateS]
           ,[Price]
           ,[Descs]
           ,[PaymentTypeId]
           ,[DateT]
           ,[RecKol]
           ,[RecMoein]
           ,[RecTafsili]
           ,[Taraf]
           ,[PayKol]
           ,[PayMoein]
           ,[PayTafsili]
           ,[PayTafsiliName]
           ,[Void])
     VALUES
           (@StartSerial+@I--<CheckNo, nvarchar(20),>
           ,N''--<DateS, nvarchar(10),>
           ,0--<Price, bigint,>
           ,N''--<Descs, nvarchar(255),>
           ,1--<PaymentTypeId, int,>
           ,N''--<DateT, nvarchar(10),>
           ,NULL--<RecKol, int,>
           ,Null--<RecMoein, int,>
           ,NULL--<RecTafsili, int,>
           ,NULL--<Taraf, nvarchar(50),>
           ,@BankKol--<BankKol, int,>
           ,@BankMoein--<BankMoein, int,>
           ,@AccountTafsiliID--<BankTafsili, int,>
           ,@BankTafsiliName--<BankTafsiliName, nvarchar(50),>   
           ,0--<Void, bit,>)
			)
	SET @I=@I+1
 IF @@ERROR <> 0 
GOTO EventHandler
	END 

 COMMIT TRAN
SET @Result=1
    RETURN @Result

    EventHandler:

    ROLLBACK TRAN
    SET @Result = -1

    RETURN @Result
end

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_PayAccountToAccountCash]'
GO
CREATE PROC [dbo].[Get_All_PayAccountToAccountCash]
(
@PayType INT,
@FromDate NVARCHAR(8),
@ToDate NVARCHAR(8)
)
AS
BEGIN
SELECT [intSerialNo],
		[DateT],
		[Sanad_Cash],
		[PayTafsili],
		[RecTafsili],
		[Resid],
		[Price]
 FROM [dbo].[tblAcc_PaymentSanad]
WHERE [PaymentTypeId]=8 AND [DateT]>=@FromDate AND [DateT]<=@ToDate
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Get_tblAcc_Sale]'
GO


ALTER PROCEDURE [dbo].[Get_tblAcc_Sale]
 AS

SELECT * ,
(SELECT [KolName] FROM [dbo].[tblAcc_Kol] WHERE [KolId] =[dbo].[TblAcc_Sale].[Kol])AS KolName,
(SELECT [TafsiliName]  FROM [dbo].[tblAcc_Tafsili] WHERE [TafsiliId] =[dbo].[TblAcc_Sale].[Tafsili])AS TafsiliName

FROM  tblAcc_Sale



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[TafsiliName]'
GO


CREATE FUNCTION [dbo].[TafsiliName](@Branch int, @TafsiliId int)  
RETURNS nvarchar(50) AS  
BEGIN 
	DECLARE rs CURSOR
	READ_ONLY
	FOR SELECT TafsiliName FROM tblAcc_Tafsili WHERE TafsiliId = @TafsiliId AND Branch = @Branch

	DECLARE @s nvarchar(50)
	DECLARE @r nvarchar(50)

	SET @r = ''
	OPEN rs

	FETCH NEXT FROM rs INTO @s
	WHILE (@@fetch_status <> -1)
	BEGIN
		IF (@@fetch_status <> -2)
		BEGIN
			SET @r = @s
		END
		FETCH NEXT FROM rs INTO @s
	END

	CLOSE rs

	DEALLOCATE rs

	RETURN @r
END




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliName]'
GO

CREATE PROCEDURE [dbo].[Get_TafsiliName](@Branch int, @TafsiliId int) AS
SELECT dbo.TafsiliName(@Branch, @TafsiliId) AS des




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Permissions_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Permissions_ByID] (
		 	
				
		@UGroupId tinyint, 		
		@FormId tinyint

		) AS
		
		SELECT 
		
		
				[UGroupId],
				[FormId],
				[Show],
				[Access]
		
		FROM 
		
		[tblAcc_Permissions]
		
		WHERE
		
		
			[UGroupId] = @UGroupId AND 
			[FormId] = @FormId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tBranchs]'
GO





-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tBranchs]
				
		AS
		
		SELECT 
		
		
				[Branch],
				[nvcBranchName],
				[Type],
				[Active]
		
		FROM 
		
		[tBranch]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_RecievedCashSanad]'
GO

CREATE PROC [dbo].[Update_tblAcc_RecievedCashSanad]
(
@intSerialNo INT,
@SanadNo INT
)
AS
BEGIN
UPDATE [dbo].[tblAcc_RecieveSanad]
SET [Sanad_Cash]=@SanadNo WHERE [intSerialNo]=@intSerialNo
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[fnDocumentKind]'
GO

CREATE FUNCTION [dbo].[fnDocumentKind] (@a tinyint)  
RETURNS nvarchar(25) AS  
BEGIN 
	declare @d nvarchar(25)
	set @d = case @a
		when 1 then N'حسابداري'
		when 2 then N'فروش'
		else N''
	end
	return @d
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_KholasehAsnad2]'
GO

CREATE PROCEDURE [dbo].[Get_All_KholasehAsnad2](@AccountYear smallint, @Branch int, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0, @DocumentId21 int, @DocumentId22 int, @State tinyint, @DocumentKind tinyint, @b1 int, @b2 int) AS
SELECT     dbo.tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(dbo.tblAcc_DocumentHeader.DocumentDate) AS dd, 
                      dbo.tblAcc_DocumentHeader.DocumentDes, 
                      dbo.fnDocumentState(dbo.tblAcc_DocumentHeader.State) AS DocumentState, 
                      dbo.tblAcc_DocumentHeader.DocumentId2, dbo.fnDocumentKind(dbo.tblAcc_DocumentHeader.DocumentKind) AS DocumentKindDes, sdBedehkar, sdBestankar
FROM         tblAcc_DocumentHeader LEFT OUTER JOIN
                      (SELECT AccountYear, Branch, DocumentId, COUNT(DocumentId) AS ct, CASE WHEN SUM(Bedehkar) IS NOT NULL THEN SUM(Bedehkar) ELSE 0 END AS sdBedehkar, CASE WHEN SUM(Bestankar) IS NOT NULL THEN SUM(Bestankar) ELSE 0 END AS sdBestankar FROM tblAcc_DocumentDetail GROUP BY AccountYear, Branch, DocumentId) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
WHERE (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND ((@d2 = 0) OR (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) AND ((@DocumentId2 = 0) OR (tblAcc_DocumentHeader.DocumentId BETWEEN @DocumentId1 AND @DocumentId2))
	AND ((@State = 0) OR (tblAcc_DocumentHeader.State = @State)) AND ((@DocumentKind = 0) OR (tblAcc_DocumentHeader.DocumentKind = @DocumentKind)) AND ((@DocumentId22 = 0) OR (tblAcc_DocumentHeader.DocumentId2 BETWEEN @DocumentId21 AND @DocumentId22)) AND ((@b2 = 0) OR (t.sdBedehkar BETWEEN @b1 AND @b2))
ORDER BY dbo.tblAcc_DocumentHeader.DocumentId


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[vw_TafsiliAll]'
GO


CREATE VIEW [dbo].[vw_TafsiliAll]
AS
SELECT DISTINCT Branch, 0 AS TafsiliId, N'' AS TafsiliName, 1 AS Active
FROM         dbo.tblAcc_Tafsili
UNION ALL
SELECT     Branch, TafsiliId, TafsiliName, Active
FROM         dbo.tblAcc_Tafsili



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazTafsili]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazTafsili](@AccountYear smallint, @Branch int, @KolId1 int= 0, @KolId2 int= 0, @MoeinId1 int= 0, @MoeinId2 int= 0, @TafsiliId int= 0, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
SELECT     KolId, MoeinId, TafsiliId, MAX(KolName) AS KolName, MAX(MoeinName) AS MoeinName, MAX(TafsiliName) AS TafsiliName, SUM(Bedehkar) AS BDAmt, SUM(Bestankar) AS BSAmt, CASE WHEN SUM(RBDA)-SUM(RBSA)>0 THEN SUM(RBDA)-SUM(RBSA) ELSE 0 END AS RBDAmt, 
                      CASE WHEN SUM(RBSA)-SUM(RBDA)>0 THEN SUM(RBSA)-SUM(RBDA) ELSE 0 END AS RBSAmt
FROM         (SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName, vw_TafsiliAll.TafsiliName, tblAcc_DocumentDetail.Bedehkar,
                      tblAcc_DocumentDetail.Bestankar, CASE WHEN tblAcc_DocumentDetail.Bedehkar - tblAcc_DocumentDetail.Bestankar > 0 THEN tblAcc_DocumentDetail.Bedehkar - tblAcc_DocumentDetail.Bestankar
                       ELSE 0 END AS RBDA, CASE WHEN tblAcc_DocumentDetail.Bestankar - tblAcc_DocumentDetail.Bedehkar > 0 THEN tblAcc_DocumentDetail.Bestankar - tblAcc_DocumentDetail.Bedehkar
                      ELSE 0 END AS RBSA
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                      tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                      tblAcc_Kol ON tblAcc_Moein.KolId = tblAcc_Kol.KolId INNER JOIN
                      vw_TafsiliAll ON tblAcc_DocumentDetail.TafsiliId = vw_TafsiliAll.TafsiliId
WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND
	((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2)) AND
	((@MoeinId1 = 0) OR (tblAcc_DocumentDetail.MoeinId BETWEEN @MoeinId1 AND @MoeinId2)) AND
	((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId)) AND
	((@d1 = 0) OR (DocumentDate BETWEEN @d1 AND @d2)) AND
	((@DocumentId1 = 0) OR (tblAcc_DocumentDetail.DocumentId BETWEEN @DocumentId1 AND @DocumentId2))) t
GROUP BY KolId, MoeinId, TafsiliId
ORDER BY KolId, MoeinId, TafsiliId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
--PRINT N'Creating [dbo].[Update_General_ReceivedSanad]'
--GO
--SET QUOTED_IDENTIFIER ON
--GO
--SET ANSI_NULLS ON
--GO
--CREATE proc [dbo].[Update_General_ReceivedSanad]
--(
--@SerialNo int,
--@DateT NVARCHAR(10),
--@DateS NVARCHAR(10),
--@Price BIGINT ,
--@BankNo int,
--@BankAccount nvarchar(50),
--@Resid NVARCHAR(255),
--@Descs nvarchar(255),
--@PayTafsili int,
--@Taraf  NVARCHAR(50) ,
--@DarTafsili int,
--@DarTafsiliName NVARCHAR(50) ,
--@BankTafsili int,
--@BankTafsiliName nvarchar(50),
--@KharjTafsili int,
--@KharjTafsiliName nvarchar(50),
--@Kharj_Date nvarchar(10),
--@Darjaryan_Date nvarchar(10),
--@Vosouli_Date nvarchar(10),
--@Bargashti_Date nvarchar(10),
--@BargashtiMoshtari_Date nvarchar(10),
--@Branch int,
--@SanadNo int,
--@UserID int,
--@AccountYear smallint,
--@DocumentDate int,
--@RowDesc nvarchar(255),
--@TafsiliBedehkar INT,
--@TafsiliBestankar INT,
--@CheckNo NVARCHAR(20)=N''
--)
--as 
--begin
--UPDATE [tblAcc_RecieveSanad]
--   SET [DateS] = (case when @DateS=N'' then DateS else @DateS end)
--	  ,[CheckNo] = (case when @CheckNo=N'' then [CheckNo] else @CheckNo end)
--      ,[Price] = (case when @Price=-1 then [Price] else @Price end)
--      ,[Descs] = (case when @Descs=N'' then [Descs] else @Descs end)
--      ,[DateT] = (case when @DateT=N'' then DateT else @DateT end)
--      ,[BankNo] = (case when @BankNo=-1 then BankNo else @BankNo end)
--      ,[BankAccount] = (case when @BankAccount=N'' then BankAccount else @BankAccount end)

--      ,[PayTafsili] = (case when @PayTafsili=-1 then PayTafsili else @PayTafsili end)
--      ,[Taraf] = (case when @Taraf=N'' then Taraf else @Taraf end)

--      ,[DarTafsili] = (case when @DarTafsili=-1 then DarTafsili else @DarTafsili end)
--      ,[DarTafsiliName] = (case when @DarTafsiliName=N'' then DarTafsiliName else @DarTafsiliName end)

--      ,[BankTafsili] = (case when @BankTafsili=-1 then BankTafsili else @BankTafsili end)
--      ,[BankTafsiliName] = (case when @BankTafsiliName=N'' then BankTafsiliName else @BankTafsiliName end)
--      ,[Darjaryan_Date] = (case when @Darjaryan_Date=N'' then Darjaryan_Date else @Darjaryan_Date end)
--      ,[Vosouli_Date] = (case when @Vosouli_Date=N'' then Vosouli_Date else @Vosouli_Date end)
--      ,[KharjTafsili] =(case when @KharjTafsili=-1 then KharjTafsili else @KharjTafsili end)
--      ,[KharjTafsiliName] = (case when @KharjTafsiliName=N'' then KharjTafsiliName else @KharjTafsiliName end)
--      ,[Kharj_Date] = (case when @Kharj_Date=N'' then Kharj_Date else @Kharj_Date end)
--      ,[Bargashti_Date] = (case when @Bargashti_Date=N'' then Bargashti_Date else @Bargashti_Date end)
--      ,[BargashtiMoshtari_Date] = (case when @BargashtiMoshtari_Date=N'' then BargashtiMoshtari_Date else @BargashtiMoshtari_Date end)

--      ,[Resid] = (case when @Resid=N'' then Resid else @Resid end)
-- WHERE intSerialNo=@SerialNo
 
 
-- UPDATE [tblAcc_DocumentHeader]
--   SET [AccountYear] = @AccountYear
--      ,[Branch] = @Branch
--      ,[DocumentDate] = @DocumentDate
--      ,[SaveDate] = dbo.ShamsiInt(GetDate())
--      ,[UserId] = @UserId
-- WHERE DocumentId=@SanadNo and Branch=@Branch and AccountYear=@AccountYear
 
-- UPDATE [tblAcc_DocumentDetail]
--   SET [AccountYear] = @AccountYear
--      ,[Branch] = @Branch

--      ,[TafsiliId] = @TafsiliBedehkar
--      ,[RowDes] = @RowDesc
--      ,[Bedehkar] = (case when @Price=-1 then [Bedehkar] else @Price end)
--      ,[Bestankar] = 0

--      ,[SaveDate] = dbo.ShamsiInt(GetDate())
--      ,[UserId] = @UserId
--      ,[CheckNo] = (case when @CheckNo=N'' then [CheckNo] else @CheckNo end)
--      ,[CheckDate] = (case when @DateS=N'' then null else cast('13'+replace(@DateS,'/','') as int)  end)
      
-- WHERE DocumentId=@SanadNo and RowId=1 and Branch=@Branch and AccountYear=@AccountYear
 
-- UPDATE [tblAcc_DocumentDetail]
--   SET [AccountYear] = @AccountYear
--      ,[Branch] = @Branch

--      ,[TafsiliId] = @TafsiliBestankar
--      ,[RowDes] = @RowDesc
--      ,[Bedehkar] = 0
--      ,[Bestankar] = (case when @Price=-1 then [Bestankar] else @Price end)

--      ,[SaveDate] = dbo.ShamsiInt(GetDate())
--      ,[UserId] = @UserId
--      ,[CheckNo] = (case when @CheckNo=N'' then [CheckNo] else @CheckNo end)
--      ,[CheckDate] = (case when @DateS=N'' then null else cast('13'+replace(@DateS,'/','') as int) end)
      
-- WHERE DocumentId=@SanadNo and RowId=2 and Branch=@Branch and AccountYear=@AccountYear
--end

--GO
--IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
--GO
--IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
--GO
PRINT N'Creating [dbo].[Gap]'
GO
SET ANSI_NULLS OFF
GO

CREATE FUNCTION [dbo].[Gap] (@AccountYear smallint, @Branch int)  
RETURNS @t TABLE (  
		[A1] [int] NOT NULL,
		[A2] [int] NOT NULL
	)
BEGIN 
	DECLARE @p int
	DECLARE @DocumentId int

	DECLARE cs CURSOR
	KEYSET
	FOR SELECT DocumentId FROM tblAcc_DocumentHeader WHERE AccountYear = @AccountYear AND Branch = @Branch ORDER BY DocumentId

	OPEN cs

	SET @p = 0
	FETCH NEXT FROM cs INTO @DocumentId
	WHILE (@@fetch_status <> -1)
	BEGIN
		IF (@@fetch_status <> -2)
		BEGIN
			IF (@p = 0)
			BEGIN
				SET @p = @DocumentId
			END
			IF (@p <> @DocumentId)
			BEGIN
				INSERT INTO @t(A1 , A2) VALUES(@p, @DocumentId - 1)
				SET @p = @DocumentId + 1
			END
			ELSE
			BEGIN
				SET @p = @p + 1
			END
		END
		FETCH NEXT FROM cs INTO @DocumentId
	END

	CLOSE cs

	DEALLOCATE cs
	RETURN
END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazTafsili8]'
GO
SET QUOTED_IDENTIFIER OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazTafsili8](@AccountYear smallint, @Branch int, @KolId1 int= 0, @KolId2 int= 0, @MoeinId1 int= 0, @MoeinId2 int= 0, @TafsiliId int= 0, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
IF ((@KolId2 > 0) AND (@MoeinId2 > 0))
BEGIN
	SELECT     t1.KolId, t1.MoeinId, t1.TafsiliId, MAX(tblAcc_Kol.KolName) AS KolName, MAX(tblAcc_Moein.MoeinName) AS MoeinName, MAX(vw_TafsiliAll.TafsiliName) AS TafsiliName, SUM(t1.bd1) AS sbd1,
	                          SUM(t1.bs1) AS sbs1, SUM(t1.bd2) AS sbd2, SUM(t1.bs2) AS sbs2, SUM(t1.bd3) AS sbd3, SUM(t1.bs3) AS sbs3
	    FROM         (SELECT     KolId, MoeinId, TafsiliId, SUM(Bedehkar) AS bd1, SUM(Bestankar) AS bs1, 0 AS bd2,
	                                                  0 AS bs2, 0 AS bd3, 0 AS bs3
	                            FROM         tblAcc_DocumentDetail INNER JOIN
	                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
								tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
								tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear < @AccountYear) AND (((tblAcc_DocumentDetail.KolId > @KolId1) AND (tblAcc_DocumentDetail.KolId < @KolId2)) OR
						((tblAcc_DocumentDetail.KolId = @KolId1) AND (tblAcc_DocumentDetail.MoeinId >= @MoeinId1)) OR ((tblAcc_DocumentDetail.KolId = @KolId2) AND (tblAcc_DocumentDetail.MoeinId <= @MoeinId2))) AND ((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId))
	                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId
	                            Union All
	                            SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2, 0 AS bd3,
	                                                  0 AS bs3
	                            FROM         tblAcc_DocumentDetail INNER JOIN
	                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
								tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
								tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < @d1) AND (((tblAcc_DocumentDetail.KolId > @KolId1) AND (tblAcc_DocumentDetail.KolId < @KolId2)) OR
						((tblAcc_DocumentDetail.KolId = @KolId1) AND (tblAcc_DocumentDetail.MoeinId >= @MoeinId1)) OR ((tblAcc_DocumentDetail.KolId = @KolId2) AND (tblAcc_DocumentDetail.MoeinId <= @MoeinId2))) AND ((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId))
	                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId
	                            Union All
	                            SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS bd1, 0 AS bs1, 0 AS bd2, 0 AS bs2, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd3, SUM(tblAcc_DocumentDetail.Bestankar)
	                                                  AS bs3
	                            FROM         tblAcc_DocumentDetail INNER JOIN
	                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
								tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
								tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND (((tblAcc_DocumentDetail.KolId > @KolId1) AND (tblAcc_DocumentDetail.KolId < @KolId2)) OR
						((tblAcc_DocumentDetail.KolId = @KolId1) AND (tblAcc_DocumentDetail.MoeinId >= @MoeinId1)) OR ((tblAcc_DocumentDetail.KolId = @KolId2) AND (tblAcc_DocumentDetail.MoeinId <= @MoeinId2))) AND ((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId))
	                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId) t1 INNER JOIN
	                          tblAcc_Kol ON t1.KolId = tblAcc_Kol.KolId INNER JOIN
	                          tblAcc_Moein ON t1.KolId = tblAcc_Moein.KolId AND t1.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
	                          vw_TafsiliAll ON t1.TafsiliId = vw_TafsiliAll.TafsiliId
	GROUP BY t1.KolId, t1.MoeinId, t1.TafsiliId
	ORDER BY t1.KolId, t1.MoeinId, t1.TafsiliId
END
ELSE
BEGIN
	SELECT     t1.KolId, t1.MoeinId, t1.TafsiliId, MAX(tblAcc_Kol.KolName) AS KolName, MAX(tblAcc_Moein.MoeinName) AS MoeinName, MAX(vw_TafsiliAll.TafsiliName) AS TafsiliName, SUM(t1.bd1) AS sbd1,
	                          SUM(t1.bs1) AS sbs1, SUM(t1.bd2) AS sbd2, SUM(t1.bs2) AS sbs2, SUM(t1.bd3) AS sbd3, SUM(t1.bs3) AS sbs3
	    FROM         (SELECT     KolId, MoeinId, TafsiliId, SUM(Bedehkar) AS bd1, SUM(Bestankar) AS bs1, 0 AS bd2,
	                                                  0 AS bs2, 0 AS bd3, 0 AS bs3
	                            FROM         tblAcc_DocumentDetail INNER JOIN
	                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
								tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
								tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear < @AccountYear) AND ((@KolId2 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2)) AND
						((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId))
	                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId
	                            Union All
	                            SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2, 0 AS bd3,
	                                                  0 AS bs3
	                            FROM         tblAcc_DocumentDetail INNER JOIN
	                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
								tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
								tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < @d1) AND ((@KolId2 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2)) AND
						((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId))
	                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId
	                            Union All
	                            SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS bd1, 0 AS bs1, 0 AS bd2, 0 AS bs2, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd3, SUM(tblAcc_DocumentDetail.Bestankar)
	                                                  AS bs3
	                            FROM         tblAcc_DocumentDetail INNER JOIN
	                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
								tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
								tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
	                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND ((@KolId2 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2)) AND
						((@TafsiliId = 0) OR (tblAcc_DocumentDetail.TafsiliId = @TafsiliId))
	                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId) t1 INNER JOIN
	                          tblAcc_Kol ON t1.KolId = tblAcc_Kol.KolId INNER JOIN
	                          tblAcc_Moein ON t1.KolId = tblAcc_Moein.KolId AND t1.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
	                          vw_TafsiliAll ON t1.TafsiliId = vw_TafsiliAll.TafsiliId
	GROUP BY t1.KolId, t1.MoeinId, t1.TafsiliId
	ORDER BY t1.KolId, t1.MoeinId, t1.TafsiliId
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsilis]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsilis]
				
		AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[TafsiliName],
				[Active]
		
		FROM 
		
		[tblAcc_Tafsili]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[InsequenceInDocumentDate]'
GO
SET ANSI_NULLS OFF
GO

CREATE FUNCTION [dbo].[InsequenceInDocumentDate](@AccountYear smallint, @Branch int, @DocumentId1 int, @DocumentId2 int)
RETURNS @ReturnTable TABLE(
	[DocumentId] [int] NOT NULL ,
	[sdt] [varchar](10) NOT NULL,
	[DocumentDes] [nvarchar] (50) COLLATE Arabic_CI_AS NOT NULL,
	[State] [tinyint] NOT NULL
)
AS
BEGIN
	DECLARE @DocumentId int
	DECLARE @DocumentDate int
	DECLARE @DocumentDes nvarchar(50)
	DECLARE @State tinyint
	DECLARE @fd int
	DECLARE @t TABLE (
		[DocumentId] [int] NOT NULL ,
		[DocumentDate] [int] NOT NULL,
		[DocumentDes] [nvarchar] (50) COLLATE Arabic_CI_AS NOT NULL,
		[State] [tinyint] NOT NULL
	)

	DECLARE cr CURSOR
	KEYSET
	FOR SELECT     DocumentId, DocumentDate, DocumentDes, State--dbo.ConvIntToDateFormat(DocumentDate) AS sdt, DocumentDes, Bedehkar, Bestankar
	FROM         tblAcc_DocumentHeader
	WHERE        (State IN (1, 2)) AND (DocumentId BETWEEN @DocumentId1 AND @DocumentId2)
	ORDER BY DocumentId

	OPEN cr

	FETCH NEXT FROM cr INTO @DocumentId, @DocumentDate, @DocumentDes, @State
	IF (@@fetch_status <> -1) AND (@@fetch_status <> -2)
	BEGIN
		SET @fd = @DocumentDate
	END
	WHILE (@@fetch_status <> -1)
	BEGIN
		IF (@@fetch_status <> -2)
		BEGIN
			IF (@DocumentDate >= @fd)
			BEGIN
				SET @fd = @DocumentDate
			END
			ELSE
			BEGIN
				INSERT INTO @t(DocumentId, DocumentDate, DocumentDes, State)
				VALUES(@DocumentId, @DocumentDate, @DocumentDes, @State)
			END
		END
		FETCH NEXT FROM cr INTO @DocumentId, @DocumentDate, @DocumentDes, @State
	END

	CLOSE cr

	INSERT INTO @RETURNTABLE
	SELECT DocumentId, dbo.ConvIntToDateFormat(DocumentDate) AS sdt, DocumentDes, State
	FROM @t
	ORDER BY DocumentId

	RETURN
END




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_MaxDocumentIdAndDocumentDateByStateEq2]'
GO
SET ANSI_NULLS ON
GO


CREATE PROCEDURE [dbo].[Get_MaxDocumentIdAndDocumentDateByStateEq2](@AccountYear smallint, @Branch int) AS
SELECT     MAX(tblAcc_DocumentHeader.DocumentId) AS maxDocumentId, MAX(tblAcc_DocumentHeader.DocumentDate) AS maxDocumentDate
FROM         tblAcc_DocumentHeader INNER JOIN
                      tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND 
                      tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE     (tblAcc_DocumentHeader.State = 2) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch)


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tBranchs_Count]'
GO





-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tBranchs_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tBranch]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazTafsili6]'
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazTafsili6](@AccountYear smallint, @Branch int, @KolId1 int= 0, @KolId2 int= 0, @MoeinId1 int= 0, @MoeinId2 int= 0, @TafsiliId int= 0, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
SELECT     t1.KolId, t1.MoeinId, t1.TafsiliId, MAX(tblAcc_Kol.KolName) AS KolName, MAX(tblAcc_Moein.MoeinName) AS MoeinName, MAX(vw_TafsiliAll.TafsiliName) AS TafsiliName, SUM(t1.bd1) AS sbd1,
                          SUM(t1.bs1) AS sbs1, SUM(t1.bd2) AS sbd2, SUM(t1.bs2) AS sbs2
    FROM         (SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd1, SUM(tblAcc_DocumentDetail.Bestankar) AS bs1, 0 AS bd2,
                                                  0 AS bs2
                            FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate < @d1)
                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId
                            Union All
                            SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar)
                                                  AS bs2
                            FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                            WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)
                            GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId) t1 INNER JOIN
                          tblAcc_Kol ON t1.KolId = tblAcc_Kol.KolId INNER JOIN
                          tblAcc_Moein ON t1.KolId = tblAcc_Moein.KolId AND t1.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                          vw_TafsiliAll ON t1.TafsiliId = vw_TafsiliAll.TafsiliId
GROUP BY t1.KolId, t1.MoeinId, t1.TafsiliId
ORDER BY t1.KolId, t1.MoeinId, t1.TafsiliId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Kols_Count]'
GO
SET ANSI_NULLS ON
GO


-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Kols_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Kol]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_JabejaeiTarikh]'
GO

CREATE PROCEDURE [dbo].[Get_All_JabejaeiTarikh](@AccountYear smallint, @Branch int, @DocumentId1 int, @DocumentId2 int) AS
SELECT     DocumentId, sdt, DocumentDes, CASE State WHEN 1 THEN N'ويرايش' WHEN 2 THEN N'ثبت موقت' WHEN 3 THEN N'تاييد' END AS StateDesp
FROM         dbo.InsequenceInDocumentDate(@AccountYear, @Branch, @DocumentId1, @DocumentId2) t



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_PaySandoghToAccountCash]'
GO
CREATE PROC [dbo].[Get_All_PaySandoghToAccountCash]
(
@PayType INT,
@FromDate NVARCHAR(8),
@ToDate NVARCHAR(8)
)
AS
BEGIN
SELECT [intSerialNo],
		[DateT],
		[Sanad_Cash],
		[PayTafsili],
		[RecTafsili],
		[Resid],
		[Price]
 FROM [dbo].[tblAcc_PaymentSanad]
WHERE [PaymentTypeId]=12 AND [DateT]>=@FromDate AND [DateT]<=@ToDate
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentHeaders_Count]'
GO
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentHeaders_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_DocumentHeader]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tBranchs_Paged]'
GO





-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tBranchs_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, nvcBranchName nvarchar(50), Type int, Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, nvcBranchName, Type, Active
		
	FROM [tBranch] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, nvcBranchName, Type, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_CommitB]'
GO
CREATE PROCEDURE  [dbo].[Get_All_CommitB](@AccountYear smallint, @Branch int, @DocumentDate int, @NewDocumentId int) AS
SELECT DocumentId, DocumentId2, dbo.ConvIntToDateFormat(DocumentDate) AS dt, DocumentDes FROM [dbo].[Commit](@AccountYear, @Branch, @DocumentDate, @NewDocumentId) t ORDER BY DocumentId2
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Kols_ByFK_GroupID]'
GO

CREATE PROCEDURE [dbo].[Get_tblAcc_Kols_ByFK_GroupID] (
	
	
	@GroupID int
		
) AS

SELECT 


		[KolID],
		[GroupID],
		[KolName],
		[Active]

FROM 

[tblAcc_Kol]

WHERE


	[GroupID] = @GroupID



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentHeaders]'
GO

CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentHeaders]
				
		AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[DocumentDate],
				[DocumentDes],
				[State],
				[DocumentId2],
				[DocumentKind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentHeader]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Kol_ByID]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_Kol_ByID] (
		@KolID int
		) AS
		
		SELECT 		
				[KolID],
				[GroupID],
				[KolName],
				[Active]
		FROM 
		[tblAcc_Kol]
		WHERE
			[KolID] = @KolID


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_BargashtiChequeByChequeNo]'
GO

CREATE PROC [dbo].[Get_BargashtiChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_RecieveSanad] 
WHERE [CheckNo]=@ChequeNo AND [RecieveTypeId]=3 AND [CheckNo]<>0
END 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_PaySandoghToSandoghCash]'
GO
CREATE PROC [dbo].[Get_All_PaySandoghToSandoghCash]
(
@PayType INT,
@FromDate NVARCHAR(8),
@ToDate NVARCHAR(8)
)
AS
BEGIN
SELECT [intSerialNo],
		[DateT],
		[Sanad_Cash],
		[PayTafsili],
		[RecTafsili],
		[Resid],
		[Price]
 FROM [dbo].[tblAcc_PaymentSanad]
WHERE [PaymentTypeId]=11 AND [DateT]>=@FromDate AND [DateT]<=@ToDate
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_CommitA]'
GO

CREATE PROCEDURE  [dbo].[Get_All_CommitA](@AccountYear smallint, @Branch int, @DocumentDate int, @NewDocumentId int) AS
SELECT DocumentId, DocumentId2, dbo.ConvIntToDateFormat(DocumentDate) AS dt, DocumentDes FROM [dbo].[Commit](@AccountYear, @Branch, @DocumentDate, @NewDocumentId) t ORDER BY DocumentId
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_BargashtiMoshtariChequeByChequeNo]'
GO

CREATE PROC [dbo].[Get_BargashtiMoshtariChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_RecieveSanad] 
WHERE [CheckNo]=@ChequeNo AND RecieveTypeId=1 AND [CheckNo]<>0
END 

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Update_tBank]'
GO

ALTER PROCEDURE [dbo].[Update_tBank] (
	@tintBank tinyint ,
	@nvcBankName nvarchar(25) ,
	@intStatus int out)

AS

Begin Tran

UPDATE tblAcc_Bank SET nvcBankName = @nvcBankName WHERE tintBank = @tintBank

if @@Error <> 0 
	Goto ErrHandler

Commit Tran



SET @intStatus = 0
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Userss_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Userss_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		UserId smallint, UserLogin nvarchar(25), UserPassword nvarchar(25), UserName nvarchar(60), UGroupId tinyint, Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT UserId, UserLogin, UserPassword, UserName, UGroupId, Active
		
	FROM [tblAcc_Users] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT UserId, UserLogin, UserPassword, UserName, UGroupId, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_ReceivedCheckByType]'
GO
CREATE PROC [dbo].[Get_ReceivedCheckByType](
@ReceivedType INT,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) 
)
AS
BEGIN
SELECT *,(SELECT [ReceiveTypeName] FROM [dbo].[tblAcc_RecieveType] WHERE [dbo].[tblAcc_RecieveType].[RecieveTypeId]=[dbo].[tblAcc_RecieveSanad].[RecieveTypeId])AS ReceiveTypeName
FROM [dbo].[tblAcc_RecieveSanad]
 
where (
	  (@ReceivedType=1 AND RecieveTypeId=1) OR
	  (@ReceivedType=2 AND RecieveTypeId=1) OR
	  (@ReceivedType=2 AND RecieveTypeId=5) OR
	  (@ReceivedType=3 AND RecieveTypeId=1) OR
      (@ReceivedType=4 AND RecieveTypeId=3) OR
	  (@ReceivedType=5 AND RecieveTypeId=3) OR
	  (@ReceivedType=6 AND RecieveTypeId=1) OR
	  (@ReceivedType=6 AND RecieveTypeId=5) OR
	  (@ReceivedType=7 AND RecieveTypeId=7) OR
	  (@ReceivedType=8 AND RecieveTypeId=8)
		)
AND [DateT]>=@FromDate AND [DateT]<=@ToDate

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_GapInDocumentId]'
GO

CREATE PROCEDURE [dbo].[Get_All_GapInDocumentId](@AccountYear smallint, @Branch int) AS
SELECT * FROM dbo.Gap(@AccountYear, @Branch) t ORDER BY A1

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_PaymentSanad]'
GO


CREATE PROCEDURE [dbo].[Update_tblAcc_PaymentSanad]
(
	@CheckNo NVARCHAR(20),
	@DateS NVARCHAR(10),
	@Price BIGINT ,
	@Descs NVARCHAR(200),
	@PaymentTypeId TINYINT,
	@DateT NVARCHAR(10),
	@RecKol INT,
	@RecMoein INT ,
	@RecTafsili INT ,
	@Taraf NVARCHAR(50) ,
	@PayKol INT ,
	@PayMoein INT ,
	@PayTafsili INT ,
	@PayTafsiliName nvarchar(50),
	@Resid NVARCHAR(50),
	@Result INT OUT 

) 
AS
UPDATE [tblAcc_PaymentSanad]
   SET 
       [DateS] = @DateS
      ,[Price] = @Price
      ,[Descs] = @Descs
      ,[PaymentTypeId] = @PaymentTypeId
      ,[DateT] = @DateT
      ,[RecKol] = @RecKol
      ,[RecMoein] = @RecMoein
      ,[RecTafsili] = @RecTafsili
      ,[Taraf] = @Taraf
      ,[PayKol] = @PayKol
      ,[PayMoein] = @PayMoein
      ,[PayTafsili] = @PayTafsili
      ,[PayTafsiliName] = @PayTafsiliName
	  ,[Resid] =@Resid

 WHERE [CheckNo]=@CheckNo 		

     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result =1
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moeins_ByFK_KolID_Paged]'
GO
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Moeins_ByFK_KolID_Paged] (
			
			
			@KolID int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, MoeinId int, MoeinName nvarchar(50), Kind tinyint, Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, MoeinId, MoeinName, Kind, Active
		
	FROM [tblAcc_Moein] 
	
	WHERE
		
		
			[KolID] = @KolID	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, MoeinId, MoeinName, Kind, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliByKolMoeinAtf2]'
GO
CREATE  PROC [dbo].[Get_TafsiliByKolMoeinAtf2]( @AtfId INT )
AS	
    BEGIN
        SELECT DISTINCT
                dbo.tblAcc_Tafsili.TafsiliName ,
                dbo.tblAcc_Tafsili.TafsiliId
		FROM    dbo.tblAcc_Tafsili
				JOIN dbo.tblAcc_Tafsili_Atf ON dbo.tblAcc_Tafsili.TafsiliId = dbo.tblAcc_Tafsili_Atf.TafsiliId AND tblAcc_Tafsili.TafsiliId <> 0
		WHERE dbo.tblAcc_Tafsili_Atf.AtfId=@AtfId 

    END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_BargashtiMoshtariPayChequeByChequeNo]'
GO
Create PROC [dbo].[Get_BargashtiMoshtariPayChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_PaymentSanad]
WHERE [CheckNo]=@ChequeNo AND [PaymentTypeId]=2 AND [CheckNo]<>0
END 

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_DocumentHeader_State_UserId]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_DocumentHeader_State_UserId] (
		
		
@AccountYear smallint, 		
@Branch int, 		
@DocumentId int, 		
@State tinyint, 		
@UserId int


) AS

UPDATE [tblAcc_DocumentHeader]

SET


		[State] = @State,
		[UserId] = @UserId


WHERE



	[AccountYear] = @AccountYear AND 
	[Branch] = @Branch AND 
	[DocumentId] = @DocumentId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_General_ReceivedSanad]'
GO
CREATE proc [dbo].[Delete_General_ReceivedSanad]
(
@SerialNo int,
@SanadNo int,
@Branch int,
@AccountYear smallint
)
as
begin


DELETE FROM [tblAcc_RecieveSanad]
      WHERE intSerialNo=@SerialNo
      
DELETE FROM [tblAcc_DocumentDetail]
      WHERE DocumentId=@SanadNo AND [Branch]=@Branch AND [AccountYear]=@AccountYear
      
DELETE FROM [tblAcc_DocumentHeader]
      WHERE DocumentId=@SanadNo AND [Branch]=@Branch AND [AccountYear]=@AccountYear
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_BargashtiPayChequeByChequeNo]'
GO

Create PROC [dbo].[Get_BargashtiPayChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_PaymentSanad]
WHERE [CheckNo]=@ChequeNo AND [PaymentTypeId]=2 AND [CheckNo]<>0
END 

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsilis_For_Branch]'
GO

CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsilis_For_Branch]
(
@Branch int
)		
	AS

	SELECT 


			[Branch],
			[TafsiliId],
			[TafsiliName],
			[Active]

	FROM 

	[tblAcc_Tafsili]

	WHERE 

	[Branch] = @Branch

	ORDER BY 

	[TafsiliId]

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[KolNameMoeinName]'
GO


CREATE FUNCTION [dbo].[KolNameMoeinName](@KolId int, @MoeinId int)  
RETURNS nvarchar(105) AS  
BEGIN 
	DECLARE rs CURSOR
	READ_ONLY
	FOR SELECT KolName FROM tblAcc_Kol WHERE KolId = @KolId

	DECLARE @s nvarchar(50)
	DECLARE @r nvarchar(105)

	SET @r = ''
	OPEN rs

	FETCH NEXT FROM rs INTO @s
	WHILE (@@fetch_status <> -1)
	BEGIN
		IF (@@fetch_status <> -2)
		BEGIN
			SET @r = @s
		END
		FETCH NEXT FROM rs INTO @s
	END

	CLOSE rs

	DEALLOCATE rs

	SET @r = @r + ' / '

	DECLARE rs CURSOR
	READ_ONLY
	FOR SELECT MoeinName FROM tblAcc_Moein WHERE KolId = @KolId AND MoeinId = @MoeinId

	OPEN rs

	FETCH NEXT FROM rs INTO @s
	WHILE (@@fetch_status <> -1)
	BEGIN
		IF (@@fetch_status <> -2)
		BEGIN
			SET @r = @r + @s
		END
		FETCH NEXT FROM rs INTO @s
	END

	CLOSE rs

	DEALLOCATE rs

	RETURN @r
END





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_KolNameMoeinName]'
GO

CREATE PROCEDURE [dbo].[Get_KolNameMoeinName](@KolId int, @MoeinId int) AS
SELECT dbo.KolNameMoeinName(@KolId, @MoeinId) AS des



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentDetails]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentDetails]
				
		AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				[kind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentDetail]




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsili_Active]'
GO
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsili_Active]
				
		AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[TafsiliName],
				[Active]
		
		FROM 
		
		[tblAcc_Tafsili]
		
		WHERE 
		
		[Active] =1
		
		ORDER BY 
		
		[TafsiliId]
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_SanadNo_Date]'
GO

CREATE PROCEDURE [dbo].[Get_tblAcc_SanadNo_Date] (
		 	
		@AccountYear smallint,
		@CurrentDate NVARCHAR(8) , 
		@Branch int

		) AS
		
DECLARE @SanadNo INT 
	SELECT @SanadNo =  MAX([DocumentId])
	FROM [tblAcc_DocumentHeader] 
	WHERE [AccountYear] = @AccountYear AND DocumentDate = @CurrentDate AND [Branch] = @Branch
IF @SanadNo IS NULL 
	SELECT @SanadNo =  MAX([DocumentId]) + 1
	FROM [tblAcc_DocumentHeader] 
	WHERE [AccountYear] = @AccountYear AND [Branch] = @Branch
IF @SanadNo IS NULL 
	SET @SanadNo = 1 

SELECT @SanadNo AS SanadNo


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsilis_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsilis_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Tafsili]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
--PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad_Kharj]'
--GO
--CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad_Kharj]
--(
--	@RecieveTypeId TINYINT,
--	@KharjTafsili INT ,
--	@KharjTafsiliName NVARCHAR(50) ,
--	@Kharj_Date NVARCHAR(10) ,
--	@Descs NVARCHAR(255),
--	@intSerialNo INT ,
--	@Resid NVARCHAR(255),
--	@Result INT OUT 

--) 
--AS
		
--UPDATE  dbo.tblAcc_RecieveSanad 
--SET 	
--	RecieveTypeId = @RecieveTypeId ,
--	KharjTafsili = @KharjTafsili ,
--	KharjTafsiliName = @KharjTafsiliName ,
--	Kharj_Date = @Kharj_Date, 
--	Resid=@Resid,
--	Descs=@Descs
--WHERE intserialNo = @intSerialNo

--     IF @@ERROR <>0
--        GoTo EventHandler

--    SET @Result =@intSerialNo
--RETURN @Result

--EventHandler:
--    SET @Result = -1
--	RETURN @Result
--GO
--IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
--GO
--IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
--GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentDetails_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentDetails_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_DocumentDetail]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_AccountYear_Branch_DocumentId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_AccountYear_Branch_DocumentId_Paged] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AccountYear smallint, Branch int, DocumentId int, RowId int, KolId int, MoeinId int, TafsiliId int, RowDes nvarchar(100), Bedehkar int, Bestankar int, kind tinyint, SaveDate int, UserId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		
	FROM [tblAcc_DocumentDetail] 
	
	WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_ReceivedCheck]'
GO
CREATE PROCEDURE [dbo].[Get_All_ReceivedCheck] 
(
@BankNo INT ,
@RecieveTypeId TINYINT ,
@BankTafsili INT ,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) 
)
AS
begin
		SELECT * ,
		(SELECT [ReceiveTypeName] FROM [dbo].[tblAcc_RecieveType] WHERE [dbo].[tblAcc_RecieveType].[RecieveTypeId]=[dbo].[tblAcc_RecieveSanad].[RecieveTypeId])AS ReceiveTypeName
FROM [dbo].[tblAcc_RecieveSanad]		
 
WHERE [tblAcc_RecieveSanad].RecieveTypeId<=6 AND [CheckNo] is NOT null
		and DateT>=@FromDate and DateT<@ToDate	
		ORDER BY [DateT]	
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Do_Reorder]'
GO

CREATE PROCEDURE [dbo].[Do_Reorder](@AccountYear smallint, @Branch int) AS

DECLARE @DocumentId int
DECLARE @dc int

SELECT @dc = CASE WHEN MAX(DocumentId) IS NOT NULL THEN MAX(DocumentId) ELSE 0 END FROM tblAcc_DocumentHeader WHERE @AccountYear = @AccountYear AND @Branch = @Branch AND State = 3

UPDATE    tblAcc_DocumentHeader
SET              DocumentId = - DocumentId
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (State < 3)

DECLARE cr CURSOR
KEYSET
FOR SELECT DocumentId FROM tblAcc_DocumentHeader WHERE AccountYear = @AccountYear AND Branch = @Branch ORDER BY DocumentDate ASC, DocumentId DESC

OPEN cr

FETCH NEXT FROM cr INTO @DocumentId
WHILE (@@fetch_status <> -1)
BEGIN
	IF (@@fetch_status <> -2)
	BEGIN
		SET @dc = @dc + 1
		UPDATE tblAcc_DocumentHeader 
		SET DocumentId = @dc
		WHERE (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId)
	END
	FETCH NEXT FROM cr INTO @DocumentId
END

CLOSE cr
DEALLOCATE cr


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Tafsili]'
GO


CREATE  PROCEDURE [dbo].[Insert_tblAcc_Tafsili] (

		@Branch int, 		
		@TafsiliId int, 		
		@TafsiliName nvarchar(50), 		
		@Active bit,
		@AtfCode int
	) 

	AS
	Begin	

	INSERT INTO [tblAcc_Tafsili]

	(
		[Branch],
		[TafsiliId],
		[TafsiliName],
		[Active]
	)		

	VALUES		
	(
		@Branch,
		@TafsiliId,
		@TafsiliName,
		@Active
	)
INSERT INTO [dbo].[tblAcc_Tafsili_Atf]
           ([Branch]
           ,[TafsiliId]
           ,[AtfId])
     VALUES
           (@Branch
           ,@TafsiliId
           ,@AtfCode)		
		End






GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_ReceivedCheckByCheckNoByType]'
GO
CREATE PROC [dbo].[Get_ReceivedCheckByCheckNoByType](
@CheckNo INT,
@ReceivedType int
)
AS
BEGIN
SELECT *,(SELECT [ReceiveTypeName] FROM [dbo].[tblAcc_RecieveType] WHERE [dbo].[tblAcc_RecieveType].[RecieveTypeId]=[dbo].[tblAcc_RecieveSanad].[RecieveTypeId])AS ReceiveTypeName 
FROM [dbo].[tblAcc_RecieveSanad] 
WHERE [CheckNo]=@CheckNo 
AND [RecieveTypeId]=
					(CASE WHEN @ReceivedType=2 THEN 1
					ELSE CASE WHEN @ReceivedType=3 THEN 1
					ELSE CASE WHEN @ReceivedType=4 THEN 3
					ELSE CASE WHEN @ReceivedType=5 THEN 3
					ELSE CASE WHEN @ReceivedType=6 THEN 1
					END END END END END)

AND [CheckNo]<>0
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazMoein6]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazMoein6](@AccountYear smallint, @Branch int, @KolId1 int, @KolId2 int, @MoeinId1 int, @MoeinId2 int, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
SELECT     tblAcc_Kol.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName, t.sbd1, t.sbs1, t.sbd2, t.sbs2
FROM         tblAcc_Kol INNER JOIN
               (SELECT KolId, MoeinId, SUM(bd1) AS sbd1, SUM(bs1) AS sbs1, SUM(bd2) AS sbd2, SUM(bs2) AS sbs2                     FROM
                          (SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd1, SUM(tblAcc_DocumentDetail.Bestankar) AS bs1, 0 AS bd2, 0 AS bs2
                             FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                             WHERE (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate < @d1)
                             GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId
                             Union All
                             SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2
                             FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                             WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)
                             GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId) t1
 GROUP BY KolId, MoeinId) t ON tblAcc_Kol.KolId = t.KolId INNER JOIN
                      tblAcc_Moein ON t.KolId = tblAcc_Moein.KolId AND t.MoeinId = tblAcc_Moein.MoeinId
ORDER BY tblAcc_Kol.KolId, tblAcc_Moein.MoeinId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_DocumentDetail]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO


CREATE PROCEDURE [dbo].[Insert_tblAcc_DocumentDetail] (
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@RowId int, 		
		@KolId int, 		
		@MoeinId int, 		
		@TafsiliId int, 		
		@RowDes nvarchar(255), 		
		@Bedehkar int, 		
		@Bestankar int, 		
		@kind tinyint, 		
		@UserId int
	) 
	
	AS
		
	INSERT INTO [tblAcc_DocumentDetail]
		
	(
		[AccountYear],
		[Branch],
		[DocumentId],
		[RowId],
		[KolId],
		[MoeinId],
		[TafsiliId],
		[RowDes],
		[Bedehkar],
		[Bestankar],
		[kind],
		[SaveDate],
		[UserId]
	)		
		
	VALUES		
	(
		@AccountYear,
		@Branch,
		@DocumentId,
		@RowId,
		@KolId,
		@MoeinId,
		@TafsiliId,
		@RowDes,
		@Bedehkar,
		@Bestankar,
		@kind,
		dbo.ShamsiInt(GetDate()),
		@UserId
	)

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_DocumentId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_DocumentId_Paged] (
			
			
			@AccountYear smallint,
			@Branch int,
			@DocumentId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, DocumentId int, Row int, KolId int, MoeinId int, TafsiliId int, RowDes nvarchar(100), Bedehkar int, Bestankar int, kind tinyint, SaveDate int, UserId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, DocumentId, Row, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		
	FROM [tblAcc_DocumentDetail] 
	
	WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, DocumentId, Row, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_MaxDocumentIdAdd1]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_MaxDocumentIdAdd1] (
		 	
		@AccountYear smallint, 
		@Branch int

		) AS
		
		SELECT 
		
		
				CASE WHEN MAX([DocumentId]) IS NULL THEN 1 ELSE MAX([DocumentId]) + 1 END AS ms
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_PaymentCheckByTypeDirect]'
GO
CREATE PROC [dbo].[Get_PaymentCheckByTypeDirect](
@PaymentType INT,
@FromDate NVARCHAR(8) ,
@ToDate NVARCHAR(8) 
)
AS
BEGIN
SELECT *,
	(SELECT [PaymentTypeName] FROM [dbo].[tblAcc_PayType] WHERE [dbo].[tblAcc_PayType].[PaymentTypeId]=[dbo].[tblAcc_PaymentSanad].PaymentTypeId)AS PaymentTypeName
FROM [dbo].[tblAcc_PaymentSanad]
where PaymentTypeId=@PaymentType
AND [DateT]>=@FromDate AND [DateT]<=@ToDate

end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_RecieveSanad_BargashtMoshtari]'
GO
CREATE PROCEDURE [dbo].[Update_tblAcc_RecieveSanad_BargashtMoshtari]
(
	@RecieveTypeId TINYINT,
	@BargashtiMoshtari_Date NVARCHAR(10) ,
	@intSerialNo INT ,
	@Descs NVARCHAR(255),
	@Resid NVARCHAR(255),
	@Result INT OUT 

) 
AS
		
UPDATE  dbo.tblAcc_RecieveSanad 
SET 	
	RecieveTypeId = @RecieveTypeId ,
	BargashtiMoshtari_Date=@BargashtiMoshtari_Date,
	Resid=@Resid,
	 Descs=@Descs
WHERE intserialNo = @intSerialNo

     IF @@ERROR <>0
        GoTo EventHandler

    SET @Result = @intSerialNo
RETURN @Result

EventHandler:
    SET @Result = -1
	RETURN @Result
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsilis_ByPK_Branch_TafsiliId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsilis_ByPK_Branch_TafsiliId] (
			
			
			@Branch int,
			@TafsiliId int
				
		) AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[TafsiliName],
				[Active]
		
		FROM 
		
		[tblAcc_Tafsili]
		
		WHERE
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_DocumentHeader]'
GO

CREATE PROCEDURE [dbo].[Update_tblAcc_DocumentHeader] (
				
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@DocumentDate int, 		
		@DocumentDes nvarchar(255), 		
		@State tinyint, 		
		@DocumentId2 int, 		
		@DocumentKind tinyint, 		
		@UserId int

		
		) AS
		
		UPDATE [tblAcc_DocumentHeader]
		
		SET
		
		
				[AccountYear] = @AccountYear,
				[Branch] = @Branch,
				[DocumentId] = @DocumentId,
				[DocumentDate] = @DocumentDate,
				[DocumentDes] = @DocumentDes,
				[State] = @State,
				[DocumentId2] = @DocumentId2,
				[DocumentKind] = @DocumentKind,
				[SaveDate] = dbo.ShamsiInt(GetDate()),
				[UserId] = @UserId

		
		WHERE
		
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazMoein8]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazMoein8](@AccountYear smallint, @Branch int, @KolId1 int, @KolId2 int, @MoeinId1 int, @MoeinId2 int, @d1 int= 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
SELECT     tblAcc_Kol.KolId, tblAcc_Moein.MoeinId, tblAcc_Kol.KolName, tblAcc_Moein.MoeinName, t.sbd1, t.sbs1, t.sbd2, t.sbs2, t.sbd3, t.sbs3
FROM         tblAcc_Kol INNER JOIN
               (SELECT KolId, MoeinId, SUM(bd1) AS sbd1, SUM(bs1) AS sbs1, SUM(bd2) AS sbd2, SUM(bs2) AS sbs2, SUM(bd3) AS sbd3, SUM(bs3) AS sbs3                     FROM
                          (SELECT    tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.Bedehkar AS bd1, tblAcc_DocumentDetail.Bestankar AS bs1, 0 AS bd2, 0 AS bs2, 0 AS bd3, 0 AS bs3                             FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                             WHERE (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear < @AccountYear) AND (((tblAcc_DocumentDetail.KolId > @KolId1) AND (tblAcc_DocumentDetail.KolId < @KolId2)) OR
					((tblAcc_DocumentDetail.KolId = @KolId1) AND (tblAcc_DocumentDetail.MoeinId >= @MoeinId1)) OR ((tblAcc_DocumentDetail.KolId = @KolId2) AND (tblAcc_DocumentDetail.MoeinId <= @MoeinId2)))
                             UNION ALL
                             SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2, 0 AS bd3, 0 AS bs3
                             FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                             WHERE (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < @d1) AND (((tblAcc_DocumentDetail.KolId > @KolId1) AND (tblAcc_DocumentDetail.KolId < @KolId2)) OR
					((tblAcc_DocumentDetail.KolId = @KolId1) AND (tblAcc_DocumentDetail.MoeinId >= @MoeinId1)) OR ((tblAcc_DocumentDetail.KolId = @KolId2) AND (tblAcc_DocumentDetail.MoeinId <= @MoeinId2)))
                             GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId
                             Union All
                             SELECT     tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, 0 AS bd1, 0 AS bs1, 0 AS bd2, 0 AS bs2, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2
                             FROM         tblAcc_DocumentDetail INNER JOIN
                                                  tblAcc_DocumentHeader ON tblAcc_DocumentDetail.AccountYear = tblAcc_DocumentHeader.AccountYear AND
							tblAcc_DocumentDetail.Branch = tblAcc_DocumentHeader.Branch AND
							tblAcc_DocumentDetail.DocumentId = tblAcc_DocumentHeader.DocumentId
                             WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND (((tblAcc_DocumentDetail.KolId > @KolId1) AND (tblAcc_DocumentDetail.KolId < @KolId2)) OR
					((tblAcc_DocumentDetail.KolId = @KolId1) AND (tblAcc_DocumentDetail.MoeinId >= @MoeinId1)) OR ((tblAcc_DocumentDetail.KolId = @KolId2) AND (tblAcc_DocumentDetail.MoeinId <= @MoeinId2)))
                             GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId) t1
GROUP BY KolId, MoeinId) t ON tblAcc_Kol.KolId = t.KolId INNER JOIN
                      tblAcc_Moein ON t.KolId = tblAcc_Moein.KolId AND t.MoeinId = tblAcc_Moein.MoeinId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_ByID]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_ByID] (
		 	
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int

		) AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[DocumentDate],
				[DocumentDes],
				[State],
				[DocumentId2],
				[DocumentKind],
				[SaveDate],
				[UserId]
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_TafsiliId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_Branch_TafsiliId_Paged] (
			
			
			@AccountYear smallint,
			@Branch int,
			@TafsiliId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AccountYear smallint, Branch int, DocumentId int, RowId int, KolId int, MoeinId int, TafsiliId int, RowDes nvarchar(100), Bedehkar int, Bestankar int, kind tinyint, SaveDate int, UserId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		
	FROM [tblAcc_DocumentDetail] 
	
	WHERE
		
		
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
--PRINT N'Creating [dbo].[Insert_tblAcc_RecieveSanad_Variz]'
--GO


--CREATE PROCEDURE [dbo].[Insert_tblAcc_RecieveSanad_Variz]
--(
--	@CheckNo NVARCHAR(20),
--	@DateS NVARCHAR(10),
--	@Price BIGINT ,
--	@Descs NVARCHAR(200),
--	@RecieveTypeId TINYINT,
--	@PayKol INT ,
--	@PayMoein INT ,
--	@PayTafsili INT ,
--	@Taraf NVARCHAR(50) ,
--	@BankKol INT ,
--	@BankMoein INT ,
--	@BankTafsili INT ,
--	@BankTafsiliName NVARCHAR(50) ,
--	@Result INT OUT 

--) 
--AS
		
--INSERT INTO dbo.tblAcc_RecieveSanad (
--	CheckNo,
--	DateS,
--	Price,
--	Descs,
--	RecieveTypeId,
--	PayKol,
--	PayMoein,
--	PayTafsili,
--	Taraf ,
--	BankKol,
--	BankMoein,
--	BankTafsili,
--	BankTafsiliName
--) VALUES ( 
--	@CheckNo ,
--	@DateS ,
--	@Price ,
--	@Descs ,
--	@RecieveTypeId ,
--	@PayKol ,
--	@PayMoein ,
--	@PayTafsili ,
--	@Taraf ,
--	@BankKol ,
--	@BankMoein ,
--	@BankTafsili ,
--	@BankTafsiliName 
--		)
--     IF @@ERROR <>0
--        GoTo EventHandler

--    SET @Result =@@IDENTITY
--RETURN @Result

--EventHandler:
--    SET @Result = -1
--	RETURN @Result



--GO
--IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
--GO
--IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
--GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsilis_ByFK_Branch]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsilis_ByFK_Branch] (
			
			
			@Branch int
				
		) AS
		
		SELECT 
		
		
				[Branch],
				[TafsiliId],
				[TafsiliName],
				[Active]
		
		FROM 
		
		[tblAcc_Tafsili]
		
		WHERE
		
		
			[Branch] = @Branch




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_SanadReceived]'
GO
CREATE PROC [dbo].[Insert_SanadReceived]
(
	@AccountYear smallint, 		
	@Branch int, 		
	@DocumentId int, 		
	@DocumentDate int, 		
	@DocumentDes nvarchar(100), 		
	@State tinyint, 		
	@DocumentId2 int, 		
	@DocumentKind tinyint, 		
	@UserId INT,
	@ds1 NVARCHAR(4000),
	@ds2 NVARCHAR(4000),
	@ds3 NVARCHAR(4000),
	@ItemNo INT,
	@SerialNo int
)
as
begin
DECLARE @number int
SET @number=
		(SELECT COUNT(*) FROM  
		[tblAcc_DocumentHeader]
		WHERE
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId)
IF @number=0 
	begin
		

INSERT INTO [tblAcc_DocumentHeader]
			
		(
			[AccountYear],
			[Branch],
			[DocumentId],
			[DocumentDate],
			[DocumentDes],
			[State],
			[DocumentId2],
			[DocumentKind],
			[SaveDate],
			[UserId] 
		)		
			
		VALUES		
		(
			@AccountYear,
			@Branch,
			@DocumentId,
			@DocumentDate,
			@DocumentDes,
			@State,
			@DocumentId2,
			@DocumentKind,
			dbo.ShamsiInt(GetDate()),
			@UserId
		)

		DELETE  FROM tblAcc_DocumentDetail
		WHERE   AccountYear = @AccountYear
				AND Branch = @Branch
				AND DocumentId = @DocumentId
		IF @ds1<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate
					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds1)
		END
		IF @ds2<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds2)
		END
		IF @ds3<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds3)
			END



-----------------------------------Do_ValidateDocumentDetail
UPDATE    tblAcc_DocumentDetail
SET              Kind = 0
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bedehkar > 0)

UPDATE    tblAcc_DocumentDetail
SET              Kind = 1
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId = @DocumentId) AND (Bestankar > 0)
--------------------------------------------------------------


------------------------------------Update_tblAcc_RecieveSanad_SanadNo
IF @ItemNo = 1 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Daryafti = @DocumentId
	WHERE intSerialNo = @SerialNo
ELSE IF @ItemNo = 2 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Kharj = @DocumentId
			WHERE intSerialNo = @SerialNo
ELSE IF @ItemNo = 3 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Vagozari = @DocumentId
	WHERE intSerialNo = @SerialNo
ELSE IF @ItemNo = 4 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Vosouli = @DocumentId
			WHERE intSerialNo = @SerialNo
ELSE IF @ItemNo = 5 
UPDATE tblAcc_RecieveSanad
	SET Sanad_Bargashti = @DocumentId
	WHERE intSerialNo = @SerialNo
ELSE IF @ItemNo = 6 
UPDATE tblAcc_RecieveSanad
	SET Sanad_BargashtiMoshtari = @DocumentId
			WHERE intSerialNo = @SerialNo

ELSE IF    @ItemNo = 7 OR @ItemNo=8
UPDATE tblAcc_RecieveSanad
	SET Sanad_Cash = @DocumentId
		WHERE intSerialNo = @SerialNo

---------------------------------------------------------------------

	END	
end

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TatbighAsnadGhatei1]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO


CREATE PROCEDURE [dbo].[Get_All_TatbighAsnadGhatei1](@AccountYear smallint, @Branch int, @d int) AS
SELECT     tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdate, t.ct, tblAcc_DocumentHeader.DocumentDes, CASE WHEN t.sBedehkar IS NOT NULL THEN t.sBedehkar ELSE 0 END AS sd
FROM         tblAcc_DocumentHeader INNER JOIN
                          (SELECT AccountYear, Branch, DocumentId, COUNT(DocumentId) AS ct, SUM(Bedehkar) AS sBedehkar FROM tblAcc_DocumentDetail GROUP BY AccountYear, Branch, DocumentId) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
WHERE (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.State = 2) AND (DocumentDate <= @d)
ORDER BY tblAcc_DocumentHeader.DocumentDate, tblAcc_DocumentHeader.DocumentId


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_DocumentDetail]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO


CREATE PROCEDURE [dbo].[Update_tblAcc_DocumentDetail] (
				
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@RowId int, 		
		@KolId int, 		
		@MoeinId int, 		
		@TafsiliId int, 		
		@RowDes nvarchar(255), 		
		@Bedehkar int, 		
		@Bestankar int, 		
		@kind tinyint, 		
		@UserId int

		
		) AS
		
		UPDATE [tblAcc_DocumentDetail]
		SET
				[AccountYear] = @AccountYear,
				[Branch] = @Branch,
				[DocumentId] = @DocumentId,
				[RowId] = @RowId,
				[KolId] = @KolId,
				[MoeinId] = @MoeinId,
				[TafsiliId] = @TafsiliId,
				[RowDes] = @RowDes,
				[Bedehkar] = @Bedehkar,
				[Bestankar] = @Bestankar,
				[kind] = @kind,
				[SaveDate] = dbo.ShamsiInt(GetDate()),
				[UserId] = @UserId

		
		WHERE
			[AccountYear] = @AccountYear AND 
			[Branch] = @Branch AND 
			[DocumentId] = @DocumentId AND 
			[RowId] = @RowId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentHeader_MaxDocumentIdByStateGreaterThan1]'
GO

CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentHeader_MaxDocumentIdByStateGreaterThan1](
		 	
		@AccountYear smallint, 
		@Branch int

		) AS
		
SELECT     DocumentId, DocumentDate
FROM         tblAcc_DocumentHeader
WHERE     (AccountYear = @AccountYear) AND (Branch = @Branch) AND (DocumentId =
                          (SELECT     MAX(DocumentId) AS mv
                             FROM         tblAcc_DocumentHeader
                             WHERE     state > 1))

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TatbighAsnadGhatei2]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO


CREATE PROCEDURE [dbo].[Get_All_TatbighAsnadGhatei2](@AccountYear smallint, @Branch int, @d int) AS
SELECT     tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdate, t.ct, tblAcc_DocumentHeader.DocumentDes, CASE WHEN t.sBedehkar IS NOT NULL THEN t.sBedehkar ELSE 0 END AS sd
FROM         tblAcc_DocumentHeader INNER JOIN
                          (SELECT AccountYear, Branch, DocumentId, COUNT(DocumentId) AS ct, SUM(Bedehkar) AS sBedehkar FROM tblAcc_DocumentDetail GROUP BY AccountYear, Branch, DocumentId) t ON tblAcc_DocumentHeader.AccountYear = t.AccountYear AND tblAcc_DocumentHeader.Branch = t.Branch AND tblAcc_DocumentHeader.DocumentId = t.DocumentId
WHERE (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.State = 2) AND (DocumentDate <= @d)
ORDER BY tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Tafsilis_Paged]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Tafsilis_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, TafsiliId int, TafsiliName nvarchar(50), Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, TafsiliId, TafsiliName, Active
		
	FROM [tblAcc_Tafsili] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, TafsiliId, TafsiliName, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetails_ByFK_KolId_MoeinId_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetails_ByFK_KolId_MoeinId_Paged] (
			
			
			@KolId int,
			@MoeinId int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AccountYear smallint, Branch int, DocumentId int, RowId int, KolId int, MoeinId int, TafsiliId int, RowDes nvarchar(100), Bedehkar int, Bestankar int, kind tinyint, SaveDate int, UserId int
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		
	FROM [tblAcc_DocumentDetail] 
	
	WHERE
		
		
			[KolId] = @KolId AND 
			[MoeinId] = @MoeinId	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AccountYear, Branch, DocumentId, RowId, KolId, MoeinId, TafsiliId, RowDes, Bedehkar, Bestankar, kind, SaveDate, UserId
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_DocumentHeader]'
GO
CREATE PROCEDURE [dbo].[Insert_tblAcc_DocumentHeader] (
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@DocumentDate int, 		
		@DocumentDes nvarchar(255), 		
		@State tinyint, 		
		@DocumentId2 int, 		
		@DocumentKind tinyint, 		
		@UserId int
	) 
	
	AS
		
	INSERT INTO [tblAcc_DocumentHeader]
		
	(
		[AccountYear],
		[Branch],
		[DocumentId],
		[DocumentDate],
		[DocumentDes],
		[State],
		[DocumentId2],
		[DocumentKind],
		[SaveDate],
		[UserId]
	)		
		
	VALUES		
	(
		@AccountYear,
		@Branch,
		@DocumentId,
		@DocumentDate,
		@DocumentDes,
		@State,
		@DocumentId2,
		@DocumentKind,
		dbo.ShamsiInt(GetDate()),
		@UserId
	)




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_General_PaymentSanad]'
GO
CREATE proc [dbo].[Delete_General_PaymentSanad]
(
@SerialNo int,
@SanadNo INT,
@Branch INT,
@AccountYear SMALLINT

)
as
begin


DELETE FROM [tblAcc_PaymentSanad]
      WHERE intSerialNo=@SerialNo
      
DELETE FROM [tblAcc_DocumentDetail]
      WHERE DocumentId=@SanadNo AND [Branch]=@Branch AND [AccountYear]=@AccountYear
      
DELETE FROM [tblAcc_DocumentHeader]
      WHERE DocumentId=@SanadNo AND [Branch]=@Branch AND [AccountYear]=@AccountYear
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Group]'
GO



CREATE PROCEDURE [dbo].[Update_tblAcc_Group] (
												@GroupID int, 		
												@GroupName nvarchar(50), 		
												@Active bit
												)
 AS
UPDATE [tblAcc_Group]
SET
		[GroupID] = @GroupID,
		[GroupName] = @GroupName,
		[Active] = @Active


WHERE
	[GroupID] = @GroupID



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TafsilisBetweenTafsiliId]'
GO



CREATE PROCEDURE [dbo].[Get_All_TafsilisBetweenTafsiliId](@Branch int, @TafsiliId1 int, @TafsiliId2 int, @SortBy tinyint) AS
IF (@SortBy = 0)
BEGIN
	IF (@TafsiliId1=0 AND @TafsiliId2=0)
	BEGIN
		SELECT     *
		FROM         tblAcc_Tafsili
		WHERE     (Branch = @Branch)
		ORDER BY TafsiliId
	END
	ELSE
	BEGIN
		SELECT     *
		FROM         tblAcc_Tafsili
		WHERE     (Branch = @Branch) AND (TafsiliId BETWEEN @TafsiliId1 AND @TafsiliId2)
		ORDER BY TafsiliId
	END
END
ELSE
BEGIN
	IF (@TafsiliId1=0 AND @TafsiliId2=0)
	BEGIN
		SELECT     *
		FROM         tblAcc_Tafsili
		WHERE     (Branch = @Branch)
		ORDER BY TafsiliName
	END
	ELSE
	BEGIN
		SELECT     *
		FROM         tblAcc_Tafsili
		WHERE     (Branch = @Branch) AND (TafsiliId BETWEEN @TafsiliId1 AND @TafsiliId2)
		ORDER BY TafsiliName
	END
END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_UGroupss_Paged]'
GO



-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_UGroupss_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		UGroupId tinyint, UGroupName nvarchar(40)
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT UGroupId, UGroupName
		
	FROM [tblAcc_UGroups] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT UGroupId, UGroupName
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_KolIdCountsInChilds]'
GO



CREATE PROCEDURE [dbo].[Get_KolIdCountsInChilds](@KolId int) AS
SELECT     SUM(c) AS ct
FROM         (SELECT     COUNT(KolId) AS c
                        FROM         tblAcc_Moein
                        WHERE     (KolId = @KolId)
                        UNION ALL
                        SELECT     COUNT(KolId) AS c
                        FROM         tblAcc_DocumentDetail
                        WHERE     (KolId = @KolId)) t



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DocumentDetails]'
GO
CREATE  PROCEDURE [dbo].[Get_All_DocumentDetails]
(@AccountYear smallint, @Branch int)
				
		AS
SELECT 


				[DocumentId],
				[RowId],
				[KolId],
				[MoeinId],
				[TafsiliId],
				[RowDes],
				[Bedehkar],
				[Bestankar],
				(SELECT [KolName] FROM [dbo].[tblAcc_Kol] WHERE [KolId]=[dbo].[tblAcc_DocumentDetail].[KolId])AS KolName,
				(SELECT [MoeinName] FROM [dbo].[tblAcc_Moein] WHERE [MoeinId]=[dbo].[tblAcc_DocumentDetail].[MoeinId])AS MoeinName,
				(SELECT [TafsiliName] FROM [dbo].[tblAcc_Tafsili] WHERE [TafsiliId]=[dbo].[tblAcc_DocumentDetail].[TafsiliId])AS TafsiliName,
				
				(SELECT [DocumentDate] FROM [dbo].[tblAcc_DocumentHeader] WHERE [DocumentId]=[dbo].[tblAcc_DocumentDetail].[DocumentId] and Branch=tblAcc_DocumentDetail.Branch and AccountYear=tblAcc_DocumentDetail.AccountYear)AS DocumentDate,
				(SELECT [DocumentDes] FROM [dbo].[tblAcc_DocumentHeader] WHERE [DocumentId]=[dbo].[tblAcc_DocumentDetail].[DocumentId] and Branch=tblAcc_DocumentDetail.Branch and AccountYear=tblAcc_DocumentDetail.AccountYear)AS DocumentDes
		FROM 
		
		[tblAcc_DocumentDetail]
		
		WHERE 
		
			[AccountYear] = @AccountYear AND
			[Branch] = @Branch
ORDER BY 
		
			[DocumentId],[RowId] --[kind], [KolId], [MoeinId], [TafsiliId], [Bedehkar], [Bestankar]
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Group]'
GO


CREATE PROCEDURE [dbo].[Insert_tblAcc_Group] (
				
		@GroupID int, 		
		@GroupName nvarchar(50), 		
		@Active bit
	) 
	
	AS
		
	INSERT INTO [tblAcc_Group]
		
	(
		[GroupID],
		[GroupName],
		[Active]
	)		
		
	VALUES		
	(
		@GroupID,
		@GroupName,
		@Active
	)
	


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsili_ByID_Count]'
GO




CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsili_ByID_Count] (
		 	
				
		@Branch int, 
		@TafsiliID int

		) AS
		
		SELECT 
		
		
				COUNT([TafsiliID]) AS ct
		
		FROM 
		
		[tblAcc_Tafsili]
		
		WHERE
		
		
			[Branch] = @Branch AND [TafsiliID] = @TafsiliID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Groups]'
GO


CREATE PROCEDURE [dbo].[Get_All_tblAcc_Groups]
AS
SELECT 
		[GroupID],
		[GroupName],
		[Active]
FROM 
[tblAcc_Group]



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Insert_tBank]'
GO
ALTER  PROCEDURE [dbo].[Insert_tBank] (
	@nvcBankName nvarchar(25) , 
	@intStatus int out)
AS

declare @tintBank int

Begin Tran

SELECT @tintBank = IsNull(Max(tintBank) + 1, 1) FROM tblAcc_Bank

Insert Into dbo.tblAcc_Bank(tintBank, nvcBankName) VALUES(@tintBank, @nvcBankName)
if @@Error <> 0 
	Goto ErrHandler

Commit Tran


SET @intStatus=@tintBank
Return

ErrHandler:
RollBack Tran
Set @intStatus = -1
Return





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_AtfCountsInDocuments]'
GO

CREATE proc [dbo].[Get_AtfCountsInDocuments](@KolId INT,
									 @MoeinId int,
									 @Atf int)
as
BEGIN
	SELECT COUNT(* ) AS [ct]
	FROM tblAcc_DocumentDetail 
	WHERE KolId =@KolId-- 63 
		AND MoeinId =@MoeinId-- 1  
		AND TafsiliId IN(SELECT TafsiliId 
						 FROM dbo.tblAcc_Tafsili_Atf 
						 WHERE AtfId=@Atf)--2)
END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_SanadHeaderSearch]'
GO


CREATE PROC [dbo].[Get_SanadHeaderSearch]
(
@AccountYear SMALLINT,
@Branch INT,
@DocumentId INT,
@DocumentDate INT
)

AS
BEGIN
SELECT 
[DocumentId],
[DocumentDate],
[DocumentDes],
(SELECT SUM([Bedehkar]) FROM [dbo].[tblAcc_DocumentDetail] 
WHERE [AccountYear]=@AccountYear AND Branch=@Branch AND [DocumentId]=[tblAcc_DocumentHeader].[DocumentId])AS SumBedehkar,

(SELECT SUM([Bestankar]) FROM [dbo].[tblAcc_DocumentDetail] 
WHERE [AccountYear]=@AccountYear AND Branch=@Branch AND [DocumentId]=[tblAcc_DocumentHeader].[DocumentId])AS SumBestankar


FROM [dbo].[tblAcc_DocumentHeader]
WHERE [dbo].[tblAcc_DocumentHeader].[AccountYear]=@AccountYear 
AND [dbo].[tblAcc_DocumentHeader].[Branch]=@Branch
AND [DocumentDate]=(CASE WHEN @DocumentDate=0 THEN [DocumentDate] ELSE @DocumentDate END)
AND [DocumentId]=(CASE WHEN @DocumentId=0 THEN [DocumentId] ELSE @DocumentId END)

END	

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Get_All_tBanks]'
GO
ALTER PROCEDURE [dbo].[Get_All_tBanks] AS

	select * from tblAcc_Bank ORDER BY tintBank
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Groups_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Groups_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Group]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsilis_ByFK_Branch_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsilis_ByFK_Branch_Count] (
			
			
			@Branch int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Tafsili]
		
		WHERE
		
		
			[Branch] = @Branch




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_AllSanadSearch]'
GO
CREATE PROC [dbo].[Get_AllSanadSearch]
(
@AccountYear SMALLINT,
@Branch INT
)
AS

BEGIN
SELECT 
[DocumentId],
[DocumentDate],
[DocumentDes],
(SELECT SUM([Bedehkar]) FROM [dbo].[tblAcc_DocumentDetail] 
WHERE [AccountYear]=@AccountYear AND Branch=@Branch AND [DocumentId]=[tblAcc_DocumentHeader].[DocumentId])AS SumBedehkar,

(SELECT SUM([Bestankar]) FROM [dbo].[tblAcc_DocumentDetail] 
WHERE [AccountYear]=@AccountYear AND Branch=@Branch AND [DocumentId]=[tblAcc_DocumentHeader].[DocumentId])AS SumBestankar


FROM [dbo].[tblAcc_DocumentHeader]
WHERE [dbo].[tblAcc_DocumentHeader].[AccountYear]=@AccountYear AND [dbo].[tblAcc_DocumentHeader].[Branch]=@Branch


END	

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_KolIdMoeinIdCountsInChilds]'
GO

CREATE PROCEDURE [dbo].[Get_KolIdMoeinIdCountsInChilds](@KolId int, @MoeinId int) AS
SELECT     SUM(c) AS ct
FROM         (SELECT     COUNT(KolId) AS c
                        FROM         tblAcc_Moein_Atf
                        WHERE     (KolId = @KolId) AND (MoeinId = @MoeinId)
                        UNION ALL
                        SELECT     COUNT(KolId) AS c
                        FROM         tblAcc_DocumentDetail
                        WHERE     (KolId = @KolId) AND (MoeinId = @MoeinId)) t


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Search_Groups]'
GO
CREATE PROC [dbo].[Search_Groups]
(
@Name  NVARCHAR(50),
@Status INT
)
AS

BEGIN

SELECT 
		[GroupID],
		[GroupName],
		[Active]
FROM 
[tblAcc_Group]
WHERE [GroupName] LIKE '%'+@Name+'%'
AND [Active]=(CASE WHEN @Status =0 THEN [Active] ELSE 
			  CASE WHEN @Status=1 THEN 1 ELSE
			  CASE WHEN @Status=2 THEN 0 END END END )
end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Altering [dbo].[Delete_tBank_By_tintBank]'
GO

ALTER PROCEDURE [dbo].[Delete_tBank_By_tintBank](
	@tintBank tinyint)
AS
	DELETE FROM dbo.tblAcc_Bank WHERE tintBank = @tintBank



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Group_ByID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Group_ByID] (
		 	
				
		@GroupID int

		) AS
		
		SELECT 
		
		
				[GroupID],
				[GroupName],
				[Active]
		
		FROM 
		
		[tblAcc_Group]
		
		WHERE
		
		
			[GroupID] = @GroupID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_SanadDetail]'
GO

CREATE PROC [dbo].[Get_SanadDetail](
@AccountYear smallint,
@Branch INt,
@DocumentId INT
)
as

begin

SELECT RowId,
KolId,
MoeinId,
TafsiliId,
RowDes,
Bedehkar,
Bestankar

FROM tblAcc_DocumentDetail 
WHERE AccountYear=@AccountYear
AND 
Branch=@Branch
AND 
DocumentId=@DocumentId

 
END
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Kols_ByFK_GroupID_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Kols_ByFK_GroupID_Paged] (
			
			
			@GroupID int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		KolID int, GroupID int, KolName nvarchar(50), Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT KolID, GroupID, KolName, Active
		
	FROM [tblAcc_Kol] 
	
	WHERE
		
		
			[GroupID] = @GroupID	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT KolID, GroupID, KolName, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_Moein]'
GO



CREATE PROCEDURE [dbo].[Insert_tblAcc_Moein]
    (
      @KolID INT ,
      @MoeinId INT ,
      @MoeinName NVARCHAR(50) ,
      @Kind TINYINT =null,
      @Active BIT=null
	
    )
AS 
    INSERT  INTO [tblAcc_Moein]
            ( [KolID] ,
              [MoeinId] ,
              [MoeinName] ,
              [Kind] ,
              [Active]
			)
    VALUES  ( @KolID ,
              @MoeinId ,
              @MoeinName ,
              ISNULL(@Kind ,0),
              ISNULL( @Active,1)
	      )
		
	

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Group_ByID_Count]'
GO




CREATE PROCEDURE [dbo].[Get_tblAcc_Group_ByID_Count] (
		 	
				
		@GroupID int

		) AS
		
		SELECT 
		
		
				COUNT([GroupID]) AS ct
		
		FROM 
		
		[tblAcc_Group]
		
		WHERE
		
		
			[GroupID] = @GroupID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_TafsiliDetails]'
GO
CREATE  PROCEDURE [dbo].[Insert_tblAcc_TafsiliDetails] (
				
		@Branch int, 		
		@TafsiliId int, 		
		@TafsiliName nvarchar(50), 		
		@Active BIT,
		@AtfId INT
--		,@KolId INT,
--		@MoeinId INT
	) 
	
	AS
BEGIN	
	INSERT INTO [tblAcc_Tafsili]
		
	(
		[Branch],
		[TafsiliId],
		[TafsiliName],
		[Active]
	)		
		
	VALUES		
	(
		@Branch,
		@TafsiliId,
		@TafsiliName,
		@Active
	)
		
	INSERT dbo.tblAcc_Tafsili_Atf
		(
			Branch,
			TafsiliId, 
			AtfId
		 )
	VALUES
		(
		 @Branch,
		 @TafsiliId,
		 @AtfId
		 )

END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Insert_tblAcc_DocumentWithDetail]'
GO


CREATE  PROCEDURE [dbo].[Insert_tblAcc_DocumentWithDetail](
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@DocumentDate int, 		
		@DocumentDes nvarchar(100), 		
		@State tinyint, 		
		@DocumentId2 int, 		
		@DocumentKind tinyint, 		
		@UserId INT,
		@ds1 NVARCHAR(4000),
		@ds2 NVARCHAR(4000),
		@ds3 NVARCHAR(4000),
		@SaleNo INT = NULL  ,
		@Status INT = 2 ,
		@SaleNoAcc INT = NULL ,
		@result INT out
	) 
	
	AS
	
	BEGIN	TRAN
		INSERT INTO [tblAcc_DocumentHeader]
			
		(
			[AccountYear],
			[Branch],
			[DocumentId],
			[DocumentDate],
			[DocumentDes],
			[State],
			[DocumentId2],
			[DocumentKind],
			[SaveDate],
			[UserId] 
		)		
			
		VALUES		
		(
			@AccountYear,
			@Branch,
			@DocumentId,
			@DocumentDate,
			@DocumentDes,
			@State,
			@DocumentId2,
			@DocumentKind,
			dbo.ShamsiInt(GetDate()),
			@UserId
		)

		IF @@ERROR>0
		BEGIN
			ROLLBACK TRAN
			SET @result=0
			RETURN
		END 

		DELETE  FROM tblAcc_DocumentDetail
		WHERE   AccountYear = @AccountYear
				AND Branch = @Branch
				AND DocumentId = @DocumentId
		IF @ds1<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate
					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds1)
			IF @@ERROR>0
			BEGIN
				ROLLBACK TRAN
				SET @result=0
				RETURN
			END 
		END
		IF @ds2<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds2)
			IF @@ERROR>0
			BEGIN
				ROLLBACK TRAN
				SET @result=0
				RETURN
			END 
		END
		IF @ds3<>N''
		BEGIN
			INSERT  INTO tblAcc_DocumentDetail
					( AccountYear ,
					  Branch ,
					  DocumentId ,
					  RowId ,
					  KolId ,
					  MoeinId ,
					  TafsiliId ,
					  RowDes ,
					  Bedehkar ,
					  Bestankar ,
					  kind ,
					  SaveDate ,
					  UserId ,
					  CheckNo ,
					  CheckDate

					)
					SELECT  AccountYear ,
						Branch ,
						DocumentId ,
						RowId ,
						KolId ,
						MoeinId ,
						TafsiliId ,
						RowDes ,
						Bedehkar ,
						Bestankar ,
						kind ,
						SaveDate ,
						UserId ,
					    CheckNo ,
					    CheckDate

				FROM    dbo.Split_Acc(@ds3)
			IF @@ERROR>0
			BEGIN
				ROLLBACK TRAN
				SET @result=0
				RETURN
			END 
		END

	IF @SaleNo > 0 
		BEGIN 
		UPDATE dbo.tFacM SET Refrence_Acc = @DocumentId WHERE Branch = @Branch AND [No] = @SaleNo AND AccountYear = @Accountyear AND Status = @Status
		UPDATE dbo.tFacM SET transferAccounting = 1 WHERE Branch = @Branch AND [No] = @SaleNo AND AccountYear = @Accountyear AND Status = @Status
		UPDATE dbo.tblAcc_DocumentHeader SET Refrence_Sale = @SaleNo WHERE 	AccountYear = @AccountYear AND Branch = @Branch AND DocumentId = @DocumentId
		END 

	IF @SaleNoAcc > 0 
		UPDATE dbo.tblAcc_DocumentHeader SET Refrence_Khazane = @SaleNoAcc WHERE AccountYear = @AccountYear AND Branch = @Branch AND DocumentId = @DocumentId

	COMMIT TRAN
	SET @result= 1
	RETURN
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_PreviousSanadHeader]'
GO
CREATE		 PROC [dbo].[Get_PreviousSanadHeader](
@AccountYear smallint,
@Branch INt,
@DocumentId INT
)
as

begin

SELECT TOP 1
DocumentDate,
DocumentId,
DocumentDes,
State,
DocumentId2,
DocumentKind 
FROM tblAcc_DocumentHeader 
WHERE DocumentId<@DocumentId AND [AccountYear]=@AccountYear AND [Branch]=@Branch
ORDER BY [DocumentId] DESC

 
END
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_DocumentDetail]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Delete_tblAcc_DocumentDetail] (
		
				
		@AccountYear smallint, 		
		@Branch int, 		
		@DocumentId int, 		
		@RowId int=0
		
		) AS
		
		IF (@RowId = 0)
			BEGIN
				DELETE [tblAcc_DocumentDetail]
				
				WHERE
				
				
					[AccountYear] = @AccountYear AND 
					[Branch] = @Branch AND 
					[DocumentId] = @DocumentId 
			END
		ELSE
			BEGIN
				DELETE [tblAcc_DocumentDetail]
				
				WHERE
				
				
					[AccountYear] = @AccountYear AND 
					[Branch] = @Branch AND 
					[DocumentId] = @DocumentId AND 
					[RowId] = @RowId
			END



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moeins_ByPK_KolID_MoeinId]'
GO

CREATE PROCEDURE [dbo].[Get_tblAcc_Moeins_ByPK_KolID_MoeinId] (
																@KolID int,
																@MoeinId int
	
) AS

SELECT 
	[KolID],
	[MoeinId],
	[MoeinName],
	[Kind],
	[Active]
FROM 
[tblAcc_Moein]
WHERE
[KolID] = @KolID AND 
[MoeinId] = @MoeinId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Search_Kols]'
GO
CREATE PROC [dbo].[Search_Kols]
(
@KolName NVARCHAR(50),
@GroupName NVARCHAR(50),
@Status INT
)
AS
BEGIN
		
SELECT 

		
		[KolID],
		[GroupID],
		(SELECT [GroupName] FROM [dbo].[tblAcc_Group] WHERE [GroupId]=tblAcc_Kol.[GroupId])AS GroupName,
		[KolName],
		[Active]

FROM 

[tblAcc_Kol]
WHERE 
(SELECT [GroupName] FROM [dbo].[tblAcc_Group] WHERE [GroupId]=tblAcc_Kol.[GroupId]) LIKE '%'+@GroupName+'%'
AND KolName LIKE '%'+@KolName+'%' 	
and[Active]=(CASE WHEN @Status =0 THEN [Active] ELSE 
			  CASE WHEN @Status=1 THEN 1 ELSE
			  CASE WHEN @Status=2 THEN 0 END END END )
ORDER BY KolId
END	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Tafsili]'
GO
CREATE PROCEDURE [dbo].[Update_tblAcc_Tafsili] (
				
				
		@Branch int, 		
		@TafsiliId int, 		
		@TafsiliName nvarchar(50), 		
		@Active bit

		
		) AS
		
		UPDATE [tblAcc_Tafsili]
		
		SET
		
		
				[Branch] = @Branch,
				[TafsiliId] = @TafsiliId,
				[TafsiliName] = @TafsiliName,
				[Active] = @Active

		
		WHERE
		
		
		
			[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Moein]'
GO


CREATE PROC	 [dbo].[Update_tblAcc_Moein] (
										@KolID int, 		
										@MoeinId int, 		
										@MoeinName nvarchar(50), 		
										@Kind tinyint=NULL, 		
										@Active bit=null
										)
AS
  UPDATE    [tblAcc_Moein]
  SET       --[KolID] = @KolID ,
--            [MoeinId] = @MoeinId ,
            [MoeinName] = @MoeinName ,
            [Kind] =ISNULL(@Kind ,Kind),
            [Active] =ISNULL( @Active,Active)
  WHERE     [KolID] = @KolID
            AND [MoeinId] = @MoeinId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Check_GroupExist]'
GO

CREATE PROC [dbo].[Check_GroupExist](@GroupId INT,@Result INT OUTPUT	)
AS
BEGIN
	SET @Result=0
	IF EXISTS(SELECT * FROM dbo.tblAcc_Group WHERE GroupId=@GroupId)
	 SET @Result=1
END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moeins_ByFK_KolID]'
GO


CREATE PROCEDURE [dbo].[Get_tblAcc_Moeins_ByFK_KolID] 
				(
					@KolID int
				)
AS

SELECT 	[KolID],
		[MoeinId],
		[MoeinName],
		[Kind],
		[Active]
FROM 	[tblAcc_Moein]

WHERE	[KolID] = @KolID



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_RooznamehRiz]'
GO


CREATE  PROCEDURE [dbo].[Get_All_RooznamehRiz](@AccountYear smallint, @Branch int, @d1 int = 0, @d2 int = 0) AS
SELECT     TOP 100 PERCENT tblAcc_DocumentHeader.DocumentId, dbo.ConvIntToDateFormat(tblAcc_DocumentHeader.DocumentDate) AS sdate, tblAcc_DocumentDetail.RowDes, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId,
             ISNULL(tblAcc_Tafsili.TafsiliName , '') AS OnvanHesab ,tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar 
             --, CASE tblAcc_DocumentDetail.TafsiliId WHEN 0 THEN tblAcc_Moein.MoeinName ELSE ISNULL(tblAcc_Tafsili.TafsiliName , '') END AS OnvanHesab
FROM         tblAcc_DocumentHeader INNER JOIN
                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
					tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
					tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId 
					INNER JOIN tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId 
					LEFT OUTER JOIN tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
WHERE (State > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)
ORDER BY tblAcc_DocumentHeader.DocumentId





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_SanadHeaderByDocumentId]'
GO
CREATE PROC [dbo].[Get_SanadHeaderByDocumentId](
@AccountYear smallint,
@Branch INt,
@DocumentId INT
)
as

begin

SELECT
DocumentDate,
DocumentId,
DocumentDes,
State,
DocumentId2,
DocumentKind 
FROM tblAcc_DocumentHeader 
WHERE DocumentId=@DocumentId AND [AccountYear]=@AccountYear AND [Branch]=@Branch	

 
END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Tafsili]'
GO


CREATE  PROCEDURE [dbo].[Delete_tblAcc_Tafsili] (
	@Branch int, 		
	@TafsiliId INT
	)
 AS
BEGIN

	DELETE FROM [dbo].[tblAcc_Tafsili_Atf]
	WHERE	[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId
	DELETE FROM  [tblAcc_Tafsili]
	WHERE	[Branch] = @Branch AND 
			[TafsiliId] = @TafsiliId
END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moeins]'
GO
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moeins]
				
		AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[MoeinName],
				[Kind],
				[Active]
		
		FROM 
		
		[tblAcc_Moein]
		

	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Tafsilis_ByFK_Branch_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_tblAcc_Tafsilis_ByFK_Branch_Paged] (
			
			
			@Branch int,
			@__PageNumber int,
			@__PageSize int	
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		Branch int, TafsiliId int, TafsiliName nvarchar(50), Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT Branch, TafsiliId, TafsiliName, Active
		
	FROM [tblAcc_Tafsili] 
	
	WHERE
		
		
			[Branch] = @Branch	


	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT Branch, TafsiliId, TafsiliName, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazKol]'
GO
SET QUOTED_IDENTIFIER OFF
GO
SET ANSI_NULLS OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazKol](@AccountYear smallint, @Branch int, @KolId1 int = 0, @KolId2 int = 0, @d1 int = 0, @d2 int = 0, @DocumentId1 int = 0, @DocumentId2 int = 0) AS
SELECT     tblAcc_Kol.KolId, MAX(tblAcc_Kol.KolName) AS KolName, CASE WHEN SUM(sbd) IS NOT NULL THEN SUM(sbd) ELSE 0 END AS sd, CASE WHEN SUM(sbs)  IS NOT NULL THEN SUM(sbs) ELSE 0 END AS ss, CASE WHEN SUM(sbd) >= SUM(sbs) THEN SUM(sbd) - SUM(sbs) ELSE 0 END AS rd, CASE WHEN SUM(sbs) >= SUM(sbd) THEN SUM(sbs) - SUM(sbd) ELSE 0 END AS rs
FROM tblAcc_Kol INNER JOIN
(SELECT     tblAcc_DocumentDetail.KolId, CASE WHEN SUM(tblAcc_DocumentDetail.Bedehkar) IS NOT NULL THEN SUM(tblAcc_DocumentDetail.Bedehkar) ELSE 0 END AS sbd, CASE WHEN SUM(tblAcc_DocumentDetail.Bestankar) IS NOT NULL THEN SUM(tblAcc_DocumentDetail.Bestankar) ELSE 0 END AS sbs
FROM         tblAcc_DocumentHeader INNER JOIN
                                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
							tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
							tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND ((@d2 = 0) OR (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) AND ((@DocumentId2 = 0) OR (tblAcc_DocumentHeader.DocumentId BETWEEN @DocumentId1 AND @DocumentId2)) AND ((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2))
GROUP BY tblAcc_DocumentDetail.KolId) t ON tblAcc_Kol.KolId = t.KolId
GROUP BY tblAcc_Kol.KolId
ORDER BY tblAcc_Kol.KolId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_DocumentHeaders_By_AccountYear_Branch]'
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_NULLS ON
GO


CREATE PROCEDURE [dbo].[Get_All_tblAcc_DocumentHeaders_By_AccountYear_Branch](@AccountYear smallint, @Branch int)
				
		AS
		
		SELECT 
		
		
				[AccountYear],
				[Branch],
				[DocumentId],
				[DocumentDate],
				[DocumentDes],
				[State],
				[SaveDate],
				[UserId] ,
				Refrence_Sale  ,
				Refrence_Khazane
		
		FROM 
		
		[tblAcc_DocumentHeader]
		
		WHERE 
		
		[AccountYear] = @AccountYear AND [Branch] = @Branch
		
		ORDER BY 
		
		[DocumentId]




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moeins_Count]'
GO
		
		CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moeins_Count]
				
		AS
		
		SELECT
		
			COUNT(*)
		
		FROM
		
			[tblAcc_Moein]
		
		
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliByKolMoeinAtf]'
GO



CREATE  	 PROC [dbo].[Get_TafsiliByKolMoeinAtf]( @AtfId INT )
AS	
    BEGIN
        SELECT DISTINCT
                dbo.tblAcc_Tafsili.TafsiliName ,
                dbo.tblAcc_Tafsili.TafsiliId , ISNULL(dbo.tblAcc_Tafsili.RemainingAmount , '') AS RemainingAmount
                , ISNULL( tblAcc_Tafsili.SanadNo , '') AS SanadNo , ISNULL( tblAcc_Tafsili.AccountYear , '') AS AccountYear
		FROM    dbo.tblAcc_Tafsili
				JOIN dbo.tblAcc_Tafsili_Atf ON dbo.tblAcc_Tafsili.TafsiliId = dbo.tblAcc_Tafsili_Atf.TafsiliId 
		WHERE dbo.tblAcc_Tafsili_Atf.AtfId=@AtfId


ORDER BY dbo.tblAcc_Tafsili.TafsiliId
    END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Delete_tblAcc_Group]'
GO

CREATE PROCEDURE [dbo].[Delete_tblAcc_Group] (
@GroupID int
) AS
DELETE [tblAcc_Group]
WHERE
	[GroupID] = @GroupID


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazKol6]'
GO
SET QUOTED_IDENTIFIER OFF
GO



CREATE PROCEDURE [dbo].[Get_All_TarazKol6](@AccountYear smallint, @Branch int, @KolId1 int = 0, @KolId2 int = 0, @d1 int = 0, @d2 int = 0) AS
SELECT     tblAcc_Kol.KolId, MAX(tblAcc_Kol.KolName) AS KolName, SUM(bd1) AS fbd1, SUM(bs1) AS fbs1, SUM(bd2) AS fbd2, SUM(bs2) AS fbs2
FROM tblAcc_Kol INNER JOIN
(SELECT     tblAcc_DocumentDetail.KolId, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd1, SUM(tblAcc_DocumentDetail.Bestankar) AS bs1, 0 AS bd2, 0 AS bs2
FROM         tblAcc_DocumentHeader INNER JOIN
                                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
							tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
							tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate < @d1) AND ((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2))
GROUP BY tblAcc_DocumentDetail.KolId
UNION ALL
SELECT     tblAcc_DocumentDetail.KolId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2
FROM         tblAcc_DocumentHeader INNER JOIN
                                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
							tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
							tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND ((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2))
GROUP BY tblAcc_DocumentDetail.KolId) t
ON tblAcc_Kol.KolId = t.KolId
GROUP BY tblAcc_Kol.KolId
ORDER BY tblAcc_Kol.KolId



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_NextSanadHeader]'
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROC [dbo].[Get_NextSanadHeader](
@AccountYear smallint,
@Branch INt,
@DocumentId INT
)
as

begin

SELECT TOP 1
DocumentDate,
DocumentId,
DocumentDes,
State,
DocumentId2,
DocumentKind 
FROM tblAcc_DocumentHeader 
WHERE DocumentId>@DocumentId AND [AccountYear]=@AccountYear AND [Branch]=@Branch

 
END
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TarazKol8]'
GO

CREATE PROCEDURE [dbo].[Get_All_TarazKol8]
(
@AccountYear smallint, 
@Branch int, 
@KolId1 int = 0, 
@KolId2 int = 0, 
@d1 int = 0, 
@d2 int = 0
) 
AS
SELECT     tblAcc_Kol.KolId, 
MAX(tblAcc_Kol.KolName) AS KolName, 
SUM(t.bd1) AS fbd1, 
SUM(t.bs1) AS fbs1, 
SUM(t.bd2) AS fbd2, 
SUM(t.bs2) AS fbs2, 
SUM(t.bd3) AS fbd3, 
SUM(t.bs3) AS fbs3
FROM         tblAcc_Kol INNER JOIN
                          (SELECT KolId, CASE WHEN SUM(Bedehkar) >= SUM(Bestankar) THEN SUM(Bedehkar)-SUM(Bestankar) ELSE 0 END AS bd1, CASE WHEN SUM(Bestankar)>=SUM(Bedehkar) THEN SUM(Bestankar)-SUM(Bedehkar) ELSE 0 END AS bs1, 0 AS bd2, 0 AS bs2, 0 AS bd3, 0 AS bs3 
                             FROM         tblAcc_DocumentHeader INNER JOIN
                                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
							tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
							tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
                             WHERE (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear < @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND ((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2))
                             GROUP BY tblAcc_DocumentDetail.KolId
                            
							 Union All

                             SELECT     tblAcc_DocumentDetail.KolId, 0 AS bd1, 0 AS bs1, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd2, SUM(tblAcc_DocumentDetail.Bestankar) AS bs2, 0 AS bd3, 0 AS bs3
                             FROM         tblAcc_DocumentHeader INNER JOIN
                                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
							tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
							tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
                             WHERE (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate < @d1) AND ((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2))
                             GROUP BY tblAcc_DocumentDetail.KolId
                             
							 Union All
							 
                             SELECT     tblAcc_DocumentDetail.KolId, 0 AS bd1, 0 AS bs1, 0 AS bd2, 0 AS bs2, SUM(tblAcc_DocumentDetail.Bedehkar) AS bd3, SUM(tblAcc_DocumentDetail.Bestankar) AS bs3
                             FROM         tblAcc_DocumentHeader INNER JOIN
                                                  tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND
							tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND
							tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId
                             WHERE     (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.AccountyEar = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2) AND ((@KolId1 = 0) OR (tblAcc_DocumentDetail.KolId BETWEEN @KolId1 AND @KolId2))
                             GROUP BY tblAcc_DocumentDetail.KolId) t ON tblAcc_Kol.KolId = t.KolId
GROUP BY tblAcc_Kol.KolId
ORDER BY tblAcc_Kol.KolId


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Groups_ByPK_GroupID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Groups_ByPK_GroupID] (
			
			
			@GroupID int
				
		) AS
		
		SELECT 
		
		
				[GroupID],
				[GroupName],
				[Active]
		
		FROM 
		
		[tblAcc_Group]
		
		WHERE
		
		
			[GroupID] = @GroupID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_KharjChequeByChequeNo]'
GO
CREATE PROC [dbo].[Get_KharjChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_RecieveSanad] 
WHERE [CheckNo]=@ChequeNo AND [RecieveTypeId]=1
AND [CheckNo]<>0
END 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_FirstSanadHeader]'
GO

CREATE PROC [dbo].[Get_FirstSanadHeader](
@AccountYear smallint,
@Branch INt
)
as

begin

SELECT 
DocumentDate,
DocumentId,
DocumentDes,
State,
DocumentId2,
DocumentKind 
FROM tblAcc_DocumentHeader 
WHERE DocumentId=
(SELECT MIN(DocumentId) FROM tblAcc_DocumentHeader WHERE AccountYear=@AccountYear and Branch=@Branch)

END

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_TafsilihayeBedoonGardesh]'
GO

CREATE PROCEDURE [dbo].[Get_All_TafsilihayeBedoonGardesh](@Branch int) AS
SELECT     *
FROM         tblAcc_Tafsili
WHERE     (Branch = @Branch) AND (NOT (TafsiliId IN
                          (SELECT     TafsiliId
                             FROM         tblAcc_DocumentDetail
                             WHERE     tblAcc_Tafsili.Branch = tblAcc_DocumentDetail.Branch)))
ORDER BY TafsiliId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_MoeinId]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_MoeinId] (
			
			
			@KolID int,
			@MoeinId int
				
		) AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[AtfID]
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[KolID] = @KolID AND 
			[MoeinId] = @MoeinId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_VosouliChequeByChequeNo]'
GO
CREATE PROC [dbo].[Get_VosouliChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_RecieveSanad] 
WHERE [CheckNo]=@ChequeNo AND [RecieveTypeId]=3
END 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Atfs_Paged]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Atfs_Paged] (
	@__PageNumber int,
	@__PageSize int
)

AS
		
	DECLARE @Start int, @End int

	BEGIN TRANSACTION GetDataSet

	SET @Start = (((@__PageNumber - 1) * @__PageSize) + 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	SET @End = (@Start + @__PageSize - 1)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler

	CREATE TABLE #TemporaryTable (
		Row int IDENTITY(1,1) PRIMARY KEY,
		AtfID int, AtfName nvarchar(50), Active bit
	)
	
	IF @@ERROR <> 0
			GOTO ErrorHandler

	INSERT INTO #TemporaryTable
		
		SELECT AtfID, AtfName, Active
		
	FROM [tblAcc_Atf] 

	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	SELECT AtfID, AtfName, Active
		FROM #TemporaryTable
		WHERE (Row >= @Start) AND (Row <= @End)
	
	IF @@ERROR <> 0
		GOTO ErrorHandler
	
	DROP TABLE #TemporaryTable
	
	COMMIT TRANSACTION GetDataSet
	RETURN 0

ErrorHandler:
ROLLBACK TRANSACTION GetDataSet
RETURN @@ERROR
	
	




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_MoeinId_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_MoeinId_Count] (
			
			
			@KolID int,
			@MoeinId int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[KolID] = @KolID AND 
			[MoeinId] = @MoeinId




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_TafsiliAtfByKolMoein]'
GO
Create Proc [dbo].[Update_TafsiliAtfByKolMoein]
(
@KolID int,
@MoeinID int,
@TafsiliID Int,
@Branch int
)
As
Begin
INSERT INTO [dbo].[tblAcc_Tafsili_Atf]
           ([Branch]
           ,[TafsiliId]
           ,[AtfId])

SELECT @Branch
      ,@TafsiliID
      ,[AtfId]
  FROM [dbo].[tblAcc_Moein_Atf]
Where MoeinId=@MoeinID ANd KolId=@KolID

End



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_KolPrint]'
GO

CREATE PROCEDURE [dbo].[Get_All_KolPrint] AS
SELECT     tblAcc_Kol.KolID, tblAcc_Group.GroupName, tblAcc_Kol.KolName, tblAcc_Kol.Active
FROM         tblAcc_Group INNER JOIN
                      tblAcc_Kol ON tblAcc_Group.GroupID = tblAcc_Kol.GroupID
ORDER BY tblAcc_Kol.KolID

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[ado]'
GO

CREATE VIEW [dbo].[ado]
AS
SELECT TOP 100 PERCENT * FROM (
SELECT     DocumentId, dbo.ConvIntToDateFormat(MAX(dt)) AS sdate, SUM(Bedehkar) AS Bedehkar, SUM(Bestankar) AS Bestankar, MAX(RowDes) AS RowDes, 
                      KolId, MoeinId, TafsiliId, MAX(DocumentDate) AS DocumentDate, MAX(t1) AS KolMoeinName, kind
FROM         (SELECT     0 AS RowId, 0 AS DocumentId, 0 AS dt, SUM(tblAcc_DocumentDetail.Bedehkar) AS Bedehkar, SUM(tblAcc_DocumentDetail.Bestankar) AS Bestankar, 
                                              N'ب' AS RowDes, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS DocumentDate, MAX(tblAcc_Kol.KolName) + ' - ' + MAX(tblAcc_Moein.MoeinName) AS t1, 0 AS kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = 1384) AND (tblAcc_DocumentHeader.Branch = 1) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < 13840101)
                        GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, RowId
                        UNION ALL
                        SELECT     RowId, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate AS dt, tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar, tblAcc_DocumentDetail.RowDes, 
                                              dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, tblAcc_DocumentHeader.DocumentDate, tblAcc_Kol.KolName + ' - ' + tblAcc_Moein.MoeinName AS t1, dbo.tblAcc_DocumentDetail.kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = 1384) AND (tblAcc_DocumentHeader.Branch = 1) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN 13840101 AND 13840730)) t
WHERE ((KolId = 0) OR (0 = 0)) AND ((MoeinId = 0) OR (0 = 0))
GROUP BY DocumentId, KolId, MoeinId, TafsiliId, Kind) dt
ORDER BY KolId, MoeinId, TafsiliId, DocumentDate, DocumentId, kind




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moeins_ForKolId]'
GO




CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moeins_ForKolId](@KolId int)
				
		AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[MoeinName],
				[Active]
		
		FROM 
		
		[tblAcc_Moein]
		
		WHERE 
		
		[KolId] = @KolId





GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID_Count] (
			
			
			@KolID int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[KolID] = @KolID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_LastSanadHeader]'
GO

CREATE PROC [dbo].[Get_LastSanadHeader](
@AccountYear smallint,
@Branch INt
)
as

begin

SELECT 
DocumentDate,
DocumentId,
DocumentDes,
State,
DocumentId2,
DocumentKind 
FROM tblAcc_DocumentHeader 
WHERE DocumentId=
(SELECT MAX(DocumentId) FROM tblAcc_DocumentHeader WHERE AccountYear=@AccountYear and Branch=@Branch)

END


GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_VagozariChequeByChequeNo]'
GO
CREATE PROC [dbo].[Get_VagozariChequeByChequeNo]
(
@ChequeNo INT
)
AS
BEGIN
SELECT * FROM [dbo].[tblAcc_RecieveSanad] 
WHERE [CheckNo]=@ChequeNo AND [RecieveTypeId]=1 AND [CheckNo]<>0
END 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Update_tblAcc_Tafsili_SanadNo]'
GO

CREATE  PROCEDURE dbo.Update_tblAcc_Tafsili_SanadNo
(
	@SanadNo	INT,
	@RemainingAmount	INT,
	@AccountYear SMALLINT ,
	@Branch  INT ,
	@Tafsili INT 

) 

AS

	UPDATE 	dbo.tblAcc_Tafsili
		SET 	SanadNo = @SanadNo ,
				RemainingAmount = @RemainingAmount ,
				AccountYear = @AccountYear
	    	          WHERE   TafsiliId = @Tafsili And Branch = @Branch

GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblAcc_Moeins_For_KolId]'
GO
CREATE PROCEDURE [dbo].[Get_All_tblAcc_Moeins_For_KolId](@KolId int)
				
		AS
		
		SELECT 
		
		
				[KolID],
				(SELECT [KolName] FROM [dbo].[tblAcc_Kol] WHERE [KolId]=[dbo].[tblAcc_Moein].[KolId])AS KolName,
				[MoeinId],
				[MoeinName],
				[Kind],
				[Active]
		
		FROM 
		
		[tblAcc_Moein]
		
		WHERE 
		
		[KolId] = @KolId
		
		ORDER BY 
		[dbo].[tblAcc_Moein].[KolId],
		[MoeinId]
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_KolID] (
			
			
			@KolID int
				
		) AS
		
		SELECT 
		
		
				[KolID],
				[MoeinId],
				[AtfID]
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[KolID] = @KolID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_TafsiliNameById]'
GO
CREATE PROC [dbo].[Get_TafsiliNameById]
(
@TafsiliId int
)
as
Begin
Select  TafsiliName from tblacc_Tafsili WHere TafsiliId=@TafsiliId

end 
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_AtfTafsili]'
GO
CREATE	PROC [dbo].[Get_All_AtfTafsili]
AS
BEGIN
SELECT *,
(SELECT COUNT(*) FROM [dbo].[tblAcc_Tafsili_Atf] WHERE [AtfId]=[dbo].[tblAcc_Atf].[AtfId] AND [TafsiliId]=[dbo].[tblAcc_Tafsili].[TafsiliId])AS Relation
 FROM [dbo].[tblAcc_Tafsili]

Cross JOIN [dbo].[tblAcc_Atf]
ORDER BY [dbo].[tblAcc_Tafsili].[TafsiliId],[dbo].[tblAcc_Atf].[AtfName]


end
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_DaftarMoein]'
GO

CREATE PROCEDURE [dbo].[Get_All_DaftarMoein](@AccountYear smallint, @Branch int, @KolId int, @MoeinId int, @d1 int, @d2 int, @title nvarchar(255)) AS
SELECT TOP 100 PERCENT * FROM (
SELECT     DocumentId, dbo.ConvIntToDateFormat(MAX(dt)) AS sdate, SUM(Bedehkar) AS Bedehkar, SUM(Bestankar) AS Bestankar, MAX(RowDes) AS RowDes, 
                      KolId, MoeinId, TafsiliId, MAX(DocumentDate) AS DocumentDate, MAX(t1) AS KolMoeinName, kind
FROM         (SELECT     0 AS RowId, 0 AS DocumentId, 0 AS dt, SUM(tblAcc_DocumentDetail.Bedehkar) AS Bedehkar, SUM(tblAcc_DocumentDetail.Bestankar) AS Bestankar, 
                                              @title AS RowDes, tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, 0 AS DocumentDate, MAX(tblAcc_Kol.KolName) + ' - ' + MAX(tblAcc_Moein.MoeinName) AS t1, 0 AS kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate < @d1)
                        GROUP BY tblAcc_DocumentDetail.KolId, tblAcc_DocumentDetail.MoeinId, tblAcc_DocumentDetail.TafsiliId, RowId
                        UNION ALL
                        SELECT     RowId, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate AS dt, tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar, tblAcc_DocumentDetail.RowDes, 
                                              dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, tblAcc_DocumentHeader.DocumentDate, tblAcc_Kol.KolName + ' - ' + tblAcc_Moein.MoeinName AS t1, dbo.tblAcc_DocumentDetail.kind
                        FROM         tblAcc_DocumentHeader INNER JOIN
                                              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId INNER JOIN
                                              tblAcc_Kol ON tblAcc_DocumentDetail.KolId = tblAcc_Kol.KolId INNER JOIN
                                              tblAcc_Moein ON tblAcc_DocumentDetail.KolId = tblAcc_Moein.KolId AND tblAcc_DocumentDetail.MoeinId = tblAcc_Moein.MoeinId INNER JOIN
                                              tblAcc_Tafsili ON tblAcc_DocumentDetail.TafsiliId = tblAcc_Tafsili.TafsiliId
                        WHERE     (tblAcc_DocumentHeader.AccountYear = @AccountYear) AND (tblAcc_DocumentHeader.Branch = @Branch) AND (tblAcc_DocumentHeader.state > 1) AND (tblAcc_DocumentHeader.DocumentDate BETWEEN @d1 AND @d2)) t
WHERE ((KolId = @KolId) OR (@KolId = 0)) AND ((MoeinId = @MoeinId) OR (@MoeinId = 0))
GROUP BY DocumentId, KolId, MoeinId, TafsiliId, Kind) dt
ORDER BY KolId, MoeinId, TafsiliId, DocumentDate, DocumentId, kind



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Moein_Atfs_ByFK_AtfID_Count]'
GO




-----------------------------------------------------------------------------------
-- This stored procedure was auto-generated by nTierGen.NET Framework Generator v1.5
-- web: http://www.nTierGen.NET/
-- email: mailto:support@gavinjoyce.com
-- forums: http://www.gavinjoyce.com/forums/
-----------------------------------------------------------------------------------
		
		CREATE PROCEDURE [dbo].[Get_tblAcc_Moein_Atfs_ByFK_AtfID_Count] (
			
			
			@AtfID int
				
		) AS
		
		SELECT 
		
			COUNT(*)
		
		FROM 
		
		[tblAcc_Moein_Atf]
		
		WHERE
		
		
			[AtfID] = @AtfID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_DocumentDetail_Received]'
GO



CREATE PROCEDURE [dbo].[Get_tblAcc_DocumentDetail_Received]
(@AccountYear smallint, @Branch int, @KolId int, @MoeinId int, @TafsiliId int, @SaleNoAcc int) AS

  SELECT     RowId, tblAcc_DocumentHeader.DocumentId, tblAcc_DocumentHeader.DocumentDate AS dt, tblAcc_DocumentDetail.Bedehkar, tblAcc_DocumentDetail.Bestankar, tblAcc_DocumentDetail.RowDes, 
          dbo.tblAcc_DocumentDetail.KolId, dbo.tblAcc_DocumentDetail.MoeinId, dbo.tblAcc_DocumentDetail.TafsiliId, tblAcc_DocumentHeader.DocumentDate, 
          dbo.tblAcc_DocumentDetail.kind
        FROM  tblAcc_DocumentHeader INNER JOIN
              tblAcc_DocumentDetail ON tblAcc_DocumentHeader.AccountYear = tblAcc_DocumentDetail.AccountYear AND tblAcc_DocumentHeader.Branch = tblAcc_DocumentDetail.Branch AND tblAcc_DocumentHeader.DocumentId = tblAcc_DocumentDetail.DocumentId 
        WHERE tblAcc_DocumentHeader.AccountYear = @AccountYear 
			AND tblAcc_DocumentHeader.Branch = @Branch 
			AND tblAcc_DocumentHeader.state > 1 
			AND KolId = @KolId AND MoeinId = @MoeinId AND TafsiliId = @TafsiliId 
			AND Refrence_Khazane = @SaleNoAcc
	ORDER BY RowId
	
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_tblAcc_Kol_ByID_Count]'
GO




CREATE PROCEDURE [dbo].[Get_tblAcc_Kol_ByID_Count] (
		 	
				
		@KolID int

		) AS
		
		SELECT 
		
		
				COUNT([KolID]) AS ct
		
		FROM 
		
		[tblAcc_Kol]
		
		WHERE
		
		
			[KolID] = @KolID




GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[tblAcc_TurnType]'
GO
CREATE TABLE [dbo].[tblAcc_TurnType]
(
[TurnTypeId] [int] NOT NULL,
[Descs] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
)
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating primary key [PK_tblAcc_TurnType] on [dbo].[tblAcc_TurnType]'
GO
ALTER TABLE [dbo].[tblAcc_TurnType] ADD CONSTRAINT [PK_tblAcc_TurnType] PRIMARY KEY CLUSTERED  ([TurnTypeId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Creating [dbo].[Get_All_tblTotal_AccountType_Active]'
GO


CREATE  PROCEDURE [dbo].[Get_All_tblTotal_AccountType_Active] AS
	SELECT * FROM tblTotal_AccountType WHERE Active = 1



GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding constraints to [dbo].[tblAcc_Atf]'
GO
ALTER TABLE [dbo].[tblAcc_Atf] ADD CONSTRAINT [IX_tAtfGroup] UNIQUE NONCLUSTERED  ([AtfId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding constraints to [dbo].[tblAcc_Kol]'
GO
ALTER TABLE [dbo].[tblAcc_Kol] ADD CONSTRAINT [IX_tblAcc_Kol] UNIQUE NONCLUSTERED  ([KolId])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_PaymentSanad]'
GO
ALTER TABLE [dbo].[tblAcc_PaymentSanad] WITH NOCHECK  ADD CONSTRAINT [FK_tblAcc_PaymentSanad_tblAcc_PayType] FOREIGN KEY ([PaymentTypeId]) REFERENCES [dbo].[tblAcc_PayType] ([PaymentTypeId]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_RecieveSanad]'
GO
ALTER TABLE [dbo].[tblAcc_RecieveSanad] ADD CONSTRAINT [FK_tblAcc_RecieveSanad_tblAcc_Bank] FOREIGN KEY ([BankNo]) REFERENCES [dbo].[tblAcc_Bank] ([tintBank])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_Moein_Atf]'
GO
ALTER TABLE [dbo].[tblAcc_Moein_Atf] ADD CONSTRAINT [FK_tblAcc_Moein_Atf_tblAcc_Atf] FOREIGN KEY ([AtfId]) REFERENCES [dbo].[tblAcc_Atf] ([AtfId]) ON UPDATE CASCADE
ALTER TABLE [dbo].[tblAcc_Moein_Atf] ADD CONSTRAINT [FK_tblAcc_Moein_Atf_tblAcc_Kol] FOREIGN KEY ([KolId]) REFERENCES [dbo].[tblAcc_Kol] ([KolId]) ON DELETE CASCADE ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_Tafsili_Atf]'
GO
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] ADD CONSTRAINT [FK_tblAcc_Tafsili_Atf_tblAcc_Atf] FOREIGN KEY ([AtfId]) REFERENCES [dbo].[tblAcc_Atf] ([AtfId])
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] ADD CONSTRAINT [FK_tblAcc_Tafsili_Atf_tblAcc_Tafsili] FOREIGN KEY ([Branch], [TafsiliId]) REFERENCES [dbo].[tblAcc_Tafsili] ([Branch], [TafsiliId]) ON DELETE CASCADE ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_DocumentDetail]'
GO
ALTER TABLE [dbo].[tblAcc_DocumentDetail] ADD CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_DocumentHeader] FOREIGN KEY ([AccountYear], [Branch], [DocumentId]) REFERENCES [dbo].[tblAcc_DocumentHeader] ([AccountYear], [Branch], [DocumentId]) ON DELETE CASCADE ON UPDATE CASCADE
ALTER TABLE [dbo].[tblAcc_DocumentDetail] ADD CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_Moein] FOREIGN KEY ([KolId], [MoeinId]) REFERENCES [dbo].[tblAcc_Moein] ([KolId], [MoeinId]) ON UPDATE CASCADE
ALTER TABLE [dbo].[tblAcc_DocumentDetail] ADD CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_Tafsili] FOREIGN KEY ([Branch], [TafsiliId]) REFERENCES [dbo].[tblAcc_Tafsili] ([Branch], [TafsiliId]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_DocumentHeader]'
GO
ALTER TABLE [dbo].[tblAcc_DocumentHeader] ADD CONSTRAINT [FK_tblAcc_DocumentHeader_tUser] FOREIGN KEY ([UserId]) REFERENCES [dbo].[tUser] ([UID]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_Kol]'
GO
ALTER TABLE [dbo].[tblAcc_Kol] ADD CONSTRAINT [FK_tblAcc_Kol_tblAcc_Group] FOREIGN KEY ([GroupId]) REFERENCES [dbo].[tblAcc_Group] ([GroupId])
ALTER TABLE [dbo].[tblAcc_Kol] ADD CONSTRAINT [FK_tblAcc_Kol_tblAcc_KolShenaseh] FOREIGN KEY ([ShenaseId]) REFERENCES [dbo].[tblAcc_KolShenaseh] ([ShenaseId]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_Moein]'
GO
ALTER TABLE [dbo].[tblAcc_Moein] ADD CONSTRAINT [FK_tblAcc_Moein_tblAcc_Kol] FOREIGN KEY ([KolId]) REFERENCES [dbo].[tblAcc_Kol] ([KolId]) ON DELETE CASCADE ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_PaymentSanad]'
GO
ALTER TABLE [dbo].[tblAcc_PaymentSanad] ADD CONSTRAINT [FK_tblAcc_CheckBook_tblAcc_PaymentSanad] FOREIGN KEY ([CheckBookId]) REFERENCES [dbo].[tblAcc_CheckBook] ([CheckBookId]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_Tafsili]'
GO
ALTER TABLE [dbo].[tblAcc_Tafsili] ADD CONSTRAINT [FK_tblAcc_Tafsili_tBranch] FOREIGN KEY ([Branch]) REFERENCES [dbo].[tBranch] ([Branch]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tblAcc_CheckBook]'
GO
ALTER TABLE [dbo].[tblAcc_CheckBook] ADD CONSTRAINT [FK_tblAcc_CheckBook_tblAcc_ChequePrintTemplate] FOREIGN KEY ([PrintTemplateID]) REFERENCES [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID])
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tSalaryD]'
GO
ALTER TABLE [dbo].[tSalaryD] ADD CONSTRAINT [FK_tSalaryM_tSalaryD] FOREIGN KEY ([SalaryId], [Branch]) REFERENCES [dbo].[tSalaryM] ([SalaryId], [Branch]) ON DELETE CASCADE ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
PRINT N'Adding foreign keys to [dbo].[tSalaryM]'
GO
ALTER TABLE [dbo].[tSalaryM] ADD CONSTRAINT [FK_tSalaryM_tAccountYear] FOREIGN KEY ([AccountYear]) REFERENCES [dbo].[tAccountYears] ([AccountYear]) ON UPDATE CASCADE
GO
IF @@ERROR<>0 AND @@TRANCOUNT>0 ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT=0 BEGIN INSERT INTO #tmpErrors (Error) SELECT 1 BEGIN TRANSACTION END
GO
IF EXISTS (SELECT * FROM #tmpErrors) ROLLBACK TRANSACTION
GO
IF @@TRANCOUNT>0 BEGIN
PRINT 'The database update succeeded'
COMMIT TRANSACTION
END
ELSE PRINT 'The database update failed'
GO
DROP TABLE #tmpErrors
GO
