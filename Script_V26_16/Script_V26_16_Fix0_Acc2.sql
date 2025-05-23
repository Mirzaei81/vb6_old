/*
Run this script on:

(local).Total_V26_16    -  This database will be modified

to synchronize it with:

(local).Total_V26_16_1

You are recommended to back up your database before running this script

Script created by SQL Data Compare version 10.4.8 from Red Gate Software Ltd at 2014/01/05 11:26:53 ب.ظ

*/
		
GO
SET NUMERIC_ROUNDABORT OFF
GO
SET ANSI_PADDING, ANSI_WARNINGS, CONCAT_NULL_YIELDS_NULL, ARITHABORT, QUOTED_IDENTIFIER, ANSI_NULLS, NOCOUNT ON
GO
SET DATEFORMAT YMD
GO
SET XACT_ABORT ON
GO
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
GO
BEGIN TRANSACTION
-- Pointer used for text / image updates. This might not be needed, but is declared here just in case
DECLARE @pv binary(16)

PRINT(N'Drop constraints from [dbo].[tblAcc_Moein_Atf]')
GO
ALTER TABLE [dbo].[tblAcc_Moein_Atf] DROP CONSTRAINT [FK_tblAcc_Moein_Atf_tblAcc_Atf]
ALTER TABLE [dbo].[tblAcc_Moein_Atf] DROP CONSTRAINT [FK_tblAcc_Moein_Atf_tblAcc_Kol]

PRINT(N'Drop constraints from [dbo].[tblAcc_Moein]')
GO
ALTER TABLE [dbo].[tblAcc_Moein] DROP CONSTRAINT [FK_tblAcc_Moein_tblAcc_Kol]

PRINT(N'Drop constraint FK_tblAcc_DocumentDetail_tblAcc_Moein from [dbo].[tblAcc_DocumentDetail]')
GO
ALTER TABLE [dbo].[tblAcc_DocumentDetail] DROP CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_Moein]

PRINT(N'Drop constraints from [dbo].[tblAcc_Tafsili_Atf]')
GO
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] DROP CONSTRAINT [FK_tblAcc_Tafsili_Atf_tblAcc_Atf]
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] DROP CONSTRAINT [FK_tblAcc_Tafsili_Atf_tblAcc_Tafsili]

PRINT(N'Drop constraints from [dbo].[tblAcc_Kol]')
GO
ALTER TABLE [dbo].[tblAcc_Kol] DROP CONSTRAINT [FK_tblAcc_Kol_tblAcc_Group]
ALTER TABLE [dbo].[tblAcc_Kol] DROP CONSTRAINT [FK_tblAcc_Kol_tblAcc_KolShenaseh]

PRINT(N'Drop constraints from [dbo].[tAccess_Object]')
GO
ALTER TABLE [dbo].[tAccess_Object] DROP CONSTRAINT [FK_tAccess_Form_tAccessLevel]
ALTER TABLE [dbo].[tAccess_Object] DROP CONSTRAINT [FK_tAccess_Form_tForms]

PRINT(N'Drop constraints from [dbo].[tblAcc_Tafsili]')
GO
ALTER TABLE [dbo].[tblAcc_Tafsili] DROP CONSTRAINT [FK_tblAcc_Tafsili_tBranch]

PRINT(N'Drop constraint FK_tblAcc_DocumentDetail_tblAcc_Tafsili from [dbo].[tblAcc_DocumentDetail]')
GO
ALTER TABLE [dbo].[tblAcc_DocumentDetail] DROP CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_Tafsili]

PRINT(N'Drop constraint FK_tblAcc_PaymentSanad_tblAcc_PayType from [dbo].[tblAcc_PaymentSanad]')
GO
ALTER TABLE [dbo].[tblAcc_PaymentSanad] DROP CONSTRAINT [FK_tblAcc_PaymentSanad_tblAcc_PayType]

PRINT(N'Drop constraint FK_tblAcc_CheckBook_tblAcc_ChequePrintTemplate from [dbo].[tblAcc_CheckBook]')
GO
ALTER TABLE [dbo].[tblAcc_CheckBook] DROP CONSTRAINT [FK_tblAcc_CheckBook_tblAcc_ChequePrintTemplate]

PRINT(N'Drop constraint FK_tblAcc_RecieveSanad_tblAcc_Bank from [dbo].[tblAcc_RecieveSanad]')
GO
ALTER TABLE [dbo].[tblAcc_RecieveSanad] DROP CONSTRAINT [FK_tblAcc_RecieveSanad_tblAcc_Bank]

PRINT(N'Add rows to [dbo].[tblAcc_Atf]')
GO
INSERT INTO [dbo].[tblAcc_Atf] ([AtfId], [AtfName], [Active]) VALUES (1, N'بانكها', 1)
INSERT INTO [dbo].[tblAcc_Atf] ([AtfId], [AtfName], [Active]) VALUES (2, N'سهامداران وپرسنل', 1)
INSERT INTO [dbo].[tblAcc_Atf] ([AtfId], [AtfName], [Active]) VALUES (3, N'اشخاص وموسسات وشركتها', 1)
INSERT INTO [dbo].[tblAcc_Atf] ([AtfId], [AtfName], [Active]) VALUES (4, N'عمومي ', 1)
PRINT(N'Operation applied to 4 rows out of 4')
GO

PRINT(N'Add rows to [dbo].[tblAcc_Bank]')
GO
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (5, N'تجارت')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (6, N'اقتصاد نوين')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (7, N'كشاورزي')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (8, N'سامان')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (9, N'رفاه كارگران')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (10, N'پارسيان')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (11, N'تات')
INSERT INTO [dbo].[tblAcc_Bank] ([tintBank], [nvcBankName]) VALUES (12, N'گردشگري')
PRINT(N'Operation applied to 8 rows out of 8')
GO

PRINT(N'Add rows to [dbo].[tblAcc_ChequePrintTemplate]')
GO
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (1, N'بانك ملي', N'Melli.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (2, N'بانك ملت', N'Mellat.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (3, N'بانك صادرات', N'Saderat.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (4, N'بانك تجارت', N'Tejarat.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (5, N'بانك سامان', N'Saman.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (6, N'بانك كشاورزي', N'Keshavarzi.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (7, N'بانك پارسيان', N'Parsian.rpt', 1)
INSERT INTO [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID], [Name], [Path], [Active]) VALUES (8, N'بانك اقتصادنوين', N'EghtesadNovin.rpt', 1)
PRINT(N'Operation applied to 8 rows out of 8')
GO

PRINT(N'Add rows to [dbo].[tblAcc_Group]')
GO
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (1, N'دارائيهاي جاري', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (2, N'دارائيهاي ثابت', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (3, N'ساير دارائيها', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (4, N'هزينه ها', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (5, N'بدهيهاي جاري', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (7, N'حقوق صاحبان سهام', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (8, N'عملكرد', 1)
INSERT INTO [dbo].[tblAcc_Group] ([GroupId], [GroupName], [Active]) VALUES (9, N'ساير حسابها', 1)
PRINT(N'Operation applied to 8 rows out of 8')
GO

PRINT(N'Add rows to [dbo].[tblAcc_PayType]')
GO
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (1, N'چك خام')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (2, N'چك پرداختي')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (3, N'چك پرداختي وصول شده')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (4, N'چك برگشتي')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (5, N'چک عودت شده')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (6, N'چك باطل شده')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (7, N'پرداخت به اشخاص و شركتها')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (8, N'پرداخت نقدي از حساب به حساب')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (9, N'پرداخت چك از حساب به صندوق')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (10, N'پرداخت نقد از حساب به صندوق')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (11, N'پرداخت نقد از صندوق به صندوق')
INSERT INTO [dbo].[tblAcc_PayType] ([PaymentTypeId], [PaymentTypeName]) VALUES (12, N'پرداخت نقد از صندوق به حساب')
PRINT(N'Operation applied to 12 rows out of 12')
GO

PRINT(N'Add rows to [dbo].[tblAcc_RecieveType]')
GO
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (1, N'چك دريافتي')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (2, N'دريافتي خرج شده')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (3, N'در جريان وصول')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (4, N'وصول شده')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (5, N'برگشتي')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (6, N'برگشت به مشتري')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (7, N'واريز نقدي به حساب')
INSERT INTO [dbo].[tblAcc_RecieveType] ([RecieveTypeId], [ReceiveTypeName]) VALUES (8, N'دريافت نقدي')
PRINT(N'Operation applied to 8 rows out of 8')
GO

DELETE FROM dbo.TblAcc_Sale
GO

PRINT(N'Add rows to [dbo].[TblAcc_Sale]')
GO
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (1, N'فروش ', 81, 8101, 0, 1, N'فروش كلي')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (2, N'تخفيفات فروش', 85, 8502, 0, 1, N'تخفيفات نقدي فروش')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (3, N'درآمد بسته بندي', 82, 8201, 0, 1, N'درآمد')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (4, N'درآمدكرايه حمل', 83, 8301, 0, 1, N'درآمد')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (5, N'اسناد پرداختني', 53, 5301, 0, 1, N'اسناد پرداختني')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (6, N'بدهكاران', 14, 1401, 0, 1, N'تجاري')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (7, N'بدهي صندوق', 11, 1101, 0, 1, N'صندوق')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (8, N'بدهي فروشنده', 11, 1101, 0, 1, N'صندوق')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (9, N'جاري كاركنان', 15, 1501, 0, 1, N'جاري كاركنان')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (10, N'وام كاركنان', 15, 1502, 0, 1, N'وام كاركنان')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (11, N'مساعده كاركنان', 15, 1503, 0, 1, N'مساعده كاركنان')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (12, N'علي الحساب كاركنان', 15, 1504, 0, 1, N'علي الحساب كاركنان')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (13, N'حقوق و دستمزد', 44, 4401, 0, 1, N'حقوق پايه')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (14, N'هزينه هاي عمومي', 41, 4101, 0, 1, N'هزينه هاي عمومي')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (15, N'بستانكاران', 51, 5101, 0, 1, N'تجاري')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (16, N'موجودي كالا و مواد اوليه', 17, 1701, 0, 1, N'موجودي كالا')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (17, N'برگشت از فروش', 85, 8501, 0, 1, N'برگشت از فروش و تخفيفات')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (18, N'برگشت از خريد', 86, 8601, 0, 1, NULL)
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (19, N'اسناد دريافتني', 16, 1601, 0, 1, N'اسناددريافتني نزد صندوق')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (20, N'تعهد بن', 11, 1, 0, 1, NULL)
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (21, N'فروش اقساطي', 12, 1, 0, 1, NULL)
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (22, N'كارت اعتباري', 11, 1, 0, 1, NULL)
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (23, N'اسناد درجريان وصول', 11, 1104, 0, 1, N'چكهاي درجريان وصول')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (24, N'عوارض فروش', 57, 5701, 0, 1, N' عوارض ارزش افزوده فروش')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (25, N'عوارض خريد', 33, 3301, 0, 1, N'  عوارض ارزش افزوده خريد')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (26, N'ماليات فروش', 57, 5702, 0, 1, N'ماليات ارزش افزوده فروش')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (27, N'ماليات خريد', 33, 3302, 0, 1, N'ماليات ارزش افزوده خريد')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (28, N'حسابهاي بانكي', 11, 1103, 0, 1, N'موجودي بانك')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (29, N'تخفيفات خريد', 86, 8602, 0, 1, N'تخفيفات خريد')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (30, N'افتتاحيه', 92, 9201, 0, 1, N'افتتاحيه')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (31, N'اختتاميه', 93, 9301, 0, 1, N'اختتاميه')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (32, N'هزينه ضايعات', 43, 4304, 0, 1, N'هزينه ضايعات')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (33, N' خلاصه سود و زيان  ', 71, 7101, 0, 1, N' سود و زيان سال جاري')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (34, N'سود و زيان انباشته', 72, 7201, 0, 1, N'سود و زيان انباشته')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (35, N'موجودی اولیه', 17, 1702, 0, 1, N'موجودی اولیه')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (36, N'هزینه های مالی', 42, 4201, 0, 1, N'هزینه های مالی')
INSERT INTO [dbo].[TblAcc_Sale] ([Code], [Description], [Kol], [Moein], [Tafsili], [Active], [MoeinDesc]) VALUES (37, N'هزینه های توزیع و فروش', 43, 4301, 0, 1, N'هزینه های توزیع و فروش')
PRINT(N'Operation applied to 37 rows out of 37')
GO

PRINT(N'Add row to [dbo].[tblAcc_Tafsili]')
GO
INSERT INTO [dbo].[tblAcc_Tafsili] ([Branch], [TafsiliId], [TafsiliName], [Active], [AccountYear], [RemainingAmount], [SanadNo]) VALUES (1, 0, N'  ', 1, NULL, NULL, NULL)

PRINT(N'Add rows to [dbo].[tblAcc_TurnType]')
GO
INSERT INTO [dbo].[tblAcc_TurnType] ([TurnTypeId], [Descs]) VALUES (1, N'چك')
INSERT INTO [dbo].[tblAcc_TurnType] ([TurnTypeId], [Descs]) VALUES (2, N'فيش')
INSERT INTO [dbo].[tblAcc_TurnType] ([TurnTypeId], [Descs]) VALUES (3, N'حواله')
INSERT INTO [dbo].[tblAcc_TurnType] ([TurnTypeId], [Descs]) VALUES (4, N'رسيد')
PRINT(N'Operation applied to 4 rows out of 4')
GO

PRINT(N'Add row to [dbo].[tblAcc_UGroups]')
GO
INSERT INTO [dbo].[tblAcc_UGroups] ([UGroupId], [UGroupName]) VALUES (1, N'مديريت')

PRINT(N'Add rows to [dbo].[tObjects]')
GO
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (470, N'AccBaseAccountingDefine', N'تعاريف پايه حسابداري', N'AccBaseAccountingDefine', 1, NULL)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (471, N'AccfrmTafsili', N'تفضيلي ها', N'AccfrmTafsili', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (472, N'AccfrmGroup', N'گروه ها', N'AccfrmTafsili', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (473, N'AccfrmKol', N'حساب هاي كل', N'AccfrmKol', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (474, N'AccfrmAtfTafsili', N'عطف هاي تفضيلي', N'AccfrmAtfTafsili', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (475, N'AccfrmMoein', N'حساب هاي معين', N'AccfrmMoein', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (476, N'AccfrmAtfMoein', N'معين و عطف ها', N'AccfrmAtfMoein', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (477, N'AccfrmAccCoding', N'كد هاي حسابداري', N'AccfrmAccCoding', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (478, N'AccfrmAtf', N'عطف ها', N'AccfrmAtf', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (480, N'AccfrmBank', N'بانك ها', N'AccfrmBank', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (481, N'AccfrmBankAccount', N'حساب هاي بانكي', N'AccfrmBankAccount', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (482, N'AccfrmCheckBook', N'دسته چك ها', N'AccfrmCheckBook', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (483, N'AccReceived', N'دريافت ها', N'AccReceived', 1, NULL)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (484, N'AccfrmReceivedCash', N'دريافت وجه نقد', N'AccfrmReceivedCash', 1, 483)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (485, N'AccfrmReceivedCashFromAccount', N'دريافت وجه نقد از طريق واريز به حساب', N'AccfrmReceivedCashFromAccount', 1, 483)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (486, N'AccfrmReceivedCheck', N'دريافت چك', N'AccfrmReceivedCheck', 1, 483)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (487, N'AccfrmReceivedCashCheckAccount', N'دريافت چك ونقد و حساب', N'AccfrmReceivedCashCheckAccount', 1, 483)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (488, N'AccPayment', N'پرداخت ها', N'AccPayment', 1, NULL)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (489, N'AccfrmPaymentCash', N'پرداخت وجه نقد', N'AccfrmPaymentCash', 1, 488)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (490, N'AccfrmPaymentCheck', N'پرداخت چك', N'AccfrmPaymentCheck', 1, 488)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (491, N'AccfrmPayAccountToAccountCash', N'وجه نقد از حساب به حساب', N'AccfrmPayAccountToAccountCash', 1, 488)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (492, N'AccfrmPayAccountToSandoghCash', N'وجه نقد از حساب به صندوق', N'AccfrmPayAccountToSandoghCash', 1, 488)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (493, N'AccfrmPaySandoghToSandoghCash', N'وجه نقد از صندوق به صندوق', N'AccfrmPaySandoghToSandoghCash', 1, 488)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (494, N'AccfrmPaySandoghToAccountCash', N'وجه نقد از صندوق به حساب', N'AccfrmPaySandoghToAccountCash', 1, 488)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (495, N'AccCheckReceived', N'چك هاي دريافتني', N'AccCheckReceived', 1, NULL)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (496, N'AccfrmAllCheckReceived', N'نمايش چك ها', N'AccfrmAllCheckReceived', 1, 495)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (497, N'AccfrmCheckReceivedOperationKharj', N'خرج چك', N'AccfrmCheckReceivedOperationKharj', 1, 495)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (498, N'AccfrmCheckReceivedOperationVagozari', N'واگذاري چك به بانك', N'AccfrmCheckReceivedOperationVagozari', 1, 495)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (499, N'AccfrmCheckReceivedOperationVosouli', N'وصولي چك', N'AccfrmCheckReceivedOperationVosouli', 1, 495)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (500, N'AccfrmCheckReceivedOperationBargashti', N'برگشتي چك', N'AccfrmCheckReceivedOperationBargashti', 1, 495)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (501, N'AccfrmCheckReceivedOperationBargashtiMoshtari', N'برگشت به مشتري چك', N'AccfrmCheckReceivedOperationBargashtiMoshtari', 1, 495)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (502, N'AccCheckPayment', N'چك هاي پرداختني', N'AccCheckPayment', 1, NULL)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (503, N'AccfrmAllCheckPayment', N'نمايش چك ها', N'AccfrmAllCheckPayment', 1, 502)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (504, N'AccfrmCheckPaymentOperationVosouli', N'وصولي چك', N'AccfrmCheckPaymentOperationVosouli', 1, 502)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (505, N'AccfrmCheckPaymentOperationBargashtiMoshtari', N'برگشت از مشتري چك', N'AccfrmCheckPaymentOperationBargashtiMoshtari', 1, 502)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (506, N'AccfrmCheckPaymentEbtal', N'ابطال چك', N'AccfrmCheckPaymentEbtal', 1, 502)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (507, N'AccfrmAsnad', N'اسناد حسابداري', N'AccfrmAsnad', 1, 470)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (508, N'AccReport', N' گزارشات حسابداري ', N'AccReport', 1, NULL)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (509, N'AccfrmKartHesabReport', N'كارت حساب', N'AccfrmKartHesabReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (510, N'AccfrmTafsiliReport', N'حساب هاي تفضيلي', N'AccfrmTafsiliReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (511, N'AccfrmMoeinReport', N'حساب هاي معين', N'AccfrmMoeinReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (512, N'AccfrmKolReport', N'حساب هاي كل', N'AccfrmKolReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (513, N'AccfrmKolReport', N'حساب هاي كل', N'AccfrmKolReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (514, N'AccfrmTarazTafsiliReport', N'تراز حساب هاي تفضيلي', N'AccfrmTarazTafsiliReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (515, N'AccfrmTarazMoeinReport', N'تراز حساب هاي معين', N'AccfrmTarazMoeinReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (516, N'AccfrmTarazKolReport', N'تراز حساب هاي كل', N'AccfrmTarazKolReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (517, N'AccfrmDaftarKolReport', N'دفتر كل', N'AccfrmDaftarKolReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (518, N'AccfrmDaftarMoeinReport', N'دفتر معين', N'AccfrmDaftarMoeinReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (519, N'AccfrmJaoftadegiSanadNoReport', N'جا افتادگي در شماره اسناد', N'AccfrmJaoftadegiSanadNoReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (520, N'AccfrmJaoftadegiTarikhReport', N'جا افتادگي در تاريخ اسناد', N'AccfrmJaoftadegiTarikhReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (521, N'AccfrmDaftarKolRizReport', N'ريز دفتر كل', N'AccfrmDaftarKolRizReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (522, N'AccfrmDaftarRuznameReport', N'دفتر روزنامه', N'AccfrmDaftarRuznameReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (523, N'AccfrmDaftarRuznameRizReport', N'ريز دفتر روزنامه', N'AccfrmDaftarRuznameRizReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (524, N'AccfrmAsnadSummaryReport', N'گزارش خلاصه اسناد', N'AccfrmAsnadSummaryReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (525, N'AccfrmTafsiliNoGardeshReport', N'تفضيلي هاي بدون گردش', N'AccfrmTafsiliNoGardeshReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (526, N'AccfrmAsnadNoTarazReport', N'اسناد تراز نشده', N'AccfrmAsnadNoTarazReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (527, N'AccfrmCheckPaymentSarresidReport', N'چك هاي پرداختني سررسيد', N'AccfrmCheckPaymentSarresidReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (528, N'AccfrmCheckReceivedSarresidReport', N'چك هاي دريافتني سررسيد', N'AccfrmCheckReceivedSarresidReport', 1, 508)
INSERT INTO [dbo].[tObjects] ([intObjectCode], [ObjectId], [ObjectName], [objectLatinName], [intObjectType], [ObjectParent]) VALUES (542, N'SalarySystem', N'محاسبه حقوق و دستمزد', N'SalarySystem', 1, NULL)
PRINT(N'Operation applied to 59 rows out of 59')
GO

PRINT(N'Add rows to [dbo].[tAccess_Object]')
GO
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 470)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 471)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 472)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 473)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 474)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 475)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 476)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 477)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 478)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 480)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 481)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 482)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 483)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 484)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 485)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 486)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 487)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 488)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 489)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 490)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 491)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 492)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 493)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 494)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 495)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 496)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 497)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 498)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 499)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 500)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 501)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 502)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 503)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 504)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 505)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 506)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 507)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 508)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 509)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 510)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 511)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 512)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 513)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 514)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 515)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 516)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 517)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 518)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 519)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 520)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 521)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 522)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 523)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 524)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 525)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 526)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 527)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 528)
INSERT INTO [dbo].[tAccess_Object] ([intAccessLevel], [intObjectCode]) VALUES (1, 542)
PRINT(N'Operation applied to 59 rows out of 59')
GO

PRINT(N'Add rows to [dbo].[tblAcc_Kol]')
GO
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (11, 1, N'موجودي نقد و بانك', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (14, 1, N'حسابهاي دريافتني', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (15, 1, N'ساير حسابهاي دريافتني', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (16, 1, N'اسناد دريافتني', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (17, 1, N'موجودي مواد وكالا', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (21, 2, N'اموال ، ماشين آلات وتجهيزات', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (33, 3, N'ماليات وعوارض افزوده خريد', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (41, 4, N'هزينه هاي اداري وتشكيلاتي', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (42, 4, N'هزينه مالي', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (43, 4, N'هزينه هاي توزيع وفروش', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (44, 4, N'هزينه حقو ق ودستمزد', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (51, 5, N'حسابهاي پرداختني', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (53, 5, N'اسناد پرداختني', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (54, 5, N'جاري شركاء', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (57, 5, N'ماليات وعوارض افزوده فروش', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (81, 8, N'فروش', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (82, 8, N'درآمد بسته بندي', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (83, 8, N'درآمد كرايه حمل', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (84, 8, N'ساير درآمدها', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (85, 8, N'برگشت از فروش وتخفيفات', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (86, 8, N'برگشت از خريد وتخفيف', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (87, 8, N'بهاي تمام شده كالاي فروش رفته', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (90, 9, N'حسابهاي انتظامي', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (91, 9, N'طرف حسابهاي انتظامي', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (92, 9, N'افتتاحيه', 1, NULL)
INSERT INTO [dbo].[tblAcc_Kol] ([KolId], [GroupId], [KolName], [Active], [ShenaseId]) VALUES (93, 9, N'اختتاميه', 1, NULL)
PRINT(N'Operation applied to 26 rows out of 26')
GO

PRINT(N'Add rows to [dbo].[tblAcc_Tafsili_Atf]')
GO
INSERT INTO [dbo].[tblAcc_Tafsili_Atf] ([Branch], [TafsiliId], [AtfId]) VALUES (1, 0, 1)
INSERT INTO [dbo].[tblAcc_Tafsili_Atf] ([Branch], [TafsiliId], [AtfId]) VALUES (1, 0, 2)
INSERT INTO [dbo].[tblAcc_Tafsili_Atf] ([Branch], [TafsiliId], [AtfId]) VALUES (1, 0, 3)
INSERT INTO [dbo].[tblAcc_Tafsili_Atf] ([Branch], [TafsiliId], [AtfId]) VALUES (1, 0, 4)
PRINT(N'Operation applied to 4 rows out of 4')
GO

PRINT(N'Add rows to [dbo].[tblAcc_Moein]')
GO
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (11, 1101, N'صندوق', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (11, 1102, N'تنخواه گردان', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (11, 1103, N'موجودي بانك', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (11, 1104, N'چكهاي درجريان وصول', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (14, 1401, N'تجاري', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (15, 1501, N'جاري كاركنان', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (15, 1502, N'وام كاركنان', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (15, 1503, N'مساعده كاركنان', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (15, 1504, N'علي الحساب كاركنان', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (16, 1601, N'اسناددريافتني نزد صندوق', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (17, 1701, N'موجودي كالا', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (21, 2101, N'اثاثيه و لوازم', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (33, 3301, N'عوارض ارزش افزوده خريد', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (33, 3302, N'ماليات ارزش افزوده خريد', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (41, 4101, N'هزينه هاي عمومي', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (41, 4102, N'آبدارخانه وپذيرايي', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (41, 4103, N'مصارف وملزومات اداري', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (42, 4201, N'كارمزد بانكي', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (43, 4301, N'تبليغات وبازاريابي', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (43, 4302, N'هزينه پورسانت', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (43, 4303, N'حمل ونقل وپيك', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4401, N'حقوق پايه', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4403, N'بن وخواروبار', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4404, N'حق مسكن', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4405, N'اضافه كار', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4406, N'23% بيمه سهم كارفرما', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4407, N'حق عائله مندي', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4408, N'سنوات خدمت', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (44, 4409, N'عيدي وپاداش', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (51, 5101, N'تجاري', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (53, 5301, N'اسناد پرداختني', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (54, 5401, N'سهامداران', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (57, 5701, N'عوارض ارزش افزوده فروش', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (57, 5702, N'ماليات ارزش افزوده فروش', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (81, 8101, N'فروش كلي', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (82, 8201, N'درآمد', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (83, 8301, N'درآمد', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (85, 8501, N'برگشت از فروش و تخفيفات ', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (85, 8502, N'تخفيفات نقدي فروش ', 1, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (86, 8601, N'برگشت از خريد و تخفيفات ', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (86, 8602, N'تخفيفات نقدي خريد ', 2, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (90, 9001, N'اسناد ديگران نزد ما ', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (90, 9002, N'اسناد ما نزد ديگران ', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (91, 9101, N'اسناد ديگران نزد ما', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (91, 9102, N'اسناد مانزد ديگران', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (92, 9201, N'افتتاحيه ', 0, 1)
INSERT INTO [dbo].[tblAcc_Moein] ([KolId], [MoeinId], [MoeinName], [Kind], [Active]) VALUES (93, 9301, N'اختتاميه ', 0, 1)
PRINT(N'Operation applied to 47 rows out of 47')
GO

PRINT(N'Add rows to [dbo].[tblAcc_Moein_Atf]')
GO
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (11, 1101, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (11, 1102, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (11, 1103, 1)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (11, 1104, 1)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (14, 1401, 3)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (14, 1402, 3)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (15, 1501, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (15, 1502, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (15, 1503, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (15, 1504, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (16, 1601, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (16, 1601, 3)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (41, 4101, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (41, 4102, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (41, 4103, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (42, 4201, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (43, 4301, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (44, 4401, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (51, 5101, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (51, 5101, 3)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (53, 5301, 1)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (54, 5401, 2)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (81, 8101, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (82, 8201, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (83, 8301, 4)
INSERT INTO [dbo].[tblAcc_Moein_Atf] ([KolId], [MoeinId], [AtfId]) VALUES (85, 8501, 4)
PRINT(N'Operation applied to 26 rows out of 26')
GO

PRINT(N'Add rows to [dbo].[tblAcc_KolShenaseh]')
GO
INSERT INTO dbo.tblAcc_KolShenaseh
        ( ShenaseId, ShenaseName )
VALUES  ( 1, -- ShenaseId - int
          N'موقت '  -- ShenaseName - nvarchar(50)
          )
          
GO

INSERT INTO dbo.tblAcc_KolShenaseh
        ( ShenaseId, ShenaseName )
VALUES  ( 2, -- ShenaseId - int
          N'ترازنامه ای'  -- ShenaseName - nvarchar(50)
          )
          
GO

PRINT(N'Add constraints to [dbo].[tblAcc_Moein_Atf]')
GO
ALTER TABLE [dbo].[tblAcc_Moein_Atf] ADD CONSTRAINT [FK_tblAcc_Moein_Atf_tblAcc_Atf] FOREIGN KEY ([AtfId]) REFERENCES [dbo].[tblAcc_Atf] ([AtfId]) ON UPDATE CASCADE
ALTER TABLE [dbo].[tblAcc_Moein_Atf] ADD CONSTRAINT [FK_tblAcc_Moein_Atf_tblAcc_Kol] FOREIGN KEY ([KolId]) REFERENCES [dbo].[tblAcc_Kol] ([KolId]) ON DELETE CASCADE ON UPDATE CASCADE

PRINT(N'Add constraints to [dbo].[tblAcc_Moein]')
GO
ALTER TABLE [dbo].[tblAcc_Moein] ADD CONSTRAINT [FK_tblAcc_Moein_tblAcc_Kol] FOREIGN KEY ([KolId]) REFERENCES [dbo].[tblAcc_Kol] ([KolId]) ON DELETE CASCADE ON UPDATE CASCADE

PRINT(N'Add constraint FK_tblAcc_DocumentDetail_tblAcc_Moein to [dbo].[tblAcc_DocumentDetail]')
GO
ALTER TABLE [dbo].[tblAcc_DocumentDetail] WITH NOCHECK ADD CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_Moein] FOREIGN KEY ([KolId], [MoeinId]) REFERENCES [dbo].[tblAcc_Moein] ([KolId], [MoeinId]) ON UPDATE CASCADE

PRINT(N'Add constraints to [dbo].[tblAcc_Tafsili_Atf]')
GO
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] ADD CONSTRAINT [FK_tblAcc_Tafsili_Atf_tblAcc_Atf] FOREIGN KEY ([AtfId]) REFERENCES [dbo].[tblAcc_Atf] ([AtfId])
ALTER TABLE [dbo].[tblAcc_Tafsili_Atf] ADD CONSTRAINT [FK_tblAcc_Tafsili_Atf_tblAcc_Tafsili] FOREIGN KEY ([Branch], [TafsiliId]) REFERENCES [dbo].[tblAcc_Tafsili] ([Branch], [TafsiliId]) ON DELETE CASCADE ON UPDATE CASCADE

PRINT(N'Add constraints to [dbo].[tblAcc_Kol]')
GO
ALTER TABLE [dbo].[tblAcc_Kol] ADD CONSTRAINT [FK_tblAcc_Kol_tblAcc_Group] FOREIGN KEY ([GroupId]) REFERENCES [dbo].[tblAcc_Group] ([GroupId])
ALTER TABLE [dbo].[tblAcc_Kol] ADD CONSTRAINT [FK_tblAcc_Kol_tblAcc_KolShenaseh] FOREIGN KEY ([ShenaseId]) REFERENCES [dbo].[tblAcc_KolShenaseh] ([ShenaseId]) ON UPDATE CASCADE

PRINT(N'Add constraints to [dbo].[tAccess_Object]')
GO
ALTER TABLE [dbo].[tAccess_Object] WITH NOCHECK ADD CONSTRAINT [FK_tAccess_Form_tAccessLevel] FOREIGN KEY ([intAccessLevel]) REFERENCES [dbo].[tAccessLevel] ([intAccessLevel]) ON DELETE CASCADE ON UPDATE CASCADE
ALTER TABLE [dbo].[tAccess_Object] WITH NOCHECK ADD CONSTRAINT [FK_tAccess_Form_tForms] FOREIGN KEY ([intObjectCode]) REFERENCES [dbo].[tObjects] ([intObjectCode]) ON DELETE CASCADE ON UPDATE CASCADE

PRINT(N'Add constraints to [dbo].[tblAcc_Tafsili]')
GO
ALTER TABLE [dbo].[tblAcc_Tafsili] ADD CONSTRAINT [FK_tblAcc_Tafsili_tBranch] FOREIGN KEY ([Branch]) REFERENCES [dbo].[tBranch] ([Branch]) ON UPDATE CASCADE

PRINT(N'Add constraint FK_tblAcc_DocumentDetail_tblAcc_Tafsili to [dbo].[tblAcc_DocumentDetail]')
GO
ALTER TABLE [dbo].[tblAcc_DocumentDetail] WITH NOCHECK ADD CONSTRAINT [FK_tblAcc_DocumentDetail_tblAcc_Tafsili] FOREIGN KEY ([Branch], [TafsiliId]) REFERENCES [dbo].[tblAcc_Tafsili] ([Branch], [TafsiliId]) ON UPDATE CASCADE

PRINT(N'Add constraint FK_tblAcc_PaymentSanad_tblAcc_PayType to [dbo].[tblAcc_PaymentSanad]')
GO
ALTER TABLE [dbo].[tblAcc_PaymentSanad] WITH NOCHECK ADD CONSTRAINT [FK_tblAcc_PaymentSanad_tblAcc_PayType] FOREIGN KEY ([PaymentTypeId]) REFERENCES [dbo].[tblAcc_PayType] ([PaymentTypeId]) ON UPDATE CASCADE

PRINT(N'Add constraint FK_tblAcc_CheckBook_tblAcc_ChequePrintTemplate to [dbo].[tblAcc_CheckBook]')
GO
ALTER TABLE [dbo].[tblAcc_CheckBook] WITH NOCHECK ADD CONSTRAINT [FK_tblAcc_CheckBook_tblAcc_ChequePrintTemplate] FOREIGN KEY ([PrintTemplateID]) REFERENCES [dbo].[tblAcc_ChequePrintTemplate] ([PrintTemplateID])

PRINT(N'Add constraint FK_tblAcc_RecieveSanad_tblAcc_Bank to [dbo].[tblAcc_RecieveSanad]')
GO
ALTER TABLE [dbo].[tblAcc_RecieveSanad] WITH NOCHECK ADD CONSTRAINT [FK_tblAcc_RecieveSanad_tblAcc_Bank] FOREIGN KEY ([BankNo]) REFERENCES [dbo].[tblAcc_Bank] ([tintBank])
COMMIT TRANSACTION
GO
