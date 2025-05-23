
--ScriptV26_16_Fix18_1_960401_اصلاح سازمانی.sql
--تولید صف انتظار از دستگاه اثرانگشت
--بارگذاری اتوماتیک پرسنل از لیست
--امکان انصراف پرسنل از لیست
--

IF NOT EXISTS(SELECT * FROM tblPub_Script2 WHERE [Version] = 26 AND Script = 16 AND FixNumber = 17 )

	INSERT  INTO dbo.tblPub_Script2
			( Version ,
			  Script ,
			  LastScriptNo ,
			  [Date] ,
			  FixNumber
			)
	VALUES  ( 26 ,
			  16 ,
			  0 ,
			  dbo.shamsi(GETDATE()) ,
			  18
			)
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_Queue]') and OBJECTPROPERTY(id, N'IsTable') = 1)
DROP TABLE Arya_Kitchen_Queue
GO

CREATE  table Arya_Kitchen_Queue
(
PK_ID int not null ,
MembershipID int not null,
nvcDate NVARCHAR(8) NOT NULL ,
nvcTime NVARCHAR(5) NOT Null,
Status INT ,
nvcOrderTime NVARCHAR(8) NULL ,
nvcEscTime NVARCHAR(8) NULL ,
nvcExitTime NVARCHAR(8) NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].Arya_Kitchen_Queue ADD CONSTRAINT [PK_Arya_Kitchen_Queue] PRIMARY KEY CLUSTERED  ([PK_ID] ) ON [PRIMARY]
GO

------------------------------------------

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_Status]') and OBJECTPROPERTY(id, N'IsTable') = 1)
DROP TABLE Arya_Kitchen_Status
GO

CREATE  table Arya_Kitchen_Status
(
StatusNo int not null,
StatusDescription NVARCHAR(20) NOT NULL 
) ON [PRIMARY]
GO


INSERT INTO dbo.Arya_Kitchen_Status
        ( StatusNo ,
          StatusDescription
        )
VALUES  ( 0 , -- StatusNo - int
          N'ورود به لیست'  -- StatusDescription - nvarchar(20)
        )
GO
INSERT INTO dbo.Arya_Kitchen_Status
        ( StatusNo ,
          StatusDescription
        )
VALUES  ( 1 , -- StatusNo - int
          N'ثبت سفارش'  -- StatusDescription - nvarchar(20)
        )
GO
INSERT INTO dbo.Arya_Kitchen_Status
        ( StatusNo ,
          StatusDescription
        )
VALUES  ( 2 , -- StatusNo - int
          N' انصراف '  -- StatusDescription - nvarchar(20)
        )
GO
INSERT INTO dbo.Arya_Kitchen_Status
        ( StatusNo ,
          StatusDescription
        )
VALUES  ( 3 , -- StatusNo - int
          N'خروج'  -- StatusDescription - nvarchar(20)
        )
GO

------------------------------------------
------------------------------------------

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_InsertNewRequest]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Arya_Kitchen_InsertNewRequest
GO


Create procedure Arya_Kitchen_InsertNewRequest
@ID as  INT 
as
Declare @MaxID as int 
set @MaxID =  (select  isnull(max(PK_ID),0)+1 from Arya_Kitchen_Queue)

insert into Arya_Kitchen_Queue (PK_ID,MembershipID, nvcDate , nvcTime , Status)values (@MaxID,@ID, dbo.Get_ShamsiDate_For_Current_Shift(GETDATE()), dbo.setTimeFormat(getdate()) ,0)
 

 Go


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_GetRequests]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Arya_Kitchen_GetRequests
GO

Create procedure Arya_Kitchen_GetRequests
 
as
DECLARE @nvcDate NVARCHAR(8)
SET @nvcDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())

select * from Arya_Kitchen_Queue
inner join tcust on Arya_Kitchen_Queue.MembershipID =tcust.MembershipId
where Status = 0 AND nvcDate = @nvcDate
order by PK_ID


 GO
 

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_ChangeStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Arya_Kitchen_ChangeStatus
GO

Create procedure Arya_Kitchen_ChangeStatus
@PK_ID INT ,
@StatusNo int 
as
UPDATE Arya_Kitchen_Queue
SET Status = @StatusNo WHERE PK_ID = @PK_ID

IF @StatusNo = 1
UPDATE Arya_Kitchen_Queue
SET nvcOrderTime = dbo.setTimeFormat(getdate())
WHERE PK_ID = @PK_ID
ELSE IF @StatusNo = 2
UPDATE Arya_Kitchen_Queue
SET nvcEscTime = dbo.setTimeFormat(getdate())
WHERE PK_ID = @PK_ID


 GO

 
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_GetPerson_ById]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Arya_Kitchen_GetPerson_ById
GO

Create procedure Arya_Kitchen_GetPerson_ById
@PK_ID INT  
as
select * from Arya_Kitchen_Queue
inner join tcust on Arya_Kitchen_Queue.MembershipID =tcust.MembershipId
where PK_ID = @PK_ID
GO


 
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Arya_Kitchen_GetFirstPerson]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].Arya_Kitchen_GetFirstPerson
GO

Create procedure Arya_Kitchen_GetFirstPerson

AS

DECLARE @nvcDate NVARCHAR(8)
SET @nvcDate = dbo.Get_ShamsiDate_For_Current_Shift(GETDATE())

SELECT TOP 1 * from Arya_Kitchen_Queue
inner join tcust on Arya_Kitchen_Queue.MembershipID =tcust.MembershipId
where nvcDate = @nvcDate AND Status = 0
ORDER BY PK_ID

GO



ALTER   view vw_Customers

as
SELECT     dbo.tCust.Code, dbo.tCust.MembershipId, dbo.tCust.MasterCode, dbo.tCust.Owner, dbo.tCust.Name, dbo.tCust.Family, dbo.tCust.Sex, 
                      dbo.tCust.WorkName, dbo.tCust.State, dbo.tCust.City,dbo.tCust.ActKind, dbo.tCust.ActDeAct, 
                      dbo.tCust.Prefix, dbo.tCust.Assansor, dbo.tCust.Address, dbo.tCust.PostalCode, dbo.tCust.Tel1, dbo.tCust.Tel2, dbo.tCust.Tel3, 
                      dbo.tCust.Tel4, dbo.tCust.Mobile, dbo.tCust.Fax, dbo.tCust.Email, dbo.tCust.CarryFee, dbo.tCust.PaykFee, dbo.tCust.Distance, dbo.tCust.Discount, 
                      dbo.tCust.BuyState, dbo.tCust.Credit, dbo.tCust.Description, dbo.tCust.[Date], dbo.tCust.[Time], dbo.tCust.[User], dbo.tCust.Unit, dbo.tCust.InternalNo, 
                      dbo.tCust.Flour, ISNULL(SUM(T.sumPrice), 0) AS Price , IsNull(T2.Bestankar,0) As Bestankar ,  CASE 
			WHEN (dbo.tCust.MasterCode is null And dbo.tCust.WorkName ='' and dbo.tCust.Name <> '')
				then   Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name When 1 then N' آقای '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name ELSE N'  '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name END
			WHEN (dbo.tCust.MasterCode is null And dbo.tCust.WorkName ='' and dbo.tCust.Name = '')
				then  Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' '  When 1 then N' آقای '  +  dbo.tcust.Family + ' '  ELSE N' '  +  dbo.tcust.Family + ' '  END
			When (dbo.tCust.MasterCode is null And dbo.tCust.WorkName <>'')
				then dbo.tCust.WorkName
			WHEN (dbo.tCust.MasterCode is not null And tCust_1.WorkName <>'' and dbo.tCust.Name <> '')
				then tCust_1.WorkName + '_' +  Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name When 1 then N' آقای '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name ELSE N' '  +  dbo.tcust.Family + ' ' +  dbo.tcust.Name END
			WHEN (dbo.tCust.MasterCode is not null And tCust_1.WorkName <>'' and dbo.tCust.Name = '')
				then tCust_1.WorkName + '_' +  Case  dbo.tCust.Sex When 0 then N' خانم '  +  dbo.tcust.Family + ' '  When 1 then N' آقای '  +  dbo.tcust.Family + ' '  ELSE N' '  +  dbo.tcust.Family + ' '  END

			End as FullName , case
			
			WHEN (dbo.tCust.MasterCode is null )
				Then
					dbo.tCust.Address
			WHEN (dbo.tCust.MasterCode is not null )
				Then
					tCust_1.Address + N' طبقه ' + isnull(dbo.tCust.Flour , '') + N' واحد ' + isnull(dbo.tCust.Unit , '')
	--				dbo.tCust.Address + ' ' + tCust_1.Address + N' طبقه ' + isnull(dbo.tCust.Flour , '') + N' واحد ' + isnull(dbo.tCust.Unit , '')
			end as FullAddress  , tCust.FamilyNo , tCust.Member , tCust.Central , tCust.SellPrice , tCust.[Branch] , tcust.Tafsili
FROM         dbo.tCust LEFT OUTER JOIN  dbo.tCust tCust_1  on dbo.tCust.MasterCode = tCust_1.Code and dbo.tCust.Branch = tCust_1.Branch
			  
			LEFT OUTER JOIN
                          (SELECT     Sum(IsNull(Bestankar,0)) As Bestankar , Code_Bes
                             FROM         dbo.tblAcc_Recieved
                             WHERE     dbo.tblAcc_Recieved.RecieveType = 3 And AccountYear = dbo.Get_AccountYear() Group By Code_Bes) T2  ON dbo.tCust.Code = T2.Code_Bes  
                      
			 LEFT OUTER JOIN
                          (SELECT    Case Status When 2 then  sumPrice When 5 then  -Sumprice end as sumprice, Customer  ,  Branch
                             FROM         dbo.tFacM
                             WHERE     dbo.tFacm.Balance = 0 And dbo.tFacm.FacPayment = 1 /* and dbo.tFacm.Branch = dbo.Get_Current_Branch()*/ ) T ON dbo.tCust.Code = T.Customer --and dbo.tCust.Branch = T.Branch
--WHERE [tCust].[Branch] = dbo.[Get_Current_Branch]()

GROUP BY dbo.tCust.Code, dbo.tCust.MembershipId, dbo.tCust.MasterCode, dbo.tCust.Owner, dbo.tCust.Name, dbo.tCust.Family, dbo.tCust.Sex, 
          dbo.tCust.WorkName, dbo.tCust.State, dbo.tCust.City, dbo.tCust.ActKind, dbo.tCust.ActDeAct, 
          dbo.tCust.Prefix, dbo.tCust.Assansor, dbo.tCust.Address, dbo.tCust.PostalCode, dbo.tCust.Tel1, dbo.tCust.Tel2, dbo.tCust.Tel3, 
          dbo.tCust.Tel4, dbo.tCust.Mobile, dbo.tCust.Fax, dbo.tCust.Email, dbo.tCust.CarryFee, dbo.tCust.PaykFee, dbo.tCust.Distance, dbo.tCust.Discount, 
          dbo.tCust.BuyState, dbo.tCust.Credit, dbo.tCust.Description, dbo.tCust.[Date], dbo.tCust.[Time], dbo.tCust.[User], dbo.tCust.Unit, dbo.tCust.InternalNo, 
          dbo.tCust.Flour , tCust_1.WorkName , tCust_1.Address , tCust.FamilyNo , tCust.Member , T2.Bestankar , tCust.Central , tCust.SellPrice , tCust.[Branch]
			, dbo.tCust.Tafsili


GO

