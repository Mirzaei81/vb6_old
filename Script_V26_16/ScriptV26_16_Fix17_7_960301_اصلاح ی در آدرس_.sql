
--اصلاح  ک و  ی   در جستجوی آدرس مشترکین


alter   Proc Get_Customer_Address
@ActDeact int ,
@Address Nvarchar(200) 
as    


Set @Address = Replace(  @Address  , N'ك', N'ک'  ) 
Set @Address = Replace(  @Address  , N'ي' , N'ی' )


Select [Code] ,MembershipId ,[Name] ,Telephone , Address , MasterCode , Prefix , Tafsili , Credit
 from dbo.vw_Get_Cust 
where  CHARINDEX ( @Address , [Address] ) > 0 and actdeact <> @ActDeact
AND vw_Get_Cust.Code <> -1 --AND Branch = @Branch


GO


