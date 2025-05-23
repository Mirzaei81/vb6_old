
--حذف نمایش فیشهای مرجوعی از فرم گردش حساب مشتریان


ALTER   Procedure [dbo].[Get_CreditFactor] ( @Customer Bigint , @AccountYear SMALLINT , @Branch INT 
	, @DateBefore NVARCHAR(10) , @DateAfter NVARCHAR(10)) 

 AS

select dbo.vw_CreditFactor.* 
	from dbo.vw_CreditFactor 
	where dbo.vw_CreditFactor.Customer = @Customer  -- And (dbo.vw_CreditFactor.Incharge  Is  Not Null  Or dbo.vw_CreditFactor.OrderType = 2) 
		AND vw_CreditFactor.Balance = 0
         And (dbo.vw_CreditFactor.Facpayment = 1 OR  (dbo.vw_CreditFactor.Facpayment = 0 AND InCharge IS NULL AND ServePlace <> 2 ))
         AND AccountYear = @AccountYear 
         AND (Branch = @Branch OR @Branch = 0)
		AND Date >= @DateBefore AND Date <= @DateAfter
		AND [vw_CreditFactor].Recursive = 0



GO
