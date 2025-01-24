
-- tInventory  InventoryNo = 99 ÓÇÎÊå ÔæÏ

--- 64 ==> 99    

--- 63 ==> 64    

--- 99 ==> 63    


UPDATE tFacd 
SET intInventoryNo = 99 WHERE intInventoryNo = 64
GO


UPDATE tFacd 
SET DestInventoryNo = 99 WHERE DestInventoryNo = 64
GO


UPDATE tFacd 
SET intInventoryNo = 64 WHERE intInventoryNo = 63
GO


UPDATE tFacd 
SET DestInventoryNo = 64 WHERE DestInventoryNo = 63
GO

UPDATE tFacd 
SET intInventoryNo = 63 WHERE intInventoryNo = 99
GO


UPDATE tFacd 
SET DestInventoryNo = 63 WHERE DestInventoryNo = 99
GO

