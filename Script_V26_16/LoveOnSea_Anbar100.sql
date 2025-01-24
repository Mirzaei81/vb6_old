

UPDATE tfacd 
SET intInventoryNo =1
FROM(
SELECT tfacd.intSerialNo , tfacd.Branch , tfacd.intInventoryNo FROM tfacm INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
 WHERE status = 1 AND  intInventoryNo = 100
)T1
WHERE tfacd.intSerialNo = T1.intSerialNo AND tfacd.Branch = T1.Branch AND T1.intInventoryNo =100


GO


UPDATE tfacd 
SET intInventoryNo =1
FROM(
SELECT tfacd.intSerialNo , tfacd.Branch , tfacd.intInventoryNo FROM tfacm INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
 WHERE status = 6 AND  intInventoryNo = 100
)T2
WHERE tfacd.intSerialNo = T2.intSerialNo AND tfacd.Branch = T2.Branch AND T2.intInventoryNo =100


GO

UPDATE tfacd 
SET DestInventoryNo =1
FROM(
SELECT tfacd.intSerialNo , tfacd.Branch , tfacd.DestInventoryNo FROM tfacm INNER JOIN tfacd ON dbo.tFacD.Branch = dbo.tFacM.Branch AND dbo.tFacD.intSerialNo = dbo.tFacM.intSerialNo
 WHERE status = 7 AND  DestInventoryNo = 100
)T3
WHERE tfacd.intSerialNo = T3.intSerialNo AND tfacd.Branch = T3.Branch AND T3.DestInventoryNo =100


GO


