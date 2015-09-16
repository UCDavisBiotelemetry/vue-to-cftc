SELECT *
FROM DetectOrphans AS Orphans
WHERE NOT EXISTS (SELECT 1 FROM Import_Deployments AS Dep WHERE Orphans.VR2SN = Dep.VR2SN AND Orphans.DetectDate BETWEEN Dep.[Start] AND Dep.[Stop]);
