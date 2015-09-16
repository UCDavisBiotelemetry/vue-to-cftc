SELECT DISTINCT DetectOrphans.VR2SN, MIN(DetectDate) AS FirstO, MAX(DetectDate) AS LastO, COUNT(*) AS totalMissed
FROM DetectOrphans
GROUP BY VR2SN;
