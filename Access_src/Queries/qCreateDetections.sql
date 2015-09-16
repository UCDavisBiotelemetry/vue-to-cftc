CREATE TABLE [Import_Detections](
	[TagID] int NOT NULL,
	[Codespace] varchar(25) NOT NULL,
	[DetectDate] datetime NOT NULL,
	[VR2SN] int NOT NULL,
	[Data] float NULL,
	[Units] varchar(50) NULL,
	[Data2] float NULL,
	[Units2] varchar(50) NULL,
 CONSTRAINT [PK_DT] PRIMARY KEY  
(
	[TagID],
	[Codespace],
	[DetectDate],
	[VR2SN]
))