CREATE TABLE [Import_Locations] (
	[Region] varchar(50) NULL,
	[Location_Long] varchar(50) NULL,
	[Location] varchar(50) NOT NULL,
	[Lat] float NULL,
	[Lon] float NULL,
	[Fatho_Depth] float NULL,
	[Chart_Depth] float NULL,
	[RiverKm] float NULL,
	[LocationType] varchar(50) NULL,
	[General_Location] varchar(50) NULL,
	[Nearest_Access] varchar(50) NULL,
	[Responsible_Agent] varchar(50) NULL,
	[Agent_Phone] varchar(255) NULL,
	[Agent_Email] varchar(255) NULL,
 CONSTRAINT [PK_tbl_Monitor_Locations] PRIMARY KEY
(
	[Location]
))