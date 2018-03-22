CREATE TABLE [_tbl_Locations_Import_20170908_PRIMARY] (
  [Location_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Plot_Name] VARCHAR (100),
  [Unit_Code] VARCHAR (12),
  [Unit_Group] VARCHAR (12),
  [Subunit_Code] VARCHAR (12),
  [New_Subunit] VARCHAR (12),
  [Unit_Note] VARCHAR (64),
  [Admin_Unit_Code] VARCHAR (12),
  [X_Coord] DOUBLE ,
  [Y_Coord] DOUBLE ,
  [Coord_Units] VARCHAR (50),
  [Coord_System] VARCHAR (50),
  [UTM_Zone] VARCHAR (50),
  [Datum] VARCHAR (50),
  [Location_Notes] LONGTEXT ,
  [Location_Status] VARCHAR (16),
  [Panel] LONG ,
  [Soil_Panel] LONG ,
  [Frame] VARCHAR (16),
  [GRTS_Order] DOUBLE ,
  [Install_Date] DATETIME ,
  [Lon_WGS84] DOUBLE ,
  [Lat_WGS84] DOUBLE ,
  [X_Coord_Access] DOUBLE ,
  [Y_Coord_Access] DOUBLE ,
  [Lon_WGS84_Access] DOUBLE ,
  [Lat_WGS84_Access] DOUBLE ,
  [Slope] LONG ,
  [Aspect] VARCHAR (50),
  [Location_Directions] LONGTEXT ,
  [Updated_Date] DATETIME ,
  [ShowLocMsg] BIT ,
  [LocMessage] LONGTEXT 
)
