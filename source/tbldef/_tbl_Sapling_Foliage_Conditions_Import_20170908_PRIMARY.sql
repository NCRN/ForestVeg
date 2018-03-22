CREATE TABLE [_tbl_Sapling_Foliage_Conditions_Import_20170908_PRIMARY] (
  [Sapling_Foliage_Condition_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Sapling_Data_ID] VARCHAR (50),
  [Condition] VARCHAR (2),
  [Percent_Afflicted] SINGLE ,
  [Updated_Date] DATETIME 
)
