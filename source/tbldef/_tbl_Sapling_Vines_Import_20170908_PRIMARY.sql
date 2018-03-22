CREATE TABLE [_tbl_Sapling_Vines_Import_20170908_PRIMARY] (
  [Sapling_Vine_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Sapling_Data_ID] VARCHAR (50),
  [TSN] LONG ,
  [Updated_Date] DATETIME 
)
