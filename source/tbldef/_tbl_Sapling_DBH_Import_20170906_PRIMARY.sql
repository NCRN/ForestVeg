CREATE TABLE [_tbl_Sapling_DBH_Import_20170906_PRIMARY] (
  [Sapling_DBH_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Sapling_Data_ID] VARCHAR (50),
  [DBH] SINGLE ,
  [Live] BIT ,
  [Updated_Date] DATETIME 
)
