CREATE TABLE [_tbl_Sapling_Conditions_Import_20170906_PRIMARY] (
  [Sapling_Condition_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Sapling_Data_ID] VARCHAR (50),
  [Condition] VARCHAR (50),
  [Updated_Date] DATETIME 
)
