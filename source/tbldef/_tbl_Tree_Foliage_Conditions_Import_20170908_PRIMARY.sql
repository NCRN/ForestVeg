CREATE TABLE [_tbl_Tree_Foliage_Conditions_Import_20170908_PRIMARY] (
  [Tree_Foliage_Condition_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Tree_Data_ID] VARCHAR (50),
  [Condition] VARCHAR (2),
  [Percent_Afflicted] SINGLE ,
  [Updated_Date] DATETIME 
)
