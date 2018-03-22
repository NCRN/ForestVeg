CREATE TABLE [_tbl_Tree_Conditions_Import_20170908_PRIMARY] (
  [Tree_Condition_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Tree_Data_ID] VARCHAR (50),
  [Condition] VARCHAR (50),
  [Updated_Date] DATETIME 
)
