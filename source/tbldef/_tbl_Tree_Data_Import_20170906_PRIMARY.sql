CREATE TABLE [_tbl_Tree_Data_Import_20170906_PRIMARY] (
  [Tree_Data_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID] VARCHAR (50),
  [Tag_ID] VARCHAR (50),
  [Crown_Class] LONG ,
  [Wind_Lightning_Damage] BIT ,
  [Status] VARCHAR (32),
  [Tree_Status] VARCHAR (32),
  [Vines_Checked] BIT ,
  [Conditions_Checked] BIT ,
  [Foliage_Conditions_Checked] BIT ,
  [Tree_Notes] LONGTEXT ,
  [Updated_Date] DATETIME ,
  [TreeVigor] SHORT 
)
