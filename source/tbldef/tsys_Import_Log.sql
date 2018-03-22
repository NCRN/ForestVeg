CREATE TABLE [tsys_Import_Log] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Table_Name] VARCHAR (128),
  [Import_Date] DATETIME ,
  [Import_Records] LONG ,
  [Delete_Table] BIT ,
  [Delete_Date] DATETIME 
)
