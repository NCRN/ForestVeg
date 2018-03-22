CREATE TABLE [tsys_Append_Log] (
  [ID] AUTOINCREMENT,
  [Table_Name] VARCHAR (50),
  [Append_Date] DATETIME ,
  [Append_Table_Name] VARCHAR (128),
  [Append_Records] LONG ,
  [Record_ID] VARCHAR (50)
)
