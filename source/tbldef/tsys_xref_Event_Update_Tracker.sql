CREATE TABLE [tsys_xref_Event_Update_Tracker] (
  [ID] AUTOINCREMENT,
  [Event_ID] GUID ,
  [Import_Event_ID] GUID ,
  [AppendTableName] VARCHAR (75),
  [Record_Count] LONG ,
  [Date] DATETIME 
)
