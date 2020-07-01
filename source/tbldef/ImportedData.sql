CREATE TABLE [ImportedData] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ImportDate] DATETIME ,
  [SourceFile] VARCHAR (50),
  [DestinationTable] VARCHAR (25),
  [NumberOfRecordsImported] SHORT ,
  [StartRecord_ID] LONG ,
  [EndRecord_ID] LONG 
)
