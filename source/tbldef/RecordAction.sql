CREATE TABLE [RecordAction] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ReferenceType] VARCHAR (25),
  [Reference_ID] LONG ,
  [Contact_ID] LONG ,
  [Activity] VARCHAR (2),
  [ActionDate] DATETIME 
)
