CREATE TABLE [DataFlags] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [RecordTable] VARCHAR (255),
  [RecordID] LONG ,
  [RecordField] VARCHAR (255),
  [FlagID] LONG ,
  [LastUpdate] DATETIME ,
  [CreateDate] DATETIME ,
  [ClearDate] DATETIME ,
  [CreatedBy] LONG ,
  [LastUpdateBy] LONG ,
  [ClearedBy] LONG 
)
