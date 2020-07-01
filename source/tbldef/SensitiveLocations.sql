CREATE TABLE [SensitiveLocations] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Park_ID] LONG ,
  [Location_ID] LONG  CONSTRAINT [Location_ID] UNIQUE ,
  [CreateDate] DATETIME ,
  [CreatedBy_ID] LONG ,
  [LastModified] DATETIME ,
  [LastModifiedBy_ID] LONG 
)
