CREATE TABLE [SensitiveSpecies] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Park_ID] LONG ,
  [Master_PLANT_Code] VARCHAR (20),
  [CreateDate] DATETIME ,
  [CreatedBy_ID] LONG ,
  [LastModified] DATETIME ,
  [LastModifiedBy_ID] LONG 
)
