CREATE TABLE [SOP] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [FullName] VARCHAR (100),
  [Code] VARCHAR (20),
  [SOPNumber] SHORT ,
  [Version] DOUBLE ,
  [EffectiveDate] DATETIME ,
  [RetireDate] DATETIME ,
  [CreateDate] DATETIME ,
  [CreatedBy_ID] LONG ,
  [LastModified] DATETIME ,
  [LastModifiedBy_ID] LONG 
)
