CREATE TABLE [TaxonomicReference] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ShortCitation] VARCHAR (25),
  [LongCitation] VARCHAR (150),
  [EffectiveDate] DATETIME ,
  [RetireDate] DATETIME 
)
