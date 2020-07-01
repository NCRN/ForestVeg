CREATE TABLE [Flags] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [FlagCategory] VARCHAR (255),
  [FlagGroup] VARCHAR (100),
  [FlagType] VARCHAR (25),
  [FlagName] VARCHAR (100),
  [Code] VARCHAR (255),
  [NumericCode] LONG ,
  [Label] VARCHAR (255),
  [IsResolvable] BYTE ,
  [EffectiveDate] DATETIME ,
  [RetireDate] DATETIME 
)
