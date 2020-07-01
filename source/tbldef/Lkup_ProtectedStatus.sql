CREATE TABLE [Lkup_ProtectedStatus] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Code] VARCHAR (2),
  [Name] VARCHAR (25),
  [Description] LONGTEXT 
)
