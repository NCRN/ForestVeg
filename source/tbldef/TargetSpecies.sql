CREATE TABLE [TargetSpecies] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [TargetList] VARCHAR (255),
  [TSN] LONG ,
  [EstablishDate] DATETIME ,
  [RetireDate] DATETIME 
)
