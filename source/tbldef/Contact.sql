CREATE TABLE [Contact] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [LastName] VARCHAR (50),
  [FirstName] VARCHAR (25),
  [MiddleInitial] VARCHAR (4),
  [Organization] VARCHAR (50),
  [PositionTitle] VARCHAR (50),
  [Email] VARCHAR (50),
  [WorkPhone] SINGLE ,
  [WorkExtension] SHORT ,
  [IsActive] BYTE ,
  [IsNPS] BYTE ,
  [Username] VARCHAR (75)
)
