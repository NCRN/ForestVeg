CREATE TABLE [tsys_db_components_new] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ComponentName] VARCHAR (255),
  [ComponentType] VARCHAR (255),
  [ComponentFrom] VARCHAR (255),
  [ComponentVersion] VARCHAR (255),
  [LastUpdate] DATETIME 
)
