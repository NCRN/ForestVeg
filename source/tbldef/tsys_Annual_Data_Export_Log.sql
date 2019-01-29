CREATE TABLE [tsys_Annual_Data_Export_Log] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [DataTable] VARCHAR (255),
  [ZipFile] VARCHAR (255),
  [StartYear] SHORT ,
  [EndYear] SHORT ,
  [DataStoreID] LONG 
)
