CREATE TABLE [tsys_Import_Tables] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Table_Name] VARCHAR (50),
  [Import] BIT 
)
