CREATE TABLE [tsys_App_Releases] (
  [Release_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Release_date] DATETIME ,
  [Database_title] VARCHAR (100),
  [Version_number] VARCHAR (20),
  [File_name] VARCHAR (50),
  [Release_by] VARCHAR (50),
  [Release_notes] LONGTEXT ,
  [Author_phone] VARCHAR (50),
  [Author_email] VARCHAR (50),
  [Author_org] VARCHAR (10),
  [Author_org_name] VARCHAR (100),
   CONSTRAINT 
)
