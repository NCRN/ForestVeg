CREATE TABLE [tsys_Bug_Reports] (
  [Bug_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Release_ID] VARCHAR (50) CONSTRAINT [tsys_App_Releasestsys_Bug_Reports] REFERENCES [tsys_App_Releases] ([Release_ID]) ON UPDATE CASCADE ,
  [Report_date] DATETIME ,
  [Found_by] VARCHAR (50),
  [Reported_by] VARCHAR (50),
  [Report_details] LONGTEXT ,
  [Fix_date] DATETIME ,
  [Fixed_by] VARCHAR (50),
  [Fix_details] LONGTEXT 
)
