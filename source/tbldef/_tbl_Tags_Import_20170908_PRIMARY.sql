CREATE TABLE [_tbl_Tags_Import_20170908_PRIMARY] (
  [Tag_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Location_ID] VARCHAR (50),
  [Tag] LONG  CONSTRAINT [Tree_Tag] UNIQUE ,
  [Azimuth] LONG ,
  [Distance] DOUBLE ,
  [Microplot_Number] LONG ,
  [TSN] LONG ,
  [Tag_Notes] LONGTEXT ,
  [Start_Date] DATETIME ,
  [Stop_Date] DATETIME ,
  [Tag_Status] VARCHAR (32),
  [RFS] BIT ,
  [Updated_Date] DATETIME ,
  [TaxonCode] LONG 
)
