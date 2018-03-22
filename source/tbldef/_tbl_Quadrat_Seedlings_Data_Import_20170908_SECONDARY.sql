CREATE TABLE [_tbl_Quadrat_Seedlings_Data_Import_20170908_SECONDARY] (
  [Quadrat_Seedlings_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Quadrat_Data_ID] VARCHAR (50),
  [TSN] LONG ,
  [Height] SINGLE ,
  [Updated_Date] DATETIME ,
  [Browsable] VARCHAR (4),
  [Browsed] VARCHAR (4)
)
