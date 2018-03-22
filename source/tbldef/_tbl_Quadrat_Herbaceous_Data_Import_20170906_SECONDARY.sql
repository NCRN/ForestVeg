CREATE TABLE [_tbl_Quadrat_Herbaceous_Data_Import_20170906_SECONDARY] (
  [Quadrat_Herbaceous_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Quadrat_Data_ID] VARCHAR (50),
  [TSN] LONG ,
  [Percent_Cover] LONG ,
  [Updated_Date] DATETIME ,
  [Browse] VARCHAR (255)
)
