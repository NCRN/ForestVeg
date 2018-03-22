CREATE TABLE [_tbl_Quadrat_Data_Import_20170908_SECONDARY] (
  [Quadrat_Data_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID] VARCHAR (50),
  [Quadrat_Number] VARCHAR (16),
  [Browse] BIT ,
  [Percent_Trees] LONG ,
  [Percent_Rock] LONG ,
  [Percent_Woody_Debris] LONG ,
  [Percent_Fine_Woody_Debris] LONG ,
  [Percent_Other] LONG ,
  [Percent_Grasses] LONG ,
  [Percent_Sedges] LONG ,
  [Percent_Herbs] LONG ,
  [Percent_Ferns] LONG ,
  [Percent_Bryophytes] LONG ,
  [Quadrat_Notes] LONGTEXT ,
  [Updated_Date] DATETIME 
)
