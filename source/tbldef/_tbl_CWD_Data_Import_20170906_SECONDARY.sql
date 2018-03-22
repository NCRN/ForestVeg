CREATE TABLE [_tbl_CWD_Data_Import_20170906_SECONDARY] (
  [CWD_Data_ID] VARCHAR (50) CONSTRAINT [CWD_Data_txt] UNIQUE  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID] VARCHAR (50),
  [Transect_Azimuth] VARCHAR (8),
  [Decay_Class] VARCHAR (16),
  [TSN] LONG ,
  [Diameter] SINGLE ,
  [Hollow] BIT ,
  [CWD_Notes] VARCHAR (255),
  [Tag_ID] VARCHAR (50),
  [Updated_Date] DATETIME 
)
