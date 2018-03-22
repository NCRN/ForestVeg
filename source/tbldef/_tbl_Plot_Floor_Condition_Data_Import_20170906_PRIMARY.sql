CREATE TABLE [_tbl_Plot_Floor_Condition_Data_Import_20170906_PRIMARY] (
  [Plot_Floor_Data_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID] VARCHAR (50),
  [Rock_Cover] VARCHAR (16),
  [Bare_Soil_Cover] VARCHAR (16),
  [Trampled] VARCHAR (16),
  [Humus] BIT ,
  [Earthworms] BIT ,
  [Updated_Date] DATETIME 
)
