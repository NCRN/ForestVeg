CREATE TABLE [tbl_wrk_Species_List] (
  [Table_Index] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Plot_ID] SHORT ,
  [Plant_Code] VARCHAR (50),
  [P1] BIT ,
  [P2] BIT ,
  [P3] BIT ,
  [P4] BIT ,
  [P5] BIT ,
  [P6] BIT ,
  [P7] BIT ,
  [P8] BIT ,
  [P9] BIT ,
  [P10] BIT ,
  [Park_Code] VARCHAR (4)
)
