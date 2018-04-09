CREATE TABLE [Copy Of tbl_Sapling_Data] (
  [Sapling_Data_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID] VARCHAR (50),
  [Tag_ID] VARCHAR (50),
  [Status] VARCHAR (32),
  [Sapling_Status] VARCHAR (32),
  [Habit] VARCHAR (16),
  [Browsable] VARCHAR (4),
  [Browsed] VARCHAR (4),
  [Sapling_Notes] LONGTEXT ,
  [DRC] SINGLE ,
  [Vines_Checked] BIT ,
  [Conditions_Checked] BIT ,
  [Foliage_Conditions_Checked] BIT ,
  [Updated_Date] DATETIME ,
  [SaplingVigor] SHORT 
)
