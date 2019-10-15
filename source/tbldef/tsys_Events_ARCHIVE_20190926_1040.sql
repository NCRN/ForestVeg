CREATE TABLE [tsys_Events_ARCHIVE_20190926_1040] (
  [Event_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Location_ID] VARCHAR (50),
  [Event_Group_ID] VARCHAR (50),
  [Protocol_Name] VARCHAR (100),
  [PseudoEvent] BYTE ,
  [Event_Date] DATETIME ,
  [Event_Time] DATETIME ,
  [Event_Notes] LONGTEXT ,
  [Pictures_Taken] BIT ,
  [CWD_Check_360] BIT ,
  [CWD_Check_120] BIT ,
  [CWD_Check_240] BIT ,
  [Deer_Impact] LONG ,
  [Is_Excluded] BIT ,
  [Early_Detect] BIT ,
  [Rare_Spp] BIT ,
  [Plot_Maint] BIT ,
  [Entered_On_Tablet] BIT ,
  [Entered_By] VARCHAR (50),
  [Entered_Date] DATETIME ,
  [Updated_By] VARCHAR (50),
  [Updated_Date] DATETIME ,
  [Verified] BIT ,
  [Verified_By] VARCHAR (50),
  [Verified_Date] DATETIME ,
  [Certified] BIT ,
  [Certified_By] VARCHAR (50),
  [Certified_Date] DATETIME 
)
