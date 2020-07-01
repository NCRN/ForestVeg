CREATE TABLE [tbl_History] (
  [History_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Table_Name] VARCHAR (64),
  [Record_ID_Field_Name] VARCHAR (64),
  [Record_ID] VARCHAR (50),
  [Field_Name] VARCHAR (64),
  [Value_New] VARCHAR (255),
  [Value_Old] VARCHAR (255),
  [Value_History_Notes] LONGTEXT ,
  [Contact_ID] VARCHAR (50),
  [Network_User_Name] VARCHAR (32),
  [Change_Date] DATETIME ,
  [Updated_Date] DATETIME 
)
