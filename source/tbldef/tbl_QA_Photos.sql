CREATE TABLE [tbl_QA_Photos] (
  [QA_Photo_ID] VARCHAR (50) CONSTRAINT [Db_Meta_ID] UNIQUE  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Location_ID] VARCHAR (50),
  [Contact_ID] VARCHAR (50),
  [AD_Name] VARCHAR (64),
  [Error_Date] DATETIME ,
  [Event_Date1] DATETIME ,
  [Event_Date2] DATETIME ,
  [Error_Detected] BIT ,
  [Error_Description] LONGTEXT ,
  [Remedy_Description] LONGTEXT ,
  [Remedy_date] DATETIME ,
  [Error_Addressed] BIT 
)
