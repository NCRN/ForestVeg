CREATE TABLE [tsys_Application_Changes] (
  [Change_timestamp] DATETIME ,
  [Application_component] VARCHAR (50),
  [Object_name] VARCHAR (100),
  [Procedure_name] VARCHAR (50),
  [Change_type] VARCHAR (10),
  [Change_description] LONGTEXT 
)
