CREATE TABLE [tbl_master_version] (
  [project_ID] LONG ,
  [version_key_number] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [version_key_date] DATETIME ,
  [narrative_version] DECIMAL (18, 2),
  [version_comments] LONGTEXT 
)
