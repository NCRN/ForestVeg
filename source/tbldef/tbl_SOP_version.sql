CREATE TABLE [tbl_SOP_version] (
  [version_key_number] LONG ,
  [SOP_number] LONG ,
  [SOP_version_number] DECIMAL (18, 2),
  [active_flag] BIT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([version_key_number], [SOP_number])
)
