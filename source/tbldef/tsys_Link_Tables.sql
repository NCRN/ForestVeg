CREATE TABLE [tsys_Link_Tables] (
  [Link_type] VARCHAR (50) CONSTRAINT [tsys_Link_Filestsys_Link_Tables] REFERENCES [tsys_Link_Files] ([Link_type]) ON UPDATE CASCADE ,
  [Link_table] VARCHAR (100),
  [Table_type] VARCHAR (50),
  [Description_text] VARCHAR (255),
  [Is_hidden] BIT ,
  [Allow_edits_lookup] BIT ,
  [Browser_view] BIT ,
  [Sort_order] BYTE ,
  [AnnualDbPurge] BYTE ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Link_type], [Link_table])
)
