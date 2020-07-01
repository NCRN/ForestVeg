CREATE TABLE [Measurement_Flags] (
  [RecordTable] VARCHAR (50),
  [Record_ID] LONG ,
  [RecordField] VARCHAR (50),
  [Flag_ID] LONG ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([RecordTable], [Record_ID], [RecordField], [Flag_ID])
)
