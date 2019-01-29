CREATE TABLE [tsys_Annual_Data_Tables] (
  [AnnualData] VARCHAR (255),
  [RelatedQuery] VARCHAR (255),
  [Sequence] SHORT ,
   CONSTRAINT [DataQuery] PRIMARY KEY ([AnnualData], [RelatedQuery])
)
