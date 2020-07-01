CREATE TABLE [ParkTaxonomicReference] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ParkCode] VARCHAR (6),
  [TaxonomicReferenceID] LONG ,
  [IsActive] BIT 
)
