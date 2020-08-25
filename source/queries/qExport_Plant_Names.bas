dbMemo "SQL" ="SELECT p.Latin_Name, p.Common AS NCRN_Common, p.NPSpecies_Common, p.Family, p.Ge"
    "nus, p.Species, p.TSN, p.Favorite, p.Woody, p.Herbaceous, p.Targeted_Herb, p.Tre"
    "e, p.Shrub, p.Vine, p.Exotic\015\012FROM tlu_Plants AS p\015\012ORDER BY p.Latin"
    "_Name;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Family"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Genus"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Species"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbInteger "ColumnWidth" ="9030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Tree"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Vine"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NCRN_Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Favorite"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Woody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Herbaceous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Targeted_Herb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.NPSpecies_Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Favorite"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Woody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Herbaceous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Targeted_Herb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Tree"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Vine"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Exotic"
        dbLong "AggregateType" ="-1"
    End
End
