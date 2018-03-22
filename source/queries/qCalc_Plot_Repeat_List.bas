Operation =1
Option =0
Having ="(((Count(qFiltered_Events.Event_ID))>1))"
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Location_ID"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Locations.Panel"
    Alias ="Event_Count"
    Expression ="Count(qFiltered_Events.Event_ID)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
End
Begin Groups
    Expression ="qFiltered_Locations.Location_ID"
    GroupLevel =0
    Expression ="qFiltered_Locations.Plot_Name"
    GroupLevel =0
    Expression ="qFiltered_Locations.Panel"
    GroupLevel =0
End
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
        dbText "Name" ="qFiltered_Locations.Location_ID"
        dbInteger "ColumnWidth" ="4200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Count"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =44
    Right =1279
    Bottom =923
    Left =-1
    Top =-1
    Right =1227
    Bottom =581
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =226
        Bottom =350
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =288
        Top =11
        Right =486
        Bottom =360
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
End
