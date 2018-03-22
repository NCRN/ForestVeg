Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Alias ="Sample_Year"
    Expression ="Year([Event_Date])"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
End
Begin OrderBy
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Flag =0
    Expression ="tbl_Locations.Panel"
    Flag =0
    Expression ="Year([Event_Date])"
    Flag =0
End
Begin Groups
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    GroupLevel =0
    Expression ="tbl_Locations.Panel"
    GroupLevel =0
    Expression ="Year([Event_Date])"
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
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =87
    Top =169
    Right =815
    Bottom =831
    Left =-1
    Top =-1
    Right =696
    Bottom =399
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =458
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
