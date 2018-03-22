Operation =1
Option =0
Where ="(((tbl_Tags.Tag_Status)=\"Tree\"))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="tbl_Tags.Location_ID"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tags.Tag"
    Alias ="X"
    Expression ="([Distance]*(Sin([Azimuth]*3.1415/180)))"
    Alias ="Y"
    Expression ="([Distance]*(Cos([Azimuth]*3.1415/180)))"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Y"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =80
    Top =85
    Right =1550
    Bottom =989
    Left =-1
    Top =-1
    Right =1438
    Bottom =536
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =217
        Top =23
        Right =395
        Bottom =314
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =13
        Top =12
        Right =188
        Bottom =489
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
