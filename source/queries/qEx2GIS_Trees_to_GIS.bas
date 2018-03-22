Operation =1
Option =0
Where ="(((tbl_Tags.Tag_Status)=\"Tree\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="X_UTM"
    Expression ="[X_Coord]+([Distance]*(Sin([Azimuth]*3.1415/180)))"
    Alias ="Y_UTM"
    Expression ="[Y_Coord]+([Distance]*(Cos([Azimuth]*3.1415/180)))"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tags.Tag_Notes"
    Expression ="tbl_Tags.Start_Date"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
    Flag =0
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
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
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
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X_UTM"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Y_UTM"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =34
    Top =7
    Right =1325
    Bottom =670
    Left =-1
    Top =-1
    Right =1259
    Bottom =297
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =20
        Top =23
        Right =228
        Bottom =273
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =294
        Top =19
        Right =475
        Bottom =291
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =545
        Top =18
        Right =746
        Bottom =291
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
