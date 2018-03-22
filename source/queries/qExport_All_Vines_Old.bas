Operation =1
Option =0
Begin InputTables
    Name ="tlu_Plants"
    Name ="tbl_Locations"
    Name ="tbl_Tags"
    Name ="qActive_Vine_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="qActive_Vine_Data.Unit_Group"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="qActive_Vine_Data.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Host_Tree_Tag"
    Expression ="tbl_Tags.Tag"
    Expression ="qActive_Vine_Data.Host_TSN"
    Expression ="qActive_Vine_Data.Host_Latin_Name"
    Expression ="qActive_Vine_Data.Host_Status"
    Expression ="qActive_Vine_Data.Condition"
    Expression ="tlu_Plants.Exotic"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="qActive_Vine_Data"
    Expression ="tbl_Tags.Tag_ID = qActive_Vine_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="qActive_Vine_Data"
    Expression ="tbl_Locations.Location_ID = qActive_Vine_Data.Location_ID"
    Flag =1
    LeftTable ="tlu_Plants"
    RightTable ="qActive_Vine_Data"
    Expression ="tlu_Plants.TSN = qActive_Vine_Data.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="qActive_Vine_Data.Sample_Year"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Tree_Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Host_TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Host_Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Host_Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =89
    Top =127
    Right =1175
    Bottom =841
    Left =-1
    Top =-1
    Right =1054
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =3
        Top =5
        Right =147
        Bottom =225
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =615
        Top =179
        Right =849
        Bottom =407
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =614
        Top =9
        Right =844
        Bottom =153
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =195
        Top =6
        Right =378
        Bottom =397
        Top =0
        Name ="qActive_Vine_Data"
        Name =""
    End
End
