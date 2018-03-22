Operation =1
Option =0
Where ="(((tbl_Sapling_Data.Sapling_Status)<>\"Removed from study\") AND ((tbl_Sapling_D"
    "ata.Habit)=\"Shrub\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tlu_Plants"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Unit_Group"
    Expression ="tbl_Locations.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Sample_Year"
    Expression ="Year([Event_Date])"
    Alias ="Date"
    Expression ="Format([tbl_Events].[Event_Date],\"yyyymmdd\")"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Status"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Sapling_Data.Habit"
    Expression ="tbl_Sapling_Data.Browsed"
    Expression ="tbl_Sapling_Data.Browsable"
End
Begin Joins
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =20
    Top =41
    Right =1539
    Bottom =633
    Left =-1
    Top =-1
    Right =1487
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =4
        Top =7
        Right =162
        Bottom =243
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =194
        Top =7
        Right =338
        Bottom =245
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =1034
        Top =-12
        Right =1206
        Bottom =237
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =410
        Top =67
        Right =567
        Bottom =234
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =681
        Top =86
        Right =825
        Bottom =230
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
