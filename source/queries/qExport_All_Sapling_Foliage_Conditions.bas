Operation =1
Option =0
Where ="(((tlu_Enumerations.Enum_Group)=\"Foliage Condition\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tlu_Enumerations"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Sapling_Foliage_Conditions"
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
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Sapling_Foliage_Conditions.Condition"
    Alias ="Condition_Description"
    Expression ="tlu_Enumerations.Enum_Description"
    Expression ="tbl_Sapling_Foliage_Conditions.Percent_Afflicted"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Status"
    Expression ="tbl_Sapling_Data.Sapling_Status"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Sapling_Foliage_Conditions"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_Foliage_Conditions.Sapling_Data_I"
        "D"
    Flag =1
    LeftTable ="tbl_Sapling_Foliage_Conditions"
    RightTable ="tlu_Enumerations"
    Expression ="tbl_Sapling_Foliage_Conditions.Condition = tlu_Enumerations.Enum_Code"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbInteger "ColumnWidth" ="780"
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
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Condition_Description"
        dbInteger "ColumnWidth" ="2250"
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
        dbText "Name" ="tbl_Sapling_Foliage_Conditions.Percent_Afflicted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Foliage_Conditions.Condition"
        dbInteger "ColumnWidth" ="1215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =63
    Top =37
    Right =1366
    Bottom =791
    Left =-1
    Top =-1
    Right =1271
    Bottom =485
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =763
        Top =17
        Right =907
        Bottom =229
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =576
        Top =14
        Right =720
        Bottom =158
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =577
        Top =181
        Right =758
        Bottom =486
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =1019
        Top =18
        Right =1163
        Bottom =440
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =-6
        Top =221
        Right =138
        Bottom =365
        Top =0
        Name ="tlu_Enumerations"
        Name =""
    End
    Begin
        Left =381
        Top =196
        Right =525
        Bottom =340
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =161
        Top =88
        Right =305
        Bottom =232
        Top =0
        Name ="tbl_Sapling_Foliage_Conditions"
        Name =""
    End
End
