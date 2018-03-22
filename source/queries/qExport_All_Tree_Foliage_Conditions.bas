Operation =1
Option =0
Where ="(((tlu_Enumerations.Enum_Group)=\"Foliage Condition\"))"
Begin InputTables
    Name ="tbl_Tree_Foliage_Conditions"
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tlu_Enumerations"
    Name ="tbl_Tree_Data"
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
    Expression ="tbl_Tree_Foliage_Conditions.Condition"
    Alias ="Condition_Description"
    Expression ="tlu_Enumerations.Enum_Description"
    Expression ="tbl_Tree_Foliage_Conditions.Percent_Afflicted"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Status"
    Expression ="tbl_Tree_Data.Tree_Status"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Tree_Foliage_Conditions"
    RightTable ="tlu_Enumerations"
    Expression ="tbl_Tree_Foliage_Conditions.Condition = tlu_Enumerations.Enum_Code"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_Foliage_Conditions"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Foliage_Conditions.Tree_Data_ID"
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
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Foliage_Conditions.Condition"
        dbInteger "ColumnWidth" ="1215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Foliage_Conditions.Percent_Afflicted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Condition_Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="1425"
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
    Left =-41
    Top =109
    Right =1542
    Bottom =863
    Left =-1
    Top =-1
    Right =1551
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =95
        Top =26
        Right =241
        Bottom =184
        Top =0
        Name ="tbl_Tree_Foliage_Conditions"
        Name =""
    End
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
        Right =721
        Bottom =401
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
        Left =65
        Top =263
        Right =209
        Bottom =407
        Top =0
        Name ="tlu_Enumerations"
        Name =""
    End
    Begin
        Left =323
        Top =24
        Right =467
        Bottom =168
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
End
