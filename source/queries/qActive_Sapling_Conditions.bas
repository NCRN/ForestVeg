Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tlu_Tree_Condition"
    Name ="tbl_Sapling_Conditions"
    Name ="tbl_Sapling_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Unit_Group"
    Expression ="tbl_Locations.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Sample_Year"
    Expression ="Year([Event_Date])"
    Expression ="tbl_Events.Event_Date"
    Alias ="Date"
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Status"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Sapling_Conditions.Condition"
    Expression ="tlu_Tree_Condition.Pest"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Tags.Tag_ID"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID"
    Expression ="tbl_Sapling_Conditions.Sapling_Condition_ID"
    Expression ="tbl_Locations.Admin_Unit_Code"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Sapling_Conditions"
    RightTable ="tlu_Tree_Condition"
    Expression ="tbl_Sapling_Conditions.Condition = tlu_Tree_Condition.Description"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Sapling_Conditions"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_Conditions.Sapling_Data_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
    Flag =0
    Expression ="tlu_Plants.Latin_Name"
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
        dbText "Name" ="tlu_Tree_Condition.Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
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
        dbText "Name" ="tbl_Tags.TSN"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Conditions.Condition"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Conditions.Sapling_Condition_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =80
    Top =4
    Right =1554
    Bottom =669
    Left =-1
    Top =-1
    Right =1442
    Bottom =367
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
        Left =479
        Top =7
        Right =623
        Bottom =151
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =461
        Top =160
        Right =609
        Bottom =424
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =971
        Top =95
        Right =1115
        Bottom =520
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =230
        Top =226
        Right =390
        Bottom =370
        Top =0
        Name ="tlu_Tree_Condition"
        Name =""
    End
    Begin
        Left =9
        Top =14
        Right =197
        Bottom =172
        Top =0
        Name ="tbl_Sapling_Conditions"
        Name =""
    End
    Begin
        Left =253
        Top =35
        Right =397
        Bottom =179
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
End
