Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Tree_Conditions"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="qActive_Tree_Data"
    Name ="tlu_Tree_Condition"
End
Begin OutputColumns
    Expression ="qActive_Tree_Data.Plot_Name"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="qActive_Tree_Data.Sample_Year"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qActive_Tree_Data.Crown_Class"
    Expression ="qActive_Tree_Data.Tree_Status"
    Expression ="tbl_Tree_Conditions.Condition"
    Expression ="tlu_Tree_Condition.Pest"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Tags.Tag_ID"
    Expression ="qActive_Tree_Data.Tree_Data_ID"
    Expression ="tbl_Tree_Conditions.Tree_Condition_ID"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="qActive_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = qActive_Tree_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="qActive_Tree_Data"
    Expression ="tbl_Events.Event_ID = qActive_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Conditions"
    RightTable ="qActive_Tree_Data"
    Expression ="tbl_Tree_Conditions.Tree_Data_ID = qActive_Tree_Data.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tree_Conditions"
    RightTable ="tlu_Tree_Condition"
    Expression ="tbl_Tree_Conditions.Condition = tlu_Tree_Condition.Description"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qActive_Tree_Data.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
    Flag =0
    Expression ="tlu_Plants.Latin_Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
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
        dbText "Name" ="qActive_Tree_Data.Plot_Name"
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
        dbText "Name" ="qActive_Tree_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Crown_Class"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Conditions.Tree_Condition_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Conditions.Condition"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
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
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =24
    Top =83
    Right =1498
    Bottom =748
    Left =-1
    Top =-1
    Right =1442
    Bottom =356
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
        Left =42
        Top =12
        Right =186
        Bottom =156
        Top =0
        Name ="tbl_Tree_Conditions"
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
        Right =725
        Bottom =445
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
        Left =265
        Top =11
        Right =423
        Bottom =274
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
    Begin
        Left =265
        Top =290
        Right =425
        Bottom =434
        Top =0
        Name ="tlu_Tree_Condition"
        Name =""
    End
End
