Operation =1
Option =0
Where ="(((tlu_Enumerations.Enum_Group)=\"Foliage Condition\"))"
Begin InputTables
    Name ="tbl_Tree_Foliage_Conditions"
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="qActive_Tree_Data"
    Name ="tlu_Enumerations"
End
Begin OutputColumns
    Expression ="qActive_Tree_Data.Plot_Name"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="qActive_Tree_Data.Sample_Year"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tree_Foliage_Conditions.Condition"
    Alias ="Condition_Description"
    Expression ="tlu_Enumerations.Enum_Description"
    Expression ="tbl_Tree_Foliage_Conditions.Percent_Afflicted"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qActive_Tree_Data.Crown_Class"
    Expression ="qActive_Tree_Data.Tree_Status"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Tags.Tag_ID"
    Expression ="qActive_Tree_Data.Tree_Data_ID"
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
    LeftTable ="tbl_Tree_Foliage_Conditions"
    RightTable ="qActive_Tree_Data"
    Expression ="tbl_Tree_Foliage_Conditions.Tree_Data_ID = qActive_Tree_Data.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tree_Foliage_Conditions"
    RightTable ="tlu_Enumerations"
    Expression ="tbl_Tree_Foliage_Conditions.Condition = tlu_Enumerations.Enum_Code"
    Flag =1
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
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Crown_Class"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
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
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =22
    Top =53
    Right =1513
    Bottom =960
    Left =-1
    Top =-1
    Right =1459
    Bottom =509
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
        Left =295
        Top =27
        Right =439
        Bottom =302
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
    Begin
        Left =157
        Top =312
        Right =301
        Bottom =456
        Top =0
        Name ="tlu_Enumerations"
        Name =""
    End
End
