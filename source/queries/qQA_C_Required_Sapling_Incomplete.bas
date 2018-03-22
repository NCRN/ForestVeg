Operation =1
Option =0
Where ="(((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Tags.TSN) Is Null) AND ((tbl_Sa"
    "pling_Data.Sapling_Status) Is Null)) OR (((tbl_Events.Event_Date) Is Not Null) A"
    "ND ((tbl_Tags.Tag) Is Null) AND ((tbl_Sapling_Data.Sapling_Status) Is Null)) OR "
    "(((tbl_Tags.Microplot_Number) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tbl_Sapling_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Tags.Microplot_Number"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Events.Event_Date"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbText "Description" ="Sapling sampling record is incomplete"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1440"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =15
    Top =351
    Right =1957
    Bottom =1015
    Left =-1
    Top =-1
    Right =1910
    Bottom =202
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =561
        Top =8
        Right =768
        Bottom =122
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =7
        Top =8
        Right =103
        Bottom =122
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =326
        Top =-5
        Right =470
        Bottom =271
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =142
        Top =33
        Right =286
        Bottom =177
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =545
        Top =133
        Right =689
        Bottom =277
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
