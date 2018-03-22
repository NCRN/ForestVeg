Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Sapling_Data.Sapling_Notes"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="tbl_Tree_Data.Tree_Notes"
End
Begin Joins
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Sapling_Data.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Sapling_Data.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
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
    Expression ="tbl_Events.Event_Date"
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
dbText "Description" ="Both a tree and sapling record exist for a tag within a single event"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1290"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =43
    Top =497
    Right =1137
    Bottom =926
    Left =-1
    Top =-1
    Right =830
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =460
        Top =20
        Right =604
        Bottom =251
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =664
        Top =17
        Right =808
        Bottom =247
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =894
        Top =22
        Right =1038
        Bottom =261
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
