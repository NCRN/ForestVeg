Operation =1
Option =0
Where ="(((tbl_Sapling_Data.Habit)=\"Shrub\") AND ((tlu_Plants.Shrub)=False)) OR (((tbl_"
    "Sapling_Data.Habit)=\"Tree\") AND ((tlu_Plants.Shrub)=True))"
Begin InputTables
    Name ="tbl_Sapling_Data"
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.Tag"
    Expression ="tlu_Plants.Genus"
    Expression ="tlu_Plants.Species"
    Expression ="tbl_Sapling_Data.Habit"
    Expression ="tlu_Plants.Shrub"
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
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Shrub"
        dbInteger "ColumnWidth" ="1080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =555
    Right =1040
    Bottom =1015
    Left =-1
    Top =-1
    Right =996
    Bottom =201
    Left =0
    Top =192
    ColumnsShown =539
    Begin
        Left =520
        Top =-19
        Right =766
        Bottom =360
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =29
        Top =-174
        Right =191
        Bottom =228
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =259
        Top =-121
        Right =404
        Bottom =89
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =856
        Top =94
        Right =1037
        Bottom =390
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =1195
        Top =-174
        Right =1408
        Bottom =410
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
