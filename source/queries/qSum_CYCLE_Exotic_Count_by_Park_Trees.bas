Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qActive_Tree_Data"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Count_of_Trees"
    Expression ="Count(qActive_Tree_Data.Tree_Data_ID)"
    Alias ="Count_of_Exotic_Trees"
    Expression ="Sum(Abs([Exotic]))"
    Alias ="Percent_Trees_Exotic"
    Expression ="Round([Count_of_Exotic_Trees]*100/[Count_of_Trees],1)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qActive_Tree_Data"
    Expression ="qFiltered_Events.Event_ID = qActive_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="qActive_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Flag =0
End
Begin Groups
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    GroupLevel =0
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
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Trees"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Exotic_Trees"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_Trees_Exotic"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =-1
    Top =-1
    Right =690
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =434
        Top =15
        Right =578
        Bottom =159
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =269
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =453
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
