Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qActive_Vine_Data"
    Name ="tbl_Tags"
    Name ="tbl_Tree_Data"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Count_of_Vines"
    Expression ="Count(tbl_Tree_Data.Tree_Data_ID)"
    Alias ="Count_of_Exotic_Vines"
    Expression ="Sum(Abs([Exotic]))"
    Alias ="Percent_Vines_Exotic"
    Expression ="Round([Count_of_Exotic_Vines]*100/[Count_of_Vines],1)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qActive_Vine_Data"
    RightTable ="tbl_Tags"
    Expression ="qActive_Vine_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="qActive_Vine_Data"
    RightTable ="tbl_Tree_Data"
    Expression ="qActive_Vine_Data.Tree_Data_ID = tbl_Tree_Data.Tree_Data_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="qFiltered_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
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
        dbText "Name" ="Count_of_Vines"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Exotic_Vines"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_Vines_Exotic"
        dbInteger "ColumnWidth" ="2010"
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
        Left =14
        Top =21
        Right =158
        Bottom =165
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =194
        Top =22
        Right =338
        Bottom =166
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =534
        Top =28
        Right =678
        Bottom =199
        Top =0
        Name ="qActive_Vine_Data"
        Name =""
    End
    Begin
        Left =707
        Top =33
        Right =851
        Bottom =235
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =368
        Top =30
        Right =512
        Bottom =174
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
End
