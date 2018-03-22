Operation =1
Option =0
Begin InputTables
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="qActive_Shrub_Data"
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Count_of_Shrubs"
    Expression ="Count(qActive_Shrub_Data.Sapling_Data_ID)"
    Alias ="Count_of_Exotic_Shrubs"
    Expression ="Sum(Abs([Exotic]))"
    Alias ="Percent_Shrubs_Exotic"
    Expression ="Round([Count_of_Exotic_Shrubs]*100/[Count_of_Shrubs],1)"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="qActive_Shrub_Data"
    Expression ="tbl_Tags.Tag_ID = qActive_Shrub_Data.Tag_ID"
    Flag =1
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qActive_Shrub_Data"
    RightTable ="qFiltered_Events"
    Expression ="qActive_Shrub_Data.Event_ID = qFiltered_Events.Event_ID"
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
        dbText "Name" ="Percent_Shrubs_Exotic"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Exotic_Shrubs"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Shrubs"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =232
    Top =84
    Right =954
    Bottom =646
    Left =-1
    Top =-1
    Right =690
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =746
        Top =11
        Right =890
        Bottom =268
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =923
        Top =7
        Right =1067
        Bottom =448
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =551
        Top =13
        Right =695
        Bottom =157
        Top =0
        Name ="qActive_Shrub_Data"
        Name =""
    End
    Begin
        Left =203
        Top =2
        Right =347
        Bottom =147
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =377
        Top =7
        Right =521
        Bottom =151
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
End
