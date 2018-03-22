Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qActive_Seedling_Data"
    Name ="tbl_Quadrat_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Count_of_Seedlings"
    Expression ="Count(qActive_Seedling_Data.Quadrat_Seedlings_ID)"
    Alias ="Count_of_Exotic_Seedlings"
    Expression ="Sum(Abs([Exotic]))"
    Alias ="Percent_Seedlings_Exotic"
    Expression ="Round([Count_of_Exotic_Seedlings]*100/[Count_of_Seedlings],1)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qActive_Seedling_Data"
    RightTable ="tbl_Quadrat_Data"
    Expression ="qActive_Seedling_Data.Quadrat_Data_ID = tbl_Quadrat_Data.Quadrat_Data_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="qFiltered_Events.Event_ID = tbl_Quadrat_Data.Event_ID"
    Flag =1
    LeftTable ="qActive_Seedling_Data"
    RightTable ="tlu_Plants"
    Expression ="qActive_Seedling_Data.TSN = tlu_Plants.TSN"
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
        dbText "Name" ="Count_of_Seedlings"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Exotic_Seedlings"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_Seedlings_Exotic"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =0
    Top =0
    Right =0
    Bottom =0
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
        Left =654
        Top =18
        Right =836
        Bottom =318
        Top =0
        Name ="qActive_Seedling_Data"
        Name =""
    End
    Begin
        Left =455
        Top =14
        Right =599
        Bottom =169
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =884
        Top =12
        Right =1028
        Bottom =427
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
