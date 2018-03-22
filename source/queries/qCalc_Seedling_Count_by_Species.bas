Operation =1
Option =0
Having ="(((tlu_Plants.Shrub)=False))"
Begin InputTables
    Name ="tlu_Plants"
    Name ="tbl_Quadrat_Seedlings_Data"
    Name ="tbl_Quadrat_Data"
    Name ="qFiltered_Events"
    Name ="qFiltered_Locations"
End
Begin OutputColumns
    Expression ="tlu_Plants.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.PLANTS_Common"
    Alias ="Seedling_Count"
    Expression ="Count(tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID)"
End
Begin Joins
    LeftTable ="tbl_Quadrat_Data"
    RightTable ="qFiltered_Events"
    Expression ="tbl_Quadrat_Data.Event_ID = qFiltered_Events.Event_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qFiltered_Locations"
    Expression ="qFiltered_Events.Location_ID = qFiltered_Locations.Location_ID"
    Flag =1
    LeftTable ="tlu_Plants"
    RightTable ="tbl_Quadrat_Seedlings_Data"
    Expression ="tlu_Plants.TSN = tbl_Quadrat_Seedlings_Data.TSN"
    Flag =1
    LeftTable ="tbl_Quadrat_Data"
    RightTable ="tbl_Quadrat_Seedlings_Data"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID = tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tlu_Plants.Latin_Name"
    Flag =0
End
Begin Groups
    Expression ="tlu_Plants.TSN"
    GroupLevel =0
    Expression ="tlu_Plants.Latin_Name"
    GroupLevel =0
    Expression ="tlu_Plants.PLANTS_Common"
    GroupLevel =0
    Expression ="tlu_Plants.Shrub"
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
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.PLANTS_Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Seedling_Count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =94
    Right =1177
    Bottom =766
    Left =-1
    Top =-1
    Right =1117
    Bottom =336
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =317
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =401
        Bottom =242
        Top =0
        Name ="tbl_Quadrat_Seedlings_Data"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =262
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
End
