Operation =1
Option =0
Where ="(((tlu_Plants.Shrub)=False))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tlu_Plants"
    Name ="tbl_Quadrat_Seedlings_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Quadrat_Seedlings_Data.*"
    Expression ="tlu_Plants.Shrub"
    Alias ="Sample_Year"
    Expression ="Year([tbl_Events].[Event_Date])"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tlu_Plants.Exotic"
    Expression ="tbl_Locations.Frame"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="tbl_Quadrat_Seedlings_Data"
    Expression ="tlu_Plants.TSN = tbl_Quadrat_Seedlings_Data.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="tbl_Events.Event_ID = tbl_Quadrat_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Quadrat_Data"
    RightTable ="tbl_Quadrat_Seedlings_Data"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID = tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID"
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
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =94
    Right =1337
    Bottom =966
    Left =-1
    Top =-1
    Right =1277
    Bottom =491
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =413
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =772
        Bottom =221
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =434
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Quadrat_Seedlings_Data"
        Name =""
    End
End
