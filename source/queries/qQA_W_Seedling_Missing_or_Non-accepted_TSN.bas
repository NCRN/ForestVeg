Operation =1
Option =0
Where ="(((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Quadrat_Seedlings_Data.TSN) Is "
    "Null)) OR (((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Quadrat_Seedlings_Dat"
    "a.TSN)=0)) OR (((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Quadrat_Seedlings"
    "_Data.TSN)=999999)) OR (((tbl_Events.Event_Date) Is Not Null) AND ((tlu_Plants.A"
    "ccepted_Found)=True))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tbl_Quadrat_Seedlings_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Quadrat_Seedlings_Data.TSN"
    Expression ="tbl_Locations.Panel"
    Expression ="tlu_Plants.Accepted_Found"
End
Begin Joins
    LeftTable ="tbl_Quadrat_Seedlings_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_Quadrat_Seedlings_Data.TSN = tlu_Plants.TSN"
    Flag =2
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
dbText "Description" ="Seedling record is missing or uses non-accepted TSN"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.TSN"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Accepted_Found"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-124
    Top =57
    Right =961
    Bottom =660
    Left =-1
    Top =-1
    Right =1053
    Bottom =267
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =880
        Top =38
        Right =1016
        Bottom =152
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =441
        Top =29
        Right =699
        Bottom =143
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =218
        Top =43
        Right =362
        Bottom =187
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =43
        Top =141
        Right =187
        Bottom =285
        Top =0
        Name ="tbl_Quadrat_Seedlings_Data"
        Name =""
    End
    Begin
        Left =408
        Top =168
        Right =662
        Bottom =312
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
