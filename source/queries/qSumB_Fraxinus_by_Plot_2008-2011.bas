Operation =1
Option =0
Having ="(((tbl_Events.Event_Date)>=#1/1/2008# And (tbl_Events.Event_Date)<=#12/31/2011#)"
    ")"
Begin InputTables
    Name ="tbl_Events"
    Name ="tbl_Locations"
    Name ="qCalc_Fraxinus_Count_by_Event"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Alias ="FraxCount"
    Expression ="Nz([Occurences],0)"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="qCalc_Fraxinus_Count_by_Event"
    Expression ="tbl_Events.Event_ID = qCalc_Fraxinus_Count_by_Event.Event_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =0
    Expression ="tbl_Events.Event_Date"
    GroupLevel =0
    Expression ="Nz([Occurences],0)"
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
dbText "Description" ="How many total fraxinus trees, saplings and seedlings were found in each plot fr"
    "om 2008-2011?"
Begin
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FraxCount"
        dbInteger "ColumnWidth" ="1260"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =26
    Top =44
    Right =871
    Bottom =795
    Left =-1
    Top =-1
    Right =1364
    Bottom =517
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =229
        Top =16
        Right =381
        Bottom =337
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =13
        Top =24
        Right =157
        Bottom =168
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =537
        Top =27
        Right =681
        Bottom =171
        Top =0
        Name ="qCalc_Fraxinus_Count_by_Event"
        Name =""
    End
End
