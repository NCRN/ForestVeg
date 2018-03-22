Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qCalc_Exotic_Herb_Count_By_Event"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Events.Event_Date"
    Alias ="Exotic_Count"
    Expression ="Int(Nz([Count_of_Specimens],0))"
    Alias ="Presence"
    Expression ="IIf(IsNull([count_of_specimens]),\"Absent\",\"Present\")"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qCalc_Exotic_Herb_Count_By_Event"
    Expression ="qFiltered_Events.Event_ID = qCalc_Exotic_Herb_Count_By_Event.Event_ID"
    Flag =2
End
Begin OrderBy
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Flag =0
    Expression ="qFiltered_Locations.Plot_Name"
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
Begin
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Presence"
        dbInteger "ColumnWidth" ="1215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic_Count"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =525
    Right =1448
    Bottom =1194
    Left =-1
    Top =-1
    Right =1364
    Bottom =587
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =432
        Top =12
        Right =683
        Bottom =164
        Top =0
        Name ="qCalc_Exotic_Herb_Count_By_Event"
        Name =""
    End
End
