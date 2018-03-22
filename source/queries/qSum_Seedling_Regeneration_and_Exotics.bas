Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qActive_Seedling_Data"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Events.Event_Year"
    Alias ="All_Seedlings"
    Expression ="Sum(IIf(IsNull([qActive_Seedling_Data]![Exotic]),0,1))"
    Alias ="Native_Seedlings"
    Expression ="Sum(IIf([qActive_Seedling_Data]![Exotic]=False,1,0))"
    Alias ="Seedlings/ha"
    Expression ="Round([All_Seedlings]/0.0012,0)"
    Alias ="Native_Seedlings/ha"
    Expression ="Round([Native_Seedlings]/0.0012,0)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qActive_Seedling_Data"
    Expression ="qFiltered_Events.Event_ID = qActive_Seedling_Data.Event_ID"
    Flag =2
End
Begin OrderBy
    Expression ="qFiltered_Locations.Plot_Name"
    Flag =0
End
Begin Groups
    Expression ="qFiltered_Locations.Plot_Name"
    GroupLevel =0
    Expression ="qFiltered_Events.Event_Year"
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
dbText "Description" ="Reports the total number of saplings, the number of exotic saplings, and the den"
    "sity in each plot. Created for IAN NRCA reports."
Begin
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="All_Seedlings"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Native_Seedlings"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Seedlings/ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Native_Seedlings/ha"
        dbInteger "ColumnWidth" ="2130"
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
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =25
        Top =24
        Right =169
        Bottom =168
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =199
        Top =28
        Right =343
        Bottom =172
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =373
        Top =28
        Right =551
        Bottom =305
        Top =0
        Name ="qActive_Seedling_Data"
        Name =""
    End
End
