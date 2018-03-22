Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="qCalc_Exotic_Count_By_Event"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Admin_Unit_Code"
    Alias ="Sample_Year"
    Expression ="Year([Event_Date])"
    Alias ="Count_of_Exotic_Plots"
    Expression ="Count(qCalc_Exotic_Count_By_Event.Count_of_Specimens)"
    Alias ="Count_of_Plots"
    Expression ="Count(tbl_Events.Protocol_Name)"
    Alias ="Percent_Plots_Exotic"
    Expression ="Round([Count_of_Exotic_Plots]*100/[Count_of_Plots])"
    Alias ="ExoticPlots_and_Plots"
    Expression ="[Count_of_Exotic_Plots] & \" / \" & [Count_of_Plots]"
End
Begin Joins
    LeftTable ="qCalc_Exotic_Count_By_Event"
    RightTable ="tbl_Events"
    Expression ="qCalc_Exotic_Count_By_Event.Event_ID=tbl_Events.Event_ID"
    Flag =3
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Admin_Unit_Code"
    Flag =0
    Expression ="Year([Event_Date])"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Admin_Unit_Code"
    GroupLevel =0
    Expression ="Year([Event_Date])"
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
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Plots"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Count_of_Exotic_Plots"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Percent_Plots_Exotic"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExoticPlots_and_Plots"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =55
    Right =1522
    Bottom =932
    Left =-1
    Top =-1
    Right =1466
    Bottom =581
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_Exotic_Count_By_Event"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
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
        Name ="tbl_Locations"
        Name =""
    End
End
