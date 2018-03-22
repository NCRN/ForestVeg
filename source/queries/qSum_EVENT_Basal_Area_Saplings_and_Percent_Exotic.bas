Operation =1
Option =0
Where ="(((qCalc_Basal_Area_per_Sapling.Dead)=\"N\")) OR (((qCalc_Basal_Area_per_Sapling"
    ".Dead) Is Null))"
Begin InputTables
    Name ="qCalc_Basal_Area_per_Sapling"
    Name ="qFiltered_Events"
    Name ="qFiltered_Locations"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Events.Event_Date"
    Expression ="qFiltered_Events.Event_Year"
    Alias ="BA_cm2"
    Expression ="Sum(Nz([SumBasalArea_cm2]))"
    Alias ="Exotic_BA_cm2"
    Expression ="Sum(IIf([Exotic]=True,[SumBasalArea_cm2],0))"
    Alias ="Percent_BA_Exotic"
    Expression ="Round(100*[Exotic_BA_cm2]/Sum([SumBasalArea_cm2]),1)"
    Alias ="Sapling_Count"
    Expression ="Count(qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2)"
    Alias ="Saplings_per_ha"
    Expression ="Round([Sapling_Count]/0.008482,1)"
    Alias ="Sapling_BA_m2_per_ha"
    Expression ="Round(([BA_cm2]/10000)/0.008482,1)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qCalc_Basal_Area_per_Sapling"
    Expression ="qFiltered_Events.Event_ID = qCalc_Basal_Area_per_Sapling.Event_ID"
    Flag =2
End
Begin Groups
    Expression ="qFiltered_Locations.Plot_Name"
    GroupLevel =0
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qFiltered_Events.Event_Date"
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
dbText "Description" ="Returns the total basal area of all tree saplings in a plot, and the percentage "
    "of that total that is made up of exotic species. Created for IAN NRCA reports."
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
        dbText "Name" ="qFiltered_Events.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_BA_Exotic"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="BA_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Exotic_BA_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sapling_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Saplings_per_ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_BA_m2_per_ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =48
    Top =26
    Right =1410
    Bottom =588
    Left =-1
    Top =-1
    Right =1330
    Bottom =292
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =552
        Top =7
        Right =696
        Bottom =231
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =248
        Top =16
        Right =426
        Bottom =282
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =57
        Top =17
        Right =201
        Bottom =364
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
End
