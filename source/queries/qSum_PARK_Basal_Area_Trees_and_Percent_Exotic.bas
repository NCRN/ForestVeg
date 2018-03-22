Operation =1
Option =0
Where ="(((qCalc_Basal_Area_per_Tree.Dead)=\"N\")) OR (((qCalc_Basal_Area_per_Tree.Dead)"
    " Is Null))"
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qCalc_Basal_Area_per_Tree"
    Name ="qSum_PARK_Event_Count"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Event_Count"
    Expression ="Max(qSum_PARK_Event_Count.Plot_Count)"
    Alias ="BA_cm2"
    Expression ="Round(Sum(Nz([SumBasalArea_cm2])),0)"
    Alias ="Exotic_BA_cm2"
    Expression ="Round(Sum(IIf([Exotic]=True,[SumBasalArea_cm2],0)),0)"
    Alias ="Percent_BA_Exotic"
    Expression ="Round(100*[Exotic_BA_cm2]/Sum([SumBasalArea_cm2]),1)"
    Alias ="Tree_Count"
    Expression ="Count(qCalc_Basal_Area_per_Tree.SumBasalArea_cm2)"
    Alias ="Trees_per_ha"
    Expression ="Round([Tree_Count]/([Event_Count]*0.07686),1)"
    Alias ="Tree_BA_m2_per_ha"
    Expression ="Round(([BA_cm2]/10000)/([Event_Count]*0.070686),1)"
End
Begin Joins
    LeftTable ="qFiltered_Events"
    RightTable ="qCalc_Basal_Area_per_Tree"
    Expression ="qFiltered_Events.Event_ID = qCalc_Basal_Area_per_Tree.Event_ID"
    Flag =2
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Locations"
    RightTable ="qSum_PARK_Event_Count"
    Expression ="qFiltered_Locations.Admin_Unit_Code = qSum_PARK_Event_Count.Admin_Unit_Code"
    Flag =2
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
dbText "Description" ="Returns the total basal area of all trees in a plot, and the percentage of that "
    "total that is made up of exotic species. Created for IAN NRCA reports."
Begin
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_BA_Exotic"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic_BA_cm2"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BA_cm2"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_Count"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trees_per_ha"
        dbInteger "ColumnWidth" ="2835"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_BA_m2_per_ha"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =71
    Top =362
    Right =1237
    Bottom =926
    Left =-1
    Top =-1
    Right =1134
    Bottom =235
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =7
        Right =198
        Bottom =395
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =283
        Top =32
        Right =462
        Bottom =469
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =543
        Top =38
        Right =765
        Bottom =252
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
    Begin
        Left =792
        Top =91
        Right =936
        Bottom =235
        Top =0
        Name ="qSum_PARK_Event_Count"
        Name =""
    End
End
