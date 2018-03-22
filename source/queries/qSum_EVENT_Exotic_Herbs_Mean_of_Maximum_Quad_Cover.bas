Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qCalc_Exotic_Herbs_Max_Cover_by_Quadrat"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Events.Event_Year"
    Alias ="Quadrats_with_Exotic_Herbs"
    Expression ="Count(qCalc_Exotic_Herbs_Max_Cover_by_Quadrat.Quadrat_Number)"
    Alias ="Mean_of_Maximum_Exotic_Quadrat_Cover_in_All_Quads"
    Expression ="Round(Sum(Nz([MaxOfPercent_Cover]))/12,1)"
    Alias ="Mean_of_Maximum_Exotic_Quadrat_Cover_in_Exotic_Quads_Only"
    Expression ="Round(Sum([MaxOfPercent_Cover])/[Quadrats_with_Exotic_Herbs],1)"
    Alias ="Exotic_Herb_Presence"
    Expression ="IIf([Quadrats_with_Exotic_Herbs]>0,\"Present\",\"Absent\")"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qCalc_Exotic_Herbs_Max_Cover_by_Quadrat"
    Expression ="qFiltered_Events.Event_ID = qCalc_Exotic_Herbs_Max_Cover_by_Quadrat.Event_ID"
    Flag =2
End
Begin OrderBy
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Flag =0
    Expression ="qFiltered_Locations.Plot_Name"
    Flag =0
    Expression ="qFiltered_Events.Event_Year"
    Flag =0
End
Begin Groups
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    GroupLevel =0
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
Begin
    Begin
        dbText "Name" ="Quadrats_with_Exotic_Herbs"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_of_Maximum_Exotic_Quadrat_Cover_in_Exotic_Quads_Only"
        dbInteger "ColumnWidth" ="6510"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_of_Maximum_Exotic_Quadrat_Cover_in_All_Quads"
        dbInteger "ColumnWidth" ="5175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic_Herb_Presence"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =64
    Top =85
    Right =1019
    Bottom =767
    Left =-1
    Top =-1
    Right =923
    Bottom =290
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =14
        Top =16
        Right =158
        Bottom =160
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =188
        Top =32
        Right =331
        Bottom =477
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =424
        Top =17
        Right =716
        Bottom =238
        Top =0
        Name ="qCalc_Exotic_Herbs_Max_Cover_by_Quadrat"
        Name =""
    End
End
