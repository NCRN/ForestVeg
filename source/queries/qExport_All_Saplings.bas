Operation =1
Option =0
Where ="(((tbl_Sapling_Data.Sapling_Status)<>\"Removed from study\") AND ((tbl_Sapling_D"
    "ata.Habit)=\"Tree\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="qCalc_Basal_Area_per_Sapling"
    Name ="tbl_Events"
    Name ="tlu_Plants"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Unit_Group"
    Expression ="tbl_Locations.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Sample_Year"
    Expression ="Year([Event_Date])"
    Alias ="Date"
    Expression ="Format([tbl_Events].[Event_Date],\"yyyymmdd\")"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qCalc_Basal_Area_per_Sapling.StemsLive"
    Expression ="qCalc_Basal_Area_per_Sapling.SumLiveBasalArea_cm2"
    Expression ="qCalc_Basal_Area_per_Sapling.Equiv_Live_DBH_cm"
    Alias ="Status"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Sapling_Data.Habit"
    Expression ="tbl_Sapling_Data.Browsed"
    Expression ="tbl_Sapling_Data.Browsable"
End
Begin Joins
    LeftTable ="qCalc_Basal_Area_per_Sapling"
    RightTable ="tbl_Sapling_Data"
    Expression ="qCalc_Basal_Area_per_Sapling.Sapling_Data_ID = tbl_Sapling_Data.Sapling_Data_ID"
    Flag =3
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.StemsLive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1488
    Bottom =991
    Left =-1
    Top =-1
    Right =1456
    Bottom =273
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =4
        Top =7
        Right =162
        Bottom =243
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =1267
        Top =33
        Right =1466
        Bottom =211
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =194
        Top =7
        Right =338
        Bottom =245
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =1035
        Top =1
        Right =1219
        Bottom =308
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =430
        Top =6
        Right =587
        Bottom =173
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =725
        Top =0
        Right =882
        Bottom =308
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
