Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="qCalc_Basal_Area_per_Sapling"
    Name ="qActive_Sapling_Data"
    Name ="tbl_Events"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Subunit_Code"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="qActive_Sapling_Data.Sample_Year"
    Alias ="Date"
    Expression ="Format([tbl_Events].[Event_Date],\"yyyymmdd\")"
    Expression ="qActive_Sapling_Data.Tag"
    Expression ="qActive_Sapling_Data.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qCalc_Basal_Area_per_Sapling.Stems"
    Expression ="qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2"
    Expression ="qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm"
    Expression ="qActive_Sapling_Data.Sapling_Status"
    Expression ="qActive_Sapling_Data.Habit"
    Expression ="qActive_Sapling_Data.Browsed"
    Expression ="qActive_Sapling_Data.Browsable"
End
Begin Joins
    LeftTable ="qActive_Sapling_Data"
    RightTable ="tbl_Events"
    Expression ="qActive_Sapling_Data.Event_ID = tbl_Events.Event_ID"
    Flag =1
    LeftTable ="qActive_Sapling_Data"
    RightTable ="tlu_Plants"
    Expression ="qActive_Sapling_Data.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="qCalc_Basal_Area_per_Sapling"
    RightTable ="qActive_Sapling_Data"
    Expression ="qCalc_Basal_Area_per_Sapling.Sapling_Data_ID = qActive_Sapling_Data.Sapling_Data"
        "_ID"
    Flag =3
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2"
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
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Stems"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Sample_Year"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Browsable"
        dbLong "AggregateType" ="-1"
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
End
Begin
    State =0
    Left =23
    Top =387
    Right =1542
    Bottom =832
    Left =-1
    Top =-1
    Right =1487
    Bottom =158
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
        Left =554
        Top =68
        Right =753
        Bottom =246
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =352
        Top =4
        Right =525
        Bottom =253
        Top =0
        Name ="qActive_Sapling_Data"
        Name =""
    End
    Begin
        Left =178
        Top =6
        Right =322
        Bottom =244
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =779
        Top =-1
        Right =951
        Bottom =248
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
