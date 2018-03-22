Operation =1
Option =0
Where ="(((qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm)>=10))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="qCalc_Basal_Area_per_Sapling"
    Name ="tbl_Tags"
    Name ="tbl_Sapling_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Microplot_Number"
    Expression ="qCalc_Basal_Area_per_Sapling.Stems"
    Expression ="qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2"
    Expression ="qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Sapling_Data.Habit"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="qCalc_Basal_Area_per_Sapling"
    Expression ="tbl_Events.Event_ID = qCalc_Basal_Area_per_Sapling.Event_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="qCalc_Basal_Area_per_Sapling"
    RightTable ="tbl_Tags"
    Expression ="qCalc_Basal_Area_per_Sapling.FirstOfTag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Events.Event_Date"
    Flag =0
    Expression ="tbl_Tags.Tag"
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
dbText "Description" ="Sapling equivalent DBH is greater than 10cm"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Stems"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =233
    Right =870
    Bottom =662
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =12
        Top =10
        Right =156
        Bottom =259
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =210
        Top =10
        Right =354
        Bottom =291
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =384
        Top =18
        Right =528
        Bottom =162
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =565
        Top =19
        Right =709
        Bottom =229
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =754
        Top =19
        Right =898
        Bottom =242
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
End
