Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Soil_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Soil_Data.[Leaf_Litter_T_Carbon_%]"
    Expression ="tbl_Soil_Data.[Leaf_Litter_T_Nitrogen_%]"
    Expression ="tbl_Soil_Data.Soil_Notes"
    Expression ="tbl_Soil_Data.Infiltration_Ring_Single_Area_cm2"
    Expression ="tbl_Soil_Data.Bulk_Density_Ring_Single_Volume_cm3"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="tbl_Soil_Data"
    Expression ="tbl_Events.Event_ID=tbl_Soil_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Events.Event_Date"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qActive_Soil_Plots].[Event_Date]"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbInteger "ColumnWidth" ="870"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbInteger "ColumnWidth" ="555"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Soil_Data.[Leaf_Litter_T_Carbon_%]"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Soil_Data.[Leaf_Litter_T_Nitrogen_%]"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Soil_Data.Soil_Notes"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Soil_Data.Infiltration_Ring_Single_Area_cm2"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Soil_Data.Bulk_Density_Ring_Single_Volume_cm3"
        dbInteger "ColumnWidth" ="1290"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =245
    Right =1541
    Bottom =860
    Left =-1
    Top =-1
    Right =1470
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Locations"
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
        Right =766
        Bottom =156
        Top =0
        Name ="tbl_Soil_Data"
        Name =""
    End
End
