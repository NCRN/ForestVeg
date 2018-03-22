Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tlu_Plants"
    Name ="qActive_Shrub_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="qActive_Shrub_Data.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([tbl_events].[Event_Date],\"yyyymmdd\"))"
    Expression ="qActive_Shrub_Data.Tag"
    Expression ="qActive_Shrub_Data.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qActive_Shrub_Data.Sapling_Status"
    Expression ="qActive_Shrub_Data.Habit"
    Expression ="qActive_Shrub_Data.Browsed"
    Expression ="qActive_Shrub_Data.Browsable"
    Expression ="qActive_Shrub_Data.Sapling_Data_ID"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qActive_Shrub_Data"
    Expression ="tlu_Plants.TSN = qActive_Shrub_Data.TSN"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="qActive_Shrub_Data"
    Expression ="tbl_Events.Event_ID = qActive_Shrub_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Sample_Year"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =22
    Top =484
    Right =1365
    Bottom =929
    Left =-1
    Top =-1
    Right =1311
    Bottom =141
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
        Left =192
        Top =6
        Right =336
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
    Begin
        Left =435
        Top =20
        Right =579
        Bottom =164
        Top =0
        Name ="qActive_Shrub_Data"
        Name =""
    End
End
