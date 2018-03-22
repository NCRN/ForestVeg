Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_CWD_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Sample_Year"
    Expression ="Year([tbl_Events].[Event_Date])"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_CWD_Data.Transect_Azimuth"
    Expression ="tbl_CWD_Data.TSN"
    Expression ="tbl_CWD_Data.Decay_Class"
    Expression ="tbl_CWD_Data.Diameter"
    Expression ="tbl_CWD_Data.Hollow"
    Alias ="CWD_Notes"
    Expression ="tbl_CWD_Data.CWD_Notes"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_CWD_Data.Event_ID"
End
Begin Joins
    LeftTable ="tbl_CWD_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_CWD_Data.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_CWD_Data"
    Expression ="tbl_Events.Event_ID = tbl_CWD_Data.Event_ID"
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
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Transect_Azimuth"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Decay_Class"
        dbInteger "ColumnWidth" ="1290"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.TSN"
        dbInteger "ColumnWidth" ="810"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Diameter"
        dbInteger "ColumnWidth" ="1185"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Hollow"
        dbInteger "ColumnWidth" ="750"
        dbInteger "ColumnOrder" ="14"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CWD_Notes"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Cycle"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =75
    Top =146
    Right =1500
    Bottom =716
    Left =-1
    Top =-1
    Right =1393
    Bottom =185
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =436
        Top =6
        Right =691
        Bottom =266
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =249
        Top =4
        Right =384
        Bottom =141
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =195
        Bottom =183
        Top =0
        Name ="tbl_CWD_Data"
        Name =""
    End
    Begin
        Left =739
        Top =12
        Right =883
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
