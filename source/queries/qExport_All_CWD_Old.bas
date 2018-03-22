Operation =1
Option =0
Begin InputTables
    Name ="qActive_CWD_Data"
End
Begin OutputColumns
    Expression ="qActive_CWD_Data.Plot_Name"
    Expression ="qActive_CWD_Data.Unit_Code"
    Expression ="qActive_CWD_Data.Admin_Unit_Code"
    Expression ="qActive_CWD_Data.Subunit_Code"
    Expression ="qActive_CWD_Data.Panel"
    Expression ="qActive_CWD_Data.Frame"
    Expression ="qActive_CWD_Data.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Expression ="qActive_CWD_Data.Transect_Azimuth"
    Expression ="qActive_CWD_Data.TSN"
    Expression ="qActive_CWD_Data.Latin_Name"
    Expression ="qActive_CWD_Data.Decay_Class"
    Expression ="qActive_CWD_Data.Diameter"
    Expression ="qActive_CWD_Data.Hollow"
    Expression ="qActive_CWD_Data.CWD_Notes"
End
Begin OrderBy
    Expression ="qActive_CWD_Data.Plot_Name"
    Flag =0
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
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
        dbText "Name" ="qActive_CWD_Data.Unit_Code"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Sample_Year"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Transect_Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.TSN"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Latin_Name"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Decay_Class"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Diameter"
        dbInteger "ColumnWidth" ="1185"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Hollow"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.CWD_Notes"
        dbInteger "ColumnWidth" ="3195"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_CWD_Data.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =49
    Top =146
    Right =1460
    Bottom =884
    Left =-1
    Top =-1
    Right =1379
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qActive_CWD_Data"
        Name =""
    End
End
