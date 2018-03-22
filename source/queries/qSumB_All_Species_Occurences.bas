Operation =1
Option =0
Begin InputTables
    Name ="qActive_Trees_Shrubs_Herbs_Vines"
    Name ="tbl_Locations"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Subunit_Code"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Habit"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Class"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.Common"
    Expression ="tlu_Plants.Exotic"
End
Begin Joins
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tbl_Locations"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Location_ID = tbl_Locations.Location_ID"
    Flag =1
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tlu_Plants"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN = tlu_Plants.TSN"
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
dbText "Description" ="Can I see a list of every species that has been identified during each sampling "
    "event?"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Habit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Class"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =72
    Right =1112
    Bottom =849
    Left =-1
    Top =-1
    Right =1072
    Bottom =509
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =310
        Bottom =457
        Top =0
        Name ="qActive_Trees_Shrubs_Herbs_Vines"
        Name =""
    End
    Begin
        Left =358
        Top =12
        Right =502
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =383
        Top =196
        Right =527
        Bottom =340
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
