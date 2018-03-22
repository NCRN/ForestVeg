Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Alias ="tlu_Plants_1"
    Name ="tlu_Plants"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Sapling_Vines"
    Name ="qCalc_Saplings_with_Vines_in_Crown"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Unit_Group"
    Expression ="tbl_Locations.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([tbl_Events]![Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Sample_Year"
    Expression ="Year([tbl_Events].[Event_Date])"
    Alias ="Date"
    Expression ="CLng(Format([tbl_Events]![Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Sapling_Vines.TSN"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tags.Tag_Status"
    Alias ="Host_Tag"
    Expression ="tbl_Tags.Tag"
    Alias ="Host_TSN"
    Expression ="tbl_Tags.TSN"
    Alias ="Host_Latin_Name"
    Expression ="tlu_Plants_1.Latin_Name"
    Alias ="Host_Status"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="qCalc_Saplings_with_Vines_in_Crown.Condition"
    Expression ="tlu_Plants.Exotic"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants_1"
    Expression ="tbl_Tags.TSN = tlu_Plants_1.TSN"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Sapling_Data"
    RightTable ="qCalc_Saplings_with_Vines_in_Crown"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID = qCalc_Saplings_with_Vines_in_Crown.Sapling_Da"
        "ta_ID"
    Flag =2
    LeftTable ="tbl_Sapling_Vines"
    RightTable ="tlu_Plants"
    Expression ="tbl_Sapling_Vines.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Sapling_Vines"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_Vines.Sapling_Data_ID"
    Flag =1
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
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_TSN"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Vines.TSN"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Saplings_with_Vines_in_Crown.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Tag"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =70
    Top =-1
    Right =1365
    Bottom =547
    Left =-1
    Top =-1
    Right =1263
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =429
        Top =17
        Right =573
        Bottom =161
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =497
        Top =178
        Right =725
        Bottom =451
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =860
        Top =144
        Right =1004
        Bottom =288
        Top =0
        Name ="tlu_Plants_1"
        Name =""
    End
    Begin
        Left =51
        Top =178
        Right =195
        Bottom =333
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =214
        Top =28
        Right =358
        Bottom =172
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =52
        Top =24
        Right =196
        Bottom =168
        Top =0
        Name ="tbl_Sapling_Vines"
        Name =""
    End
    Begin
        Left =258
        Top =133
        Right =402
        Bottom =394
        Top =0
        Name ="qCalc_Saplings_with_Vines_in_Crown"
        Name =""
    End
End
