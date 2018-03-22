Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_Vines"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Alias ="tlu_Plants_1"
    Name ="tlu_Plants"
    Name ="qCalc_Trees_with_Vines_in_Crown"
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
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tree_Vines.TSN"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Host_Tag"
    Expression ="tbl_Tags.Tag"
    Alias ="Host_TSN"
    Expression ="tbl_Tags.TSN"
    Alias ="Host_Latin_Name"
    Expression ="tlu_Plants_1.Latin_Name"
    Alias ="Host_Status"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="qCalc_Trees_with_Vines_in_Crown.Condition"
    Expression ="tlu_Plants.Exotic"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants_1"
    Expression ="tbl_Tags.TSN = tlu_Plants_1.TSN"
    Flag =1
    LeftTable ="tbl_Tree_Vines"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tree_Vines.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="qCalc_Trees_with_Vines_in_Crown"
    Expression ="tbl_Tree_Data.Tree_Data_ID = qCalc_Trees_with_Vines_in_Crown.Tree_Data_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_Vines"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Vines.Tree_Data_ID"
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
        dbText "Name" ="tbl_Tree_Vines.TSN"
        dbInteger "ColumnWidth" ="855"
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
        dbText "Name" ="qCalc_Trees_with_Vines_in_Crown.Condition"
        dbInteger "ColumnWidth" ="1680"
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
        dbText "Name" ="Host_Tree_Tag"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
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
    Left =207
    Top =10
    Right =1502
    Bottom =803
    Left =-1
    Top =-1
    Right =1263
    Bottom =436
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Tree_Vines"
        Name =""
    End
    Begin
        Left =431
        Top =181
        Right =608
        Bottom =441
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =627
        Top =181
        Right =771
        Bottom =325
        Top =0
        Name ="tlu_Plants_1"
        Name =""
    End
    Begin
        Left =51
        Top =189
        Right =195
        Bottom =333
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =230
        Top =187
        Right =387
        Bottom =339
        Top =0
        Name ="qCalc_Trees_with_Vines_in_Crown"
        Name =""
    End
End
