Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Tags"
    Name ="qSum_Trees_with_Vines_in_Crown"
    Name ="qCalc_Basal_Area_Per_Tree"
    Name ="tlu_Plants"
    Name ="tbl_Events"
    Name ="qList_Crown_Class_Descriptions"
    Name ="tbl_Tree_Data"
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
    Expression ="CLng(Format([tbl_Events].[Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qCalc_Basal_Area_Per_Tree.Stems"
    Expression ="qCalc_Basal_Area_Per_Tree.SumLiveBasalArea_cm2"
    Expression ="qCalc_Basal_Area_Per_Tree.SumDeadBasalArea_cm2"
    Expression ="qCalc_Basal_Area_Per_Tree.Equiv_Live_DBH_cm"
    Expression ="qCalc_Basal_Area_Per_Tree.Equiv_Dead_DBH_cm"
    Expression ="qSum_Trees_with_Vines_in_Crown.Condition"
    Alias ="Status"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="tbl_Tree_Data.Crown_Class"
    Expression ="qList_Crown_Class_Descriptions.Crown_Description"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="qList_Crown_Class_Descriptions"
    Expression ="tbl_Tree_Data.Crown_Class = qList_Crown_Class_Descriptions.Crown_Class"
    Flag =2
    LeftTable ="tbl_Tree_Data"
    RightTable ="qCalc_Basal_Area_Per_Tree"
    Expression ="tbl_Tree_Data.Tree_Data_ID = qCalc_Basal_Area_Per_Tree.Tree_Data_ID"
    Flag =2
    LeftTable ="tbl_Tree_Data"
    RightTable ="qSum_Trees_with_Vines_in_Crown"
    Expression ="tbl_Tree_Data.Tree_Data_ID = qSum_Trees_with_Vines_in_Crown.Tree_Data_ID"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
End
Begin OrderBy
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
Begin
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="705"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Per_Tree.Stems"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qSum_Trees_with_Vines_in_Crown.Condition"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qList_Crown_Class_Descriptions.Crown_Description"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Crown_Class"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Per_Tree.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Per_Tree.SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Per_Tree.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Per_Tree.Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =11
    Top =8
    Right =1495
    Bottom =819
    Left =-1
    Top =-1
    Right =1452
    Bottom =462
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =7
        Top =31
        Right =151
        Bottom =175
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =701
        Top =362
        Right =845
        Bottom =527
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =701
        Top =190
        Right =959
        Bottom =343
        Top =0
        Name ="qSum_Trees_with_Vines_in_Crown"
        Name =""
    End
    Begin
        Left =809
        Top =5
        Right =1048
        Bottom =185
        Top =0
        Name ="qCalc_Basal_Area_Per_Tree"
        Name =""
    End
    Begin
        Left =1176
        Top =23
        Right =1377
        Bottom =463
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =181
        Top =32
        Right =325
        Bottom =176
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =184
        Top =235
        Right =328
        Bottom =379
        Top =0
        Name ="qList_Crown_Class_Descriptions"
        Name =""
    End
    Begin
        Left =432
        Top =95
        Right =576
        Bottom =239
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
End
