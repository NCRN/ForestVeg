Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_DBH"
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
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Status"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="tbl_Tree_DBH.DBH"
    Expression ="tbl_Tree_DBH.Live"
    Alias ="Habit"
    Expression ="\"Tree\""
    Alias ="Class"
    Expression ="\"Tree\""
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =1
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
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_DBH"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Tags.Tag"
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
dbMemo "OrderBy" ="[qExport_All_Tree_Stems].[Sample_Year]"
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
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =26
    Top =-3
    Right =1350
    Bottom =457
    Left =-1
    Top =-1
    Right =1292
    Bottom =240
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
        Left =871
        Top =106
        Right =1015
        Bottom =271
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =1041
        Top =90
        Right =1185
        Bottom =234
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
        Left =432
        Top =95
        Right =576
        Bottom =239
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =722
        Top =5
        Right =866
        Bottom =149
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
End
