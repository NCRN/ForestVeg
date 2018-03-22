Operation =1
Option =16
RowCount ="100"
Begin InputTables
    Name ="tlu_Plants"
    Name ="tbl_Events"
    Name ="tbl_Locations"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_DBH"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Alias ="DBH (cm)"
    Expression ="tbl_Tree_DBH.DBH"
    Alias ="Species"
    Expression ="[Latin_Name] & \" (\" & [Common] & \")\""
    Alias ="Location"
    Expression ="[Plot_Name] & \" #\" & [Tag]"
    Expression ="tbl_Tree_Data.Tree_Status"
    Alias ="Year"
    Expression ="Year([Event_Date])"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
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
    Expression ="tbl_Tree_DBH.DBH"
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
dbText "Description" ="What are the 100 largest STEMS ever measured during NCRN monitoring?"
Begin
    Begin
        dbText "Name" ="DBH (cm)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species"
        dbInteger "ColumnWidth" ="4215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =72
    Right =950
    Bottom =558
    Left =-1
    Top =-1
    Right =910
    Bottom =249
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =740
        Top =56
        Right =902
        Bottom =379
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =172
        Top =11
        Right =316
        Bottom =155
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =5
        Top =9
        Right =149
        Bottom =209
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =341
        Top =14
        Right =527
        Bottom =246
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =555
        Top =9
        Right =699
        Bottom =153
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
    Begin
        Left =529
        Top =159
        Right =673
        Bottom =303
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
