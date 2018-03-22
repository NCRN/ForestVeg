Operation =1
Option =0
Where ="(((Year([Event_Date]))=2008))"
Begin InputTables
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="tbl_Tree_Data.Tag_ID"
    Expression ="tbl_Tree_Data.Tree_Data_ID"
    Alias ="Crown_Class_06"
    Expression ="tbl_Tree_Data.Crown_Class"
    Alias ="Tree_Status_06"
    Expression ="tbl_Tree_Data.Tree_Status"
    Alias ="Tag_Status_06"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tree_Data.Updated_Date"
    Alias ="Sample_Year_06"
    Expression ="Year([Event_Date])"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
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
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Crown_Class_06"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_Status_06"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tag_Status_06"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year_06"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =90
    Right =1434
    Bottom =883
    Left =-1
    Top =-1
    Right =1359
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =656
        Top =12
        Right =895
        Bottom =427
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =365
        Top =15
        Right =608
        Bottom =265
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =95
        Top =12
        Right =239
        Bottom =262
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
