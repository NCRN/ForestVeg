Operation =6
Option =0
Begin InputTables
    Name ="qSum_Tree_and_Sapling_Cycles"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="qSum_Tree_and_Sapling_Cycles.Tag_ID"
    GroupLevel =2
    Expression ="tbl_Tags.Tag"
    GroupLevel =2
    Expression ="qSum_Tree_and_Sapling_Cycles.Cycle"
    GroupLevel =1
    Alias ="MaxOfEvent_Year"
    Expression ="Max(qSum_Tree_and_Sapling_Cycles.Event_Year)"
    GroupLevel =2
    Alias ="FirstOfClass_Status"
    Expression ="First(qSum_Tree_and_Sapling_Cycles.Class_Status)"
End
Begin Joins
    LeftTable ="qSum_Tree_and_Sapling_Cycles"
    RightTable ="tbl_Tags"
    Expression ="qSum_Tree_and_Sapling_Cycles.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
End
Begin Groups
    Expression ="qSum_Tree_and_Sapling_Cycles.Tag_ID"
    GroupLevel =2
    Expression ="tbl_Tags.Tag"
    GroupLevel =2
    Expression ="tbl_Tags.Tag_Status"
    GroupLevel =2
    Expression ="qSum_Tree_and_Sapling_Cycles.Cycle"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Class_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbInteger "ColumnWidth" ="3300"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbInteger "ColumnWidth" ="4290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfEvent_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1518
    Bottom =965
    Left =-1
    Top =-1
    Right =1486
    Bottom =591
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =275
        Bottom =223
        Top =0
        Name ="qSum_Tree_and_Sapling_Cycles"
        Name =""
    End
    Begin
        Left =323
        Top =12
        Right =583
        Bottom =356
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
