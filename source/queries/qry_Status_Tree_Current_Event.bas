Operation =1
Option =0
Where ="(((tbl_Events.Event_ID)=[TempVars]![EventID]))"
Begin InputTables
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
End
Begin OutputColumns
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Tree_Data.Tag_ID"
    Expression ="tbl_Tree_Data.Tree_Status"
End
Begin Joins
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
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =91
    Top =-27
    Right =1180
    Bottom =632
    Left =-1
    Top =-1
    Right =1071
    Bottom =363
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =284
        Top =43
        Right =428
        Bottom =187
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
End
