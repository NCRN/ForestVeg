Operation =1
Option =0
Where ="(((tbl_Tags.Tag)<>Int([Value_New])) AND ((tbl_Tags_History.Field_Name)=\"Tag\"))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Tags_History"
End
Begin OutputColumns
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags_History.Field_Name"
    Expression ="tbl_Tags_History.Value_New"
    Expression ="tbl_Tags_History.Value_Old"
    Expression ="tbl_Tags_History.Value_History_Notes"
    Expression ="tbl_Tags_History.Change_Date"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tags_History"
    Expression ="tbl_Tags.Tag_ID = tbl_Tags_History.Record_ID"
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
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_New"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_Old"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_History_Notes"
        dbInteger "ColumnWidth" ="4980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Field_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Change_Date"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =119
    Top =187
    Right =1164
    Bottom =882
    Left =-1
    Top =-1
    Right =1013
    Bottom =442
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =184
        Bottom =323
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =534
        Bottom =338
        Top =0
        Name ="tbl_Tags_History"
        Name =""
    End
End
