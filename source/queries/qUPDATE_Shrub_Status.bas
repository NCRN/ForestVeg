Operation =4
Option =0
Where ="(((tlu_Plants.Genus)=\"Lindera\" Or (tlu_Plants.Genus)=\"Kalmia\"))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Sapling_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Name ="tbl_Tags.Tag_Notes"
    Expression ="\"No longer monitoring as individual shrubs - Tag retired\""
    Name ="tbl_Tags.Stop_Date"
    Expression ="#4/1/2015#"
    Name ="tbl_Tags.Tag_Status"
    Expression ="\"Retired (In Office)\""
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="tlu_Plants"
    RightTable ="tbl_Tags"
    Expression ="tlu_Plants.TSN = tbl_Tags.TSN"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Notes"
        dbInteger "ColumnWidth" ="5925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Stop_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.RFS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Genus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =41
    Top =16
    Right =1612
    Bottom =842
    Left =-1
    Top =-1
    Right =1539
    Bottom =543
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =461
        Top =16
        Right =772
        Bottom =469
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =175
        Top =9
        Right =384
        Bottom =403
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =901
        Top =-7
        Right =1167
        Bottom =511
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
