Operation =4
Option =0
Where ="(((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\") AND ((tbl_Tags.Tag_Status)<>"
    "\"Retired (In Office)\"))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Sapling_Data"
End
Begin OutputColumns
    Name ="tbl_Tags.Tag_Status"
    Expression ="\"Retired (In Office)\""
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbInteger "ColumnWidth" ="3705"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1462
    Bottom =910
    Left =-1
    Top =-1
    Right =1430
    Bottom =587
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =307
        Bottom =408
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =649
        Bottom =454
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
End
