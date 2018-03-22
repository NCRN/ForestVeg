Operation =1
Option =0
Where ="(((tbl_Tags.Tag_Status)=\"Retired (In Office)\"))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Locations"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Tags.Tag"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Location_Status"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Locations"
    Expression ="tbl_Tags.Location_ID = tbl_Locations.Location_ID"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
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
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Status"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Notes"
        dbInteger "ColumnWidth" ="5145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =47
    Top =9
    Right =1282
    Bottom =948
    Left =-1
    Top =-1
    Right =1203
    Bottom =605
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =284
        Top =90
        Right =428
        Bottom =234
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =64
        Top =78
        Right =208
        Bottom =222
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =695
        Top =26
        Right =839
        Bottom =170
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
