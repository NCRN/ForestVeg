Operation =1
Option =1
Where ="(((tbl_Tags.[RFS])=True) AND ((tbl_Tags.[Tag_Status])<>'Removed from study'))"
Begin InputTables
    Name ="tbl_Tags"
End
Begin OutputColumns
End
Begin OrderBy
    Expression ="tbl_Tags.Tag_Status"
    Flag =0
    Expression ="tbl_Tags.Start_Date"
    Flag =0
    Expression ="tbl_Tags.Tag"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query10].[Updated_Date] Is Null)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Notes"
        dbInteger "ColumnWidth" ="15165"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="tbl_Tags.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =1
    Left =480
    Top =1040
    Right =640
    Bottom =1068
    Left =-1
    Top =-1
    Right =1246
    Bottom =575
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
