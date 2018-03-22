Operation =1
Option =0
Where ="(((tbl_Tags.Tag) In (SELECT [Tag] FROM [tbl_Tags] As Tmp GROUP BY [Tag] HAVING C"
    "ount(*)>1 )))"
Begin InputTables
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Tag_ID"
    Expression ="tbl_Tags.Location_ID"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tags.Microplot_Number"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tags.Updated_Date"
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
dbText "Description" ="Duplicate tag numbers exist in tbl_Tags."
Begin
    Begin
        dbText "Name" ="tbl_Tags.[Tag]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Tag_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Location_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Azimuth]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Distance]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Microplot_Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Tag_Status]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.[Updated_Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Location_ID"
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
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =233
    Right =870
    Bottom =662
    Left =-1
    Top =-1
    Right =830
    Bottom =196
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
