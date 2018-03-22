Operation =1
Option =0
Where ="(((tbl_Tags.Azimuth) Is Null) AND ((tbl_Tags_0830.Azimuth) Is Not Null)) OR (((t"
    "bl_Tags.Azimuth) Is Null) AND ((tbl_Tags_0628.Azimuth) Is Not Null)) OR (((tbl_T"
    "ags.Azimuth) Is Null) AND ((tbl_Tags_0730.Azimuth) Is Not Null)) OR (((tbl_Tags."
    "Azimuth) Is Null) AND ((tbl_Tags_0607.Azimuth) Is Not Null)) OR (((tbl_Tags.Azim"
    "uth) Is Null) AND ((tbl_Tags_0528.Azimuth) Is Not Null))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Tags_0830"
    Name ="tbl_Tags_0628"
    Name ="tbl_Tags_0730"
    Name ="tbl_Tags_0528"
    Name ="tbl_Tags_0607"
End
Begin OutputColumns
    Expression ="tbl_Tags_0830.Tag"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tags_0830.Azimuth"
    Expression ="tbl_Tags_0830.Distance"
    Expression ="tbl_Tags_0730.Azimuth"
    Expression ="tbl_Tags_0730.Distance"
    Expression ="tbl_Tags_0628.Azimuth"
    Expression ="tbl_Tags_0628.Distance"
    Expression ="tbl_Tags_0607.Azimuth"
    Expression ="tbl_Tags_0607.Distance"
    Expression ="tbl_Tags_0528.Azimuth"
    Expression ="tbl_Tags_0528.Distance"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tags_0830"
    Expression ="tbl_Tags.Tag_ID = tbl_Tags_0830.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags_0628"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tags_0628.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags_0730"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tags_0730.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags_0528"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tags_0528.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags_0607"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tags_0607.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Tags_0830.Tag"
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
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="495"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="435"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0830.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_0830.Azimuth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0830.Distance"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="675"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0628.Azimuth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="285"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0628.Distance"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="330"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0730.Azimuth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0730.Distance"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="660"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0607.Azimuth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="510"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0607.Distance"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="510"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0528.Azimuth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="510"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_0528.Distance"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="510"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =57
    Top =81
    Right =1440
    Bottom =901
    Left =-1
    Top =-1
    Right =1351
    Bottom =520
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =200
        Bottom =284
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =238
        Top =129
        Right =525
        Bottom =441
        Top =0
        Name ="tbl_Tags_0830"
        Name =""
    End
    Begin
        Left =580
        Top =121
        Right =748
        Bottom =457
        Top =0
        Name ="tbl_Tags_0628"
        Name =""
    End
    Begin
        Left =767
        Top =12
        Right =1043
        Bottom =361
        Top =0
        Name ="tbl_Tags_0730"
        Name =""
    End
    Begin
        Left =1091
        Top =12
        Right =1235
        Bottom =156
        Top =0
        Name ="tbl_Tags_0528"
        Name =""
    End
    Begin
        Left =1091
        Top =156
        Right =1235
        Bottom =300
        Top =0
        Name ="tbl_Tags_0607"
        Name =""
    End
End
