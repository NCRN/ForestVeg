Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
End
Begin OutputColumns
    Alias ="WPt"
    Expression ="\"<wpt lat=\"\"\" & [Lat_WGS84] & \"\"\" lon=\"\"\" & [Lon_WGS84] & \"\"\"><name"
        ">\" & [Plot_Name] & \"</name><cmt>Panel \" & [Panel] & \" Frame=\" & Left([Frame"
        "],1) & \" GRTS=\" & [GRTS_ORDER] & \"</cmt><sym>Park</sym></wpt>\""
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
        dbText "Name" ="WPt"
        dbInteger "ColumnWidth" ="18000"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =37
    Top =69
    Right =1494
    Bottom =844
    Left =-1
    Top =-1
    Right =1425
    Bottom =473
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =252
        Bottom =534
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
