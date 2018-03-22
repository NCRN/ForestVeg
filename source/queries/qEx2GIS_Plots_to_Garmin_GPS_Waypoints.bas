Operation =1
Option =0
Where ="(((tbl_Locations.Panel)>0))"
Begin InputTables
    Name ="tbl_Locations"
End
Begin OutputColumns
    Alias ="type"
    Expression ="\"WAYPOINT\""
    Alias ="ident"
    Expression ="tbl_Locations.Plot_Name"
    Alias ="lat"
    Expression ="tbl_Locations.Lat_WGS84"
    Alias ="long"
    Expression ="tbl_Locations.Lon_WGS84"
    Alias ="y_proj"
    Expression ="tbl_Locations.Y_Coord"
    Alias ="x_proj"
    Expression ="tbl_Locations.X_Coord"
    Alias ="comment"
    Expression ="\"Pan\" & [Panel] & \" Fr=\" & Left([Frame],1) & \" G=\" & [GRTS_ORDER]"
    Alias ="display"
    Expression ="0"
    Alias ="symbol"
    Expression ="159"
    Alias ="unused1"
    Expression ="0"
    Alias ="dist"
    Expression ="0"
    Alias ="prox_index"
    Expression ="0"
    Alias ="color"
    Expression ="31"
    Alias ="altitude"
    Expression ="0"
    Alias ="depth"
    Expression ="0"
    Alias ="wpt_class"
    Expression ="0"
    Alias ="sub_Class"
    Expression ="\"\""
    Alias ="attrib"
    Expression ="128"
    Alias ="link"
    Expression ="\"\""
    Alias ="state"
    Expression ="\"\""
    Alias ="county"
    Expression ="\"\""
    Alias ="city"
    Expression ="\"\""
    Alias ="address"
    Expression ="\"\""
    Alias ="facility"
    Expression ="\"\""
    Alias ="crossroad"
    Expression ="\"\""
    Alias ="unused2"
    Expression ="0"
    Alias ="ete"
    Expression ="-1"
    Alias ="dtype"
    Expression ="1"
    Alias ="model"
    Expression ="\"GPSMap60CSX\""
    Alias ="filename"
    Expression ="\"\""
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="y_proj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="x_proj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ident"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="long"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="comment"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="display"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="symbol"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="unused1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="prox_index"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="altitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="depth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="wpt_class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sub_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="attrib"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="link"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="state"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="county"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="city"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="address"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="facility"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="crossroad"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="unused2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="filename"
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
    Bottom =490
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
