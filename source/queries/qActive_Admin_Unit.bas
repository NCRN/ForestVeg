Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="tbl_Locations.Admin_Unit_Code"
End
Begin OrderBy
    Expression ="tbl_Locations.Admin_Unit_Code"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Admin_Unit_Code"
    GroupLevel =0
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
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =56
    Top =148
    Right =855
    Bottom =690
    Left =-1
    Top =-1
    Right =767
    Bottom =251
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =316
        Bottom =205
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
