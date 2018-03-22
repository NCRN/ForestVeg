Operation =1
Option =0
Begin InputTables
    Name ="qSum_PARK_Event_List"
End
Begin OutputColumns
    Expression ="qSum_PARK_Event_List.Admin_Unit_Code"
    Alias ="Plot_Count"
    Expression ="Count(qSum_PARK_Event_List.Certified)"
End
Begin OrderBy
    Expression ="qSum_PARK_Event_List.Admin_Unit_Code"
    Flag =0
End
Begin Groups
    Expression ="qSum_PARK_Event_List.Admin_Unit_Code"
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
        dbText "Name" ="qSum_PARK_Event_List.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1041
    Bottom =875
    Left =-1
    Top =-1
    Right =933
    Bottom =460
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =435
        Bottom =308
        Top =0
        Name ="qSum_PARK_Event_List"
        Name =""
    End
End
