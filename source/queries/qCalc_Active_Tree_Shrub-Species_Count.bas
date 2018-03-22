Operation =1
Option =0
Begin InputTables
    Name ="qActive_Trees_and_Shrubs"
End
Begin OutputColumns
    Expression ="qActive_Trees_and_Shrubs.Sample_Year"
    Expression ="qActive_Trees_and_Shrubs.Habit"
    Expression ="qActive_Trees_and_Shrubs.TSN"
    Alias ="CountOfSpecimens"
    Expression ="Count(qActive_Trees_and_Shrubs.Location_ID)"
End
Begin Groups
    Expression ="qActive_Trees_and_Shrubs.Sample_Year"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.Habit"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.TSN"
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
        dbText "Name" ="qActive_Trees_and_Shrubs.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfSpecimens"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =109
    Top =100
    Right =1124
    Bottom =832
    Left =-1
    Top =-1
    Right =991
    Bottom =453
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =578
        Bottom =280
        Top =0
        Name ="qActive_Trees_and_Shrubs"
        Name =""
    End
End
