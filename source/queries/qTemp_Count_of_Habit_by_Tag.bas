Operation =1
Option =0
Having ="(((Count(qTemp_Tag_by_Habit.Tag))>1))"
Begin InputTables
    Name ="qTemp_Tag_by_Habit"
End
Begin OutputColumns
    Expression ="qTemp_Tag_by_Habit.Tag"
    Alias ="CountOfTag"
    Expression ="Count(qTemp_Tag_by_Habit.Tag)"
End
Begin OrderBy
    Expression ="qTemp_Tag_by_Habit.Tag"
    Flag =0
End
Begin Groups
    Expression ="qTemp_Tag_by_Habit.Tag"
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
        dbText "Name" ="qTemp_Tag_by_Habit.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTag"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1439
    Bottom =966
    Left =-1
    Top =-1
    Right =1293
    Bottom =505
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qTemp_Tag_by_Habit"
        Name =""
    End
End
