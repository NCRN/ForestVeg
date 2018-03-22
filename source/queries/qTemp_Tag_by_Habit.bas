Operation =1
Option =0
Begin InputTables
    Name ="qAll_Trees_and_Shrubs_Tagged"
End
Begin OutputColumns
    Expression ="qAll_Trees_and_Shrubs_Tagged.Tag"
    Expression ="qAll_Trees_and_Shrubs_Tagged.Habit"
End
Begin OrderBy
    Expression ="qAll_Trees_and_Shrubs_Tagged.Tag"
    Flag =0
End
Begin Groups
    Expression ="qAll_Trees_and_Shrubs_Tagged.Tag"
    GroupLevel =0
    Expression ="qAll_Trees_and_Shrubs_Tagged.Habit"
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
        dbText "Name" ="qAll_Trees_and_Shrubs_Tagged.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qAll_Trees_and_Shrubs_Tagged.Habit"
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
        Right =265
        Bottom =259
        Top =0
        Name ="qAll_Trees_and_Shrubs_Tagged"
        Name =""
    End
End
