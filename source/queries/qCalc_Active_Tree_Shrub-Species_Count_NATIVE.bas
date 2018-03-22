Operation =1
Option =0
Where ="(((tlu_Plants.Exotic)=False))"
Begin InputTables
    Name ="qCalc_Active_Tree_Shrub-Species_Count"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].Sample_Year"
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].Habit"
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].TSN"
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].CountOfSpecimens"
    Expression ="tlu_Plants.Exotic"
End
Begin Joins
    LeftTable ="qCalc_Active_Tree_Shrub-Species_Count"
    RightTable ="tlu_Plants"
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].TSN = tlu_Plants.TSN"
    Flag =2
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
        dbText "Name" ="[qCalc_Active_Tree_Shrub-Species_Count].Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qCalc_Active_Tree_Shrub-Species_Count].Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qCalc_Active_Tree_Shrub-Species_Count].TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qCalc_Active_Tree_Shrub-Species_Count].CountOfSpecimens"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1463
    Bottom =997
    Left =-1
    Top =-1
    Right =1431
    Bottom =640
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =376
        Bottom =190
        Top =0
        Name ="qCalc_Active_Tree_Shrub-Species_Count"
        Name =""
    End
    Begin
        Left =424
        Top =12
        Right =598
        Bottom =438
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
