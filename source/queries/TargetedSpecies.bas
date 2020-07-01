Operation =1
Option =0
Begin InputTables
    Name ="TargetSpecies"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="TargetSpecies.ID"
    Expression ="TargetSpecies.TargetList"
    Expression ="TargetSpecies.TSN"
    Expression ="TargetSpecies.EstablishDate"
    Expression ="TargetSpecies.RetireDate"
    Expression ="tlu_Plants.PLANTS_Code"
    Expression ="tlu_Plants.Latin_Name"
End
Begin Joins
    LeftTable ="TargetSpecies"
    RightTable ="tlu_Plants"
    Expression ="TargetSpecies.ID = tlu_Plants.ID"
    Flag =1
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.PLANTS_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetSpecies.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetSpecies.TargetList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetSpecies.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetSpecies.EstablishDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TargetSpecies.RetireDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-112
    Top =114
    Right =399
    Bottom =505
    Left =-1
    Top =-1
    Right =487
    Bottom =164
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="TargetSpecies"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
