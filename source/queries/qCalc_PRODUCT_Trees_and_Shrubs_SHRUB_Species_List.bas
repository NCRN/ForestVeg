Operation =1
Option =0
Having ="(((qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.Habit)=\"Shrub\"))"
Begin InputTables
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.TSN"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.Habit"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.Habit"
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
        dbText "Name" ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.Habit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1478
    Bottom =997
    Left =-1
    Top =-1
    Right =1332
    Bottom =468
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =269
        Bottom =258
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount"
        Name =""
    End
End
