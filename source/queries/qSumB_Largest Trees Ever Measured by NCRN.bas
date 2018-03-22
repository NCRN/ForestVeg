Operation =1
Option =0
Begin InputTables
    Name ="qActive_Tree_Data"
    Name ="tlu_Plants"
    Name ="qCalc_Basal_Area_per_Tree"
End
Begin OutputColumns
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.Common"
    Expression ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
    Expression ="qCalc_Basal_Area_per_Tree.Equiv_DBH_cm"
    Expression ="qActive_Tree_Data.Plot_Name"
    Expression ="qActive_Tree_Data.StemList"
    Expression ="qActive_Tree_Data.Tag"
    Expression ="qActive_Tree_Data.ConditionAndPest_List"
End
Begin Joins
    LeftTable ="qActive_Tree_Data"
    RightTable ="qCalc_Basal_Area_per_Tree"
    Expression ="qActive_Tree_Data.Tree_Data_ID = qCalc_Basal_Area_per_Tree.Tree_Data_ID"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="tlu_Plants"
    Expression ="qActive_Tree_Data.TSN = tlu_Plants.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
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
dbText "Description" ="What are the largest trees ever measured during NCRN monitoring?"
Begin
    Begin
        dbText "Name" ="qActive_Tree_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.StemList"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.ConditionAndPest_List"
        dbInteger "ColumnWidth" ="3900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.Equiv_DBH_cm"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =72
    Right =950
    Bottom =558
    Left =-1
    Top =-1
    Right =910
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =365
        Bottom =471
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
    Begin
        Left =415
        Top =199
        Right =577
        Bottom =522
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =605
        Top =12
        Right =825
        Bottom =275
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
End
