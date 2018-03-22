Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Conditions_prequery"
    Name ="tlu_Tree_Condition"
End
Begin OutputColumns
    Alias ="Category"
    Expression ="IIf([Pest]=True,\"Pest\",\"Condition\")"
    Expression ="tlu_Tree_Condition.Description"
    Alias ="Plot Count"
    Expression ="Count(qCalc_PRODUCT_Conditions_prequery.Plot_Name)"
    Alias ="Total Occurences"
    Expression ="CInt(Nz(Sum([CountOfTree_Condition_ID]),0))"
End
Begin Joins
    LeftTable ="tlu_Tree_Condition"
    RightTable ="qCalc_PRODUCT_Conditions_prequery"
    Expression ="tlu_Tree_Condition.Description = qCalc_PRODUCT_Conditions_prequery.Condition"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([Pest]=True,\"Pest\",\"Condition\")"
    Flag =0
    Expression ="tlu_Tree_Condition.Description"
    Flag =0
    Expression ="Count(qCalc_PRODUCT_Conditions_prequery.Plot_Name)"
    Flag =0
End
Begin Groups
    Expression ="IIf([Pest]=True,\"Pest\",\"Condition\")"
    GroupLevel =0
    Expression ="tlu_Tree_Condition.Description"
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
        dbText "Name" ="Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Tree_Condition.Description"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot Count"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Total Occurences"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =266
        Top =30
        Right =604
        Bottom =174
        Top =0
        Name ="qCalc_PRODUCT_Conditions_prequery"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tlu_Tree_Condition"
        Name =""
    End
End
