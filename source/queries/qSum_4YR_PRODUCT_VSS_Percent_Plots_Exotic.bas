Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_VSS_Exotic_Count_By_Plot"
    Name ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code"
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Plot_Count"
    Alias ="Ex_Plot_Cnt"
    Expression ="Sum(IIf([ExoticYN_Bin]>0,1,0))"
    Alias ="Ex_Plot_Percent"
    Expression ="[Ex_Plot_Cnt]*100/[Plot_Count]"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
    RightTable ="qCalc_PRODUCT_VSS_Exotic_Count_By_Plot"
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code = qCalc_PRODUCT_VSS_Exo"
        "tic_Count_By_Plot.[Admin Unit Code]"
    Flag =2
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Plot_Count"
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
        dbText "Name" ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Plot_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ex_Plot_Cnt"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ex_Plot_Percent"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1518
    Bottom =947
    Left =-1
    Top =-1
    Right =1448
    Bottom =578
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =428
        Top =20
        Right =572
        Bottom =164
        Top =0
        Name ="qCalc_PRODUCT_VSS_Exotic_Count_By_Plot"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
        Name =""
    End
End
