Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
    Name ="qCalc_PRODUCT_Seedlings_Prequery"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code"
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Plot_Count"
    Alias ="Seedling_Count"
    Expression ="Sum(qCalc_PRODUCT_Seedlings_Prequery.Samp_Count)"
    Alias ="Seedling_per_ha"
    Expression ="Round([Seedling_Count]/([Plot_Count]*0.0012),2)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
    RightTable ="qCalc_PRODUCT_Seedlings_Prequery"
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code = qCalc_PRODUCT_Seedlin"
        "gs_Prequery.Admin_Unit_Code"
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
        dbText "Name" ="Seedling_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Seedling_per_ha"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =29
    Top =72
    Right =1509
    Bottom =933
    Left =-1
    Top =-1
    Right =1448
    Bottom =578
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qCalc_PRODUCT_Seedlings_Prequery"
        Name =""
    End
End
