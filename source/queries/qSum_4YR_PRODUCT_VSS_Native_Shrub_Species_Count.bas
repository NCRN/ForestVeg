﻿Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_VSS_Shrub_Species_By_AdmUnit_Prequery"
    Name ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code"
    Alias ="CountOfLatin Name"
    Expression ="Count(qCalc_PRODUCT_VSS_Shrub_Species_By_AdmUnit_Prequery.[Latin Name])"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
    RightTable ="qCalc_PRODUCT_VSS_Shrub_Species_By_AdmUnit_Prequery"
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code = qCalc_PRODUCT_VSS_Shr"
        "ub_Species_By_AdmUnit_Prequery.[Admin Unit Code]"
    Flag =2
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit.Admin_Unit_Code"
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
        dbText "Name" ="CountOfLatin Name"
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
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =434
        Top =32
        Right =578
        Bottom =176
        Top =0
        Name ="qCalc_PRODUCT_VSS_Shrub_Species_By_AdmUnit_Prequery"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qSum_4YR_PRODUCT_Plot_Count_by_AdminUnit"
        Name =""
    End
End
