Operation =1
Option =0
Begin InputTables
    Name ="qActive_Sapling_Data"
    Name ="tbl_Sapling_DBH"
End
Begin OutputColumns
    Expression ="tbl_Sapling_DBH.Sapling_Data_ID"
    Expression ="qActive_Sapling_Data.Event_ID"
    Alias ="FirstOfTag_ID"
    Expression ="First(qActive_Sapling_Data.Tag_ID)"
    Expression ="qActive_Sapling_Data.Exotic"
    Alias ="Stems"
    Expression ="Count(tbl_Sapling_DBH.DBH)"
    Alias ="SumBasalArea_cm2"
    Expression ="Round(Sum(3.1415926*(([DBH]/2)^2)),1)"
    Alias ="Equiv_DBH_cm"
    Expression ="Round((([SumBasalArea_cm2]/3.1415)^0.5)*2,1)"
End
Begin Joins
    LeftTable ="qActive_Sapling_Data"
    RightTable ="tbl_Sapling_DBH"
    Expression ="qActive_Sapling_Data.Sapling_Data_ID = tbl_Sapling_DBH.Sapling_Data_ID"
    Flag =1
End
Begin Groups
    Expression ="tbl_Sapling_DBH.Sapling_Data_ID"
    GroupLevel =0
    Expression ="qActive_Sapling_Data.Event_ID"
    GroupLevel =0
    Expression ="qActive_Sapling_Data.Exotic"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qCalc_Basal_Area_per_Sapling].[SumBasalArea_cm2] DESC"
Begin
    Begin
        dbText "Name" ="tbl_Sapling_DBH.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="FirstOfTag_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4200"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qActive_Sapling_Data].Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =101
    Right =1065
    Bottom =584
    Left =-1
    Top =-1
    Right =1005
    Bottom =119
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qActive_Sapling_Data"
        Name =""
    End
    Begin
        Left =306
        Top =22
        Right =450
        Bottom =166
        Top =0
        Name ="tbl_Sapling_DBH"
        Name =""
    End
End
