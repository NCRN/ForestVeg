Operation =1
Option =0
Having ="(((Year([event_date]))=2006))"
Begin InputTables
    Name ="tbl_Events"
    Name ="qCalc_Basal_Area_per_Sapling"
End
Begin OutputColumns
    Expression ="tbl_Events.Location_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Events.Event_Date"
    Alias ="Sample_Year"
    Expression ="Year([event_date])"
    Alias ="Sapling_Count_2006"
    Expression ="Count(qCalc_Basal_Area_per_Sapling.Sapling_Data_ID)"
    Alias ="Sapling_Stem_Count_2006"
    Expression ="CInt(nz(Sum([Stems])))"
    Alias ="Sapling_BasalArea_cm2_Sum_2006"
    Expression ="CLng(nz(Sum([SumBasalArea_cm2])))"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="qCalc_Basal_Area_per_Sapling"
    Expression ="tbl_Events.Event_ID=qCalc_Basal_Area_per_Sapling.Event_ID"
    Flag =2
End
Begin Groups
    Expression ="tbl_Events.Location_ID"
    GroupLevel =0
    Expression ="tbl_Events.Event_ID"
    GroupLevel =0
    Expression ="tbl_Events.Event_Date"
    GroupLevel =0
    Expression ="Year([event_date])"
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
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sapling_Count_2006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sapling_Stem_Count_2006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sapling_BasalArea_cm2_Sum_2006"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =108
    Top =188
    Right =1237
    Bottom =763
    Left =-1
    Top =-1
    Right =1105
    Bottom =166
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =11
        Top =9
        Right =155
        Bottom =153
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =208
        Top =9
        Right =563
        Bottom =153
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
End
