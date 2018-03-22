Operation =1
Option =0
Having ="(((Year([event_date]))=2012))"
Begin InputTables
    Name ="qCalc_Basal_Area_per_Tree"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Events.Location_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Events.Event_Date"
    Alias ="Sample_Year"
    Expression ="Year([event_date])"
    Alias ="Tree_Count_2012"
    Expression ="Count(qCalc_Basal_Area_per_Tree.Tree_Data_ID)"
    Alias ="Tree_Stem_Count_2012"
    Expression ="CInt(Nz(Sum([Stems])))"
    Alias ="Tree_BasalArea_cm2_Sum_2012"
    Expression ="CLng(Nz(Sum([SumBasalArea_cm2])))"
End
Begin Joins
    LeftTable ="qCalc_Basal_Area_per_Tree"
    RightTable ="tbl_Events"
    Expression ="qCalc_Basal_Area_per_Tree.Event_ID = tbl_Events.Event_ID"
    Flag =3
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
        dbText "Name" ="Tree_Count_2012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_Stem_Count_2012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_BasalArea_cm2_Sum_2012"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =164
    Top =197
    Right =1293
    Bottom =772
    Left =-1
    Top =-1
    Right =1097
    Bottom =217
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =268
        Top =16
        Right =461
        Bottom =177
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
    Begin
        Left =29
        Top =20
        Right =173
        Bottom =164
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
