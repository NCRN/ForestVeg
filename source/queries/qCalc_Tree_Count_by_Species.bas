Operation =1
Option =0
Begin InputTables
    Name ="tlu_Plants"
    Name ="qActive_Tree_Data"
    Name ="qCalc_Basal_Area_per_Tree"
    Name ="qFiltered_Events"
    Name ="qFiltered_Locations"
End
Begin OutputColumns
    Expression ="tlu_Plants.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.PLANTS_Common"
    Alias ="Tree_Count"
    Expression ="Count(qCalc_Basal_Area_per_Tree.Tree_Data_ID)"
    Alias ="Tree_SumBasalArea_cm2"
    Expression ="Sum(Round([SumBasalArea_cm2]))"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qActive_Tree_Data"
    Expression ="tlu_Plants.TSN=qActive_Tree_Data.TSN"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="qCalc_Basal_Area_per_Tree"
    Expression ="qActive_Tree_Data.Tree_Data_ID=qCalc_Basal_Area_per_Tree.Tree_Data_ID"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="qFiltered_Events"
    Expression ="qActive_Tree_Data.Event_ID=qFiltered_Events.Event_ID"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="qFiltered_Locations"
    Expression ="qActive_Tree_Data.Location_ID=qFiltered_Locations.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tlu_Plants.Latin_Name"
    Flag =0
End
Begin Groups
    Expression ="tlu_Plants.TSN"
    GroupLevel =0
    Expression ="tlu_Plants.Latin_Name"
    GroupLevel =0
    Expression ="tlu_Plants.PLANTS_Common"
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
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.PLANTS_Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =34
    Top =63
    Right =1134
    Bottom =753
    Left =-1
    Top =-1
    Right =1501
    Bottom =352
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =19
        Top =18
        Right =163
        Bottom =335
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =240
        Top =5
        Right =384
        Bottom =340
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
    Begin
        Left =474
        Top =9
        Right =618
        Bottom =153
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
    Begin
        Left =489
        Top =198
        Right =633
        Bottom =342
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =706
        Top =198
        Right =850
        Bottom =342
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
End
