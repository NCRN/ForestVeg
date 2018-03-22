Operation =1
Option =0
Begin InputTables
    Name ="tbl_Tree_DBH"
    Name ="tbl_Tree_Data"
    Name ="qEnumCrownClass"
End
Begin OutputColumns
    Expression ="tbl_Tree_Data.Tree_Data_ID"
    Expression ="tbl_Tree_Data.Event_ID"
    Alias ="CrownClass"
    Expression ="qEnumCrownClass.Enum_Description"
    Alias ="Stems"
    Expression ="Count(tbl_Tree_DBH.DBH)"
    Alias ="StemsLive"
    Expression ="Sum(IIf([Live]=True,1,0))"
    Alias ="StemsDead"
    Expression ="Sum(IIf([Live]=False,1,0))"
    Alias ="SumBasalArea_cm2"
    Expression ="Round(Sum(3.1415926*(([DBH]/2)^2)),1)"
    Alias ="FirstOfTag_ID"
    Expression ="First(tbl_Tree_Data.Tag_ID)"
    Alias ="Equiv_DBH_cm"
    Expression ="Round((([SumBasalArea_cm2]/3.1415)^0.5)*2,1)"
    Alias ="SumLiveBasalArea_cm2"
    Expression ="Round(Sum(3.1415926*(((IIf([Live]=True,[DBH],0))/2)^2)),1)"
    Alias ="SumDeadBasalArea_cm2"
    Expression ="Round(Sum(3.1415926*(((IIf([Live]=False,[DBH],0))/2)^2)),1)"
    Alias ="Equiv_Live_DBH_cm"
    Expression ="Round((([SumLiveBasalArea_cm2]/3.1415)^0.5)*2,1)"
    Alias ="Equiv_Dead_DBH_cm"
    Expression ="Round((([SumDeadBasalArea_cm2]/3.1415)^0.5)*2,1)"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_DBH"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="qEnumCrownClass"
    Expression ="tbl_Tree_Data.Crown_Class = qEnumCrownClass.CrownClassCode"
    Flag =1
End
Begin Groups
    Expression ="tbl_Tree_Data.Tree_Data_ID"
    GroupLevel =0
    Expression ="tbl_Tree_Data.Event_ID"
    GroupLevel =0
    Expression ="qEnumCrownClass.Enum_Description"
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
        dbText "Name" ="Stems"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="FirstOfTag_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1245"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Event_ID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StemsDead"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="StemsLive"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="CrownClass"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =41
    Top =154
    Right =1360
    Bottom =918
    Left =-1
    Top =-1
    Right =1287
    Bottom =301
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =375
        Top =21
        Right =519
        Bottom =165
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
    Begin
        Left =24
        Top =16
        Right =309
        Bottom =228
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =610
        Top =12
        Right =754
        Bottom =156
        Top =0
        Name ="qEnumCrownClass"
        Name =""
    End
End
