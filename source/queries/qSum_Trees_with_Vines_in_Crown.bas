Operation =1
Option =0
Having ="(((qActive_Tree_Conditions.Condition)=\"Vines in the crown\"))"
Begin InputTables
    Name ="qActive_Tree_Conditions"
End
Begin OutputColumns
    Expression ="qActive_Tree_Conditions.Tree_Data_ID"
    Expression ="qActive_Tree_Conditions.Plot_Name"
    Expression ="qActive_Tree_Conditions.Panel"
    Expression ="qActive_Tree_Conditions.Frame"
    Expression ="qActive_Tree_Conditions.Sample_Year"
    Expression ="qActive_Tree_Conditions.Tag"
    Expression ="qActive_Tree_Conditions.Condition"
    Alias ="CountOfTSN"
    Expression ="Count(qActive_Tree_Conditions.TSN)"
End
Begin Groups
    Expression ="qActive_Tree_Conditions.Tree_Data_ID"
    GroupLevel =0
    Expression ="qActive_Tree_Conditions.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Tree_Conditions.Panel"
    GroupLevel =0
    Expression ="qActive_Tree_Conditions.Frame"
    GroupLevel =0
    Expression ="qActive_Tree_Conditions.Sample_Year"
    GroupLevel =0
    Expression ="qActive_Tree_Conditions.Tag"
    GroupLevel =0
    Expression ="qActive_Tree_Conditions.Condition"
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
        dbText "Name" ="qActive_Tree_Conditions.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Tree_Data_ID"
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
        Left =48
        Top =12
        Right =268
        Bottom =425
        Top =0
        Name ="qActive_Tree_Conditions"
        Name =""
    End
End
