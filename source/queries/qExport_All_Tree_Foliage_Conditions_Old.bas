Operation =1
Option =0
Begin InputTables
    Name ="qActive_Tree_Foliage_Conditions"
End
Begin OutputColumns
    Expression ="qActive_Tree_Foliage_Conditions.Plot_Name"
    Expression ="qActive_Tree_Foliage_Conditions.Unit_Code"
    Expression ="qActive_Tree_Foliage_Conditions.Admin_Unit_Code"
    Expression ="qActive_Tree_Foliage_Conditions.Panel"
    Expression ="qActive_Tree_Foliage_Conditions.Frame"
    Expression ="qActive_Tree_Foliage_Conditions.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Expression ="qActive_Tree_Foliage_Conditions.Tag"
    Expression ="qActive_Tree_Foliage_Conditions.Condition"
    Expression ="qActive_Tree_Foliage_Conditions.Condition_Description"
    Expression ="qActive_Tree_Foliage_Conditions.Percent_Afflicted"
    Expression ="qActive_Tree_Foliage_Conditions.TSN"
    Expression ="qActive_Tree_Foliage_Conditions.Latin_Name"
    Expression ="qActive_Tree_Foliage_Conditions.Crown_Class"
    Expression ="qActive_Tree_Foliage_Conditions.Tree_Status"
End
Begin OrderBy
    Expression ="qActive_Tree_Foliage_Conditions.Plot_Name"
    Flag =0
    Expression ="qActive_Tree_Foliage_Conditions.Sample_Year"
    Flag =0
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Flag =0
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
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Condition_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Percent_Afflicted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Foliage_Conditions.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =25
    Top =52
    Right =1540
    Bottom =859
    Left =-1
    Top =-1
    Right =1483
    Bottom =538
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =281
        Bottom =371
        Top =0
        Name ="qActive_Tree_Foliage_Conditions"
        Name =""
    End
End
