Operation =1
Option =0
Begin InputTables
    Name ="qActive_Tree_Conditions"
End
Begin OutputColumns
    Expression ="qActive_Tree_Conditions.Plot_Name"
    Expression ="qActive_Tree_Conditions.Unit_Code"
    Expression ="qActive_Tree_Conditions.Unit_Group"
    Expression ="qActive_Tree_Conditions.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="qActive_Tree_Conditions.Panel"
    Expression ="qActive_Tree_Conditions.Frame"
    Expression ="qActive_Tree_Conditions.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([Event_Date],\"yyyymmdd\"))"
    Expression ="qActive_Tree_Conditions.Tag"
    Expression ="qActive_Tree_Conditions.TSN"
    Expression ="qActive_Tree_Conditions.Latin_Name"
    Expression ="qActive_Tree_Conditions.Crown_Class"
    Alias ="Status"
    Expression ="qActive_Tree_Conditions.Tree_Status"
    Expression ="qActive_Tree_Conditions.Condition"
    Expression ="qActive_Tree_Conditions.Pest"
End
Begin OrderBy
    Expression ="qActive_Tree_Conditions.Plot_Name"
    Flag =0
    Expression ="qActive_Tree_Conditions.Sample_Year"
    Flag =0
    Expression ="qActive_Tree_Conditions.Tag"
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
        dbText "Name" ="qActive_Tree_Conditions.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Conditions.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =20
    Right =1536
    Bottom =490
    Left =-1
    Top =-1
    Right =1483
    Bottom =242
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =342
        Bottom =437
        Top =0
        Name ="qActive_Tree_Conditions"
        Name =""
    End
End
