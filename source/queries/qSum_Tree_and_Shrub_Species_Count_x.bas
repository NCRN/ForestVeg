Operation =6
Option =0
Begin InputTables
    Name ="qCalc_Active_Tree_Shrub-Species_Count"
End
Begin OutputColumns
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].Habit"
    GroupLevel =2
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].Sample_Year"
    GroupLevel =1
    Alias ="CountOfTSN"
    Expression ="Count([qCalc_Active_Tree_Shrub-Species_Count].TSN)"
End
Begin Groups
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].Habit"
    GroupLevel =2
    Expression ="[qCalc_Active_Tree_Shrub-Species_Count].Sample_Year"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbText "Description" ="Union of all tree and shrub records including species across size class. Expecte"
    "d to be used for specimen and species counts. Not appropriate for density calcul"
    "ation without correcting for the different sample areas of Trees, saplings and s"
    "eedlings."
Begin
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2007"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2009"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfTSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qCalc_Active_Tree_Shrub-Species_Count].Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qCalc_Active_Tree_Shrub-Species_Count].Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2011"
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
    ColumnsShown =559
    Begin
        Left =52
        Top =14
        Right =196
        Bottom =158
        Top =0
        Name ="qCalc_Active_Tree_Shrub-Species_Count"
        Name =""
    End
End
