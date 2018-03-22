Operation =1
Option =0
Begin InputTables
    Name ="qActive_Trees_Shrubs_Herbs_Vines"
End
Begin OutputColumns
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Location_ID"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
    Alias ="Date"
    Expression ="Format([Event_Date],\"mm/dd/yyyy\")"
    Alias ="Habit_Class"
    Expression ="[Habit] & \" / \" & [Class]"
    Alias ="Occurence_Count"
    Expression ="Count(qActive_Trees_Shrubs_Herbs_Vines.Habit)"
End
Begin OrderBy
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
    Flag =0
    Expression ="Format([Event_Date],\"mm/dd/yyyy\")"
    Flag =0
    Expression ="[Habit] & \" / \" & [Class]"
    Flag =0
End
Begin Groups
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN"
    GroupLevel =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Location_ID"
    GroupLevel =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
    GroupLevel =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
    GroupLevel =0
    Expression ="Format([Event_Date],\"mm/dd/yyyy\")"
    GroupLevel =0
    Expression ="[Habit] & \" / \" & [Class]"
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
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Location_ID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.TSN"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occurence_Count"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =30
    Top =67
    Right =1472
    Bottom =927
    Left =-1
    Top =-1
    Right =1410
    Bottom =577
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =285
        Bottom =218
        Top =0
        Name ="qActive_Trees_Shrubs_Herbs_Vines"
        Name =""
    End
End
