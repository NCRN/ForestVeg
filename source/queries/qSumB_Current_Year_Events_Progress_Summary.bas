Operation =1
Option =0
Begin InputTables
    Name ="qSumB_Current_Year_Event_Progress"
End
Begin OutputColumns
    Expression ="qSumB_Current_Year_Event_Progress.Unit_Code"
    Alias ="Pending"
    Expression ="Sum(IIf(IsNull([Event_Date]),1,0))"
    Alias ="Completed"
    Expression ="Sum(IIf(IsNull([Event_Date]),0,1))"
End
Begin Groups
    Expression ="qSumB_Current_Year_Event_Progress.Unit_Code"
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
dbText "Description" ="How many events are completed and pending for each park for the current year?"
Begin
    Begin
        dbText "Name" ="qSumB_Current_Year_Event_Progress.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Completed"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pending"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1546
    Bottom =980
    Left =-1
    Top =-1
    Right =1400
    Bottom =519
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qSumB_Current_Year_Event_Progress"
        Name =""
    End
End
