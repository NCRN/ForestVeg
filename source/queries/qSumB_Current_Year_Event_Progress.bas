Operation =1
Option =0
Where ="(((tbl_Locations.Location_Status)=\"New\") AND ((tbl_Locations.Panel)=[Forms]![f"
    "rm_Switchboard]![Panel])) OR (((tbl_Locations.Location_Status)=\"Active\") AND ("
    "(tbl_Locations.Panel)=[Forms]![frm_Switchboard]![Panel]))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="qSumB_Current_Year_Events_Completed"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="qSumB_Current_Year_Events_Completed.Event_Date"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="qSumB_Current_Year_Events_Completed"
    Expression ="tbl_Locations.Location_ID = qSumB_Current_Year_Events_Completed.Location_ID"
    Flag =2
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
dbText "Description" ="Which plots were sampled in the current year?"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSumB_Current_Year_Events_Completed.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1546
    Bottom =980
    Left =-1
    Top =-1
    Right =1438
    Bottom =565
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =28
        Top =9
        Right =232
        Bottom =451
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =280
        Top =12
        Right =482
        Bottom =308
        Top =0
        Name ="qSumB_Current_Year_Events_Completed"
        Name =""
    End
End
