Operation =1
Option =0
Where ="(((tbl_Tasks.Task_Status)=\"Active\"))"
Begin InputTables
    Name ="tbl_Tasks"
    Name ="tbl_Locations"
    Name ="tlu_Contacts"
End
Begin OutputColumns
    Expression ="tbl_Tasks.Task_Date"
    Expression ="tbl_Locations.Plot_Name"
    Alias ="Name"
    Expression ="[First_Name] & \" \" & [Last_Name]"
    Expression ="tbl_Tasks.Task_Status"
    Expression ="tbl_Tasks.Task_Notes"
    Expression ="tbl_Tasks.Followup_Date"
    Expression ="tbl_Tasks.Followup_Notes"
End
Begin Joins
    LeftTable ="tbl_Tasks"
    RightTable ="tbl_Locations"
    Expression ="tbl_Tasks.Location_ID=tbl_Locations.Location_ID"
    Flag =2
    LeftTable ="tbl_Tasks"
    RightTable ="tlu_Contacts"
    Expression ="tbl_Tasks.Task_Contact_ID=tlu_Contacts.Contact_ID"
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
Begin
    Begin
        dbText "Name" ="tbl_Tasks.Task_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tasks.Task_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tasks.Task_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tasks.Followup_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tasks.Followup_Notes"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1518
    Bottom =971
    Left =-1
    Top =-1
    Right =1494
    Bottom =652
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =108
        Top =6
        Right =314
        Bottom =264
        Top =0
        Name ="tbl_Tasks"
        Name =""
    End
    Begin
        Left =855
        Top =-9
        Right =999
        Bottom =135
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =628
        Top =239
        Right =772
        Bottom =383
        Top =0
        Name ="tlu_Contacts"
        Name =""
    End
End
