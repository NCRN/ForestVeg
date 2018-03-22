Operation =1
Option =0
Having ="(((tbl_Tags_History.Field_Name)=\"Tag\") AND ((tbl_Tags_History.Value_New)<>Form"
    "at([tbl_Tags].[Tag])))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Tags_History"
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_Name"
    Alias ="LastOfEvent_Date"
    Expression ="Last(tbl_Events.Event_Date)"
    Expression ="tbl_Tags_History.Change_Date"
    Expression ="tbl_Tags_History.Tags_History_ID"
    Expression ="tbl_Tags_History.Table_Name"
    Expression ="tbl_Tags_History.Record_ID_Field_Name"
    Expression ="tbl_Tags_History.Record_ID"
    Expression ="tbl_Tags_History.Field_Name"
    Expression ="tbl_Tags_History.Value_New"
    Expression ="tbl_Tags_History.Value_Old"
    Expression ="tbl_Tags_History.Value_History_Notes"
    Expression ="tbl_Tags.Tag"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tags_History"
    Expression ="tbl_Tags.Tag_ID = tbl_Tags_History.Record_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =0
    Expression ="tbl_Tags_History.Change_Date"
    GroupLevel =0
    Expression ="tbl_Tags_History.Tags_History_ID"
    GroupLevel =0
    Expression ="tbl_Tags_History.Table_Name"
    GroupLevel =0
    Expression ="tbl_Tags_History.Record_ID_Field_Name"
    GroupLevel =0
    Expression ="tbl_Tags_History.Record_ID"
    GroupLevel =0
    Expression ="tbl_Tags_History.Field_Name"
    GroupLevel =0
    Expression ="tbl_Tags_History.Value_New"
    GroupLevel =0
    Expression ="tbl_Tags_History.Value_Old"
    GroupLevel =0
    Expression ="tbl_Tags_History.Value_History_Notes"
    GroupLevel =0
    Expression ="tbl_Tags.Tag"
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
        dbText "Name" ="tbl_Tags_History.Value_History_Notes"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Tags_History_ID"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Table_Name"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Record_ID_Field_Name"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Record_ID"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Field_Name"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_New"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_Old"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfEvent_Date"
        dbInteger "ColumnWidth" ="2280"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Change_Date"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfEvent_Date"
        dbInteger "ColumnWidth" ="2280"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-61
    Top =661
    Right =1529
    Bottom =1572
    Left =-1
    Top =-1
    Right =1558
    Bottom =611
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =786
        Top =31
        Right =1013
        Bottom =342
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =1148
        Top =14
        Right =1369
        Bottom =323
        Top =0
        Name ="tbl_Tags_History"
        Name =""
    End
    Begin
        Left =54
        Top =15
        Right =239
        Bottom =497
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =375
        Top =111
        Right =653
        Bottom =533
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
