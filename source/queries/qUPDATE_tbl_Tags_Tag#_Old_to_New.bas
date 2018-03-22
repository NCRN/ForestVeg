Operation =4
Option =0
Where ="(((tbl_Tags_History.Field_Name)=\"Tag\") AND ((tbl_Tags_History.Value_New)<>Form"
    "at([tbl_Tags].[Tag])))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="tbl_Tags_History"
    Name ="tbl_Locations"
End
Begin OutputColumns
    Name ="tbl_Tags.Tag"
    Expression ="[tbl_Tags_History].[Value_New]"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tags_History"
    Expression ="tbl_Tags.Tag_ID = tbl_Tags_History.Record_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Tags_History.Value_History_Notes"
        dbInteger "ColumnWidth" ="4650"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Tags_History_ID"
        dbInteger "ColumnWidth" ="3525"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="2"
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
        dbText "Name" ="tbl_Tags_History.Table_Name"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Record_ID_Field_Name"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Record_ID"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Field_Name"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_New"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_Old"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =14
    Top =3
    Right =1604
    Bottom =914
    Left =-1
    Top =-1
    Right =1558
    Bottom =560
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =127
        Top =28
        Right =354
        Bottom =339
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =385
        Top =172
        Right =606
        Bottom =481
        Top =0
        Name ="tbl_Tags_History"
        Name =""
    End
    Begin
        Left =809
        Top =137
        Right =1090
        Bottom =490
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
