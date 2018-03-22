Operation =1
Option =0
Where ="(((tbl_Tree_Conditions.Condition)=\"Vines in the crown\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
    Name ="tbl_Tree_Conditions"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Tree_Data.Tree_Data_ID"
    Expression ="tbl_Tree_Conditions.Tree_Condition_ID"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Tags.Tag"
    Alias ="Host_Latin_Name"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tree_Conditions.Condition"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID=tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID=tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_Conditions"
    Expression ="tbl_Tree_Data.Tree_Data_ID=tbl_Tree_Conditions.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN=tlu_Plants.TSN"
    Flag =1
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
        dbText "Name" ="tbl_Locations.Location_ID"
        dbInteger "ColumnWidth" ="1215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Conditions.Tree_Condition_ID"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Conditions.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Latin_Name"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =32
    Right =1325
    Bottom =850
    Left =-1
    Top =-1
    Right =1257
    Bottom =522
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Tree_Conditions"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =414
        Bottom =238
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =203
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =45
        Top =216
        Right =189
        Bottom =466
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =280
        Top =317
        Right =424
        Bottom =461
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
