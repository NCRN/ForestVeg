Operation =1
Option =0
Where ="(((Year([Event_Date]))=2015) AND ((tbl_Events.Event_Date) Is Not Null) AND ((tbl"
    "_Tags.Tag) Is Null) AND ((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\")) OR ("
    "((Year([Event_Date]))=2015) AND ((tbl_Events.Event_Date) Is Not Null) AND ((tbl_"
    "Tags.TSN) Is Null) AND ((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\")) OR (("
    "(Year([Event_Date]))=2015) AND ((tbl_Sapling_Data.Sapling_Status) Is Null)) OR ("
    "((Year([Event_Date]))=2015) AND ((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\""
    ") AND ((tbl_Sapling_Data.Browsed) Is Not Null)) OR (((Year([Event_Date]))=2015) "
    "AND ((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\") AND ((tbl_Sapling_Data.Br"
    "owsable) Is Not Null)) OR (((Year([Event_Date]))=2015) AND ((tbl_Sapling_Data.Sa"
    "pling_Status) Like \"Dead*\") AND ((tbl_Tags.Microplot_Number) Is Null)) OR (((Y"
    "ear([Event_Date]))=2015) AND ((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\") "
    "AND ((tbl_Sapling_Data.Vines_Checked)=True)) OR (((Year([Event_Date]))=2015) AND"
    " ((tbl_Sapling_Data.Sapling_Status) Like \"Dead*\") AND ((tbl_Sapling_Data.Condi"
    "tions_Checked)=True)) OR (((Year([Event_Date]))=2015) AND ((tbl_Sapling_Data.Sap"
    "ling_Status) Like \"Dead*\") AND ((tbl_Sapling_Data.Foliage_Conditions_Checked)="
    "True))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tbl_Sapling_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Alias ="Year"
    Expression ="Year([Event_Date])"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Sapling_Data.Browsed"
    Expression ="tbl_Sapling_Data.Browsable"
    Expression ="tbl_Tags.Microplot_Number"
    Expression ="tbl_Sapling_Data.Vines_Checked"
    Expression ="tbl_Sapling_Data.Conditions_Checked"
    Expression ="tbl_Sapling_Data.Foliage_Conditions_Checked"
    Expression ="tbl_Sapling_Data.Sapling_Notes"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Events.Event_Date"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbText "Description" ="Sapling sampling record is incomplete"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1440"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Notes"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsable"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsed"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =-7
    Top =18
    Right =1755
    Bottom =1039
    Left =-1
    Top =-1
    Right =1512
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =561
        Top =8
        Right =768
        Bottom =122
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =7
        Top =8
        Right =103
        Bottom =122
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =326
        Top =-5
        Right =470
        Bottom =271
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =142
        Top =33
        Right =334
        Bottom =349
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =545
        Top =133
        Right =689
        Bottom =277
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
