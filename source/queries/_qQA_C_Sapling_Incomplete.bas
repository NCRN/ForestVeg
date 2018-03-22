Operation =1
Option =0
Where ="(((tbl_Sapling_Data.Sapling_Status) Is Null)) OR (((tbl_Sapling_Data.Sapling_Sta"
    "tus)<>\"Removed from study\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead\" And"
    " (tbl_Sapling_Data.Sapling_Status)<>\"Dead Fallen\" And (tbl_Sapling_Data.Saplin"
    "g_Status)<>\"Dead Standing\" And (tbl_Sapling_Data.Sapling_Status)<>\"Downgraded"
    " to Non-Sampled\" And (tbl_Sapling_Data.Sapling_Status)<>\"Missing\") AND ((tbl_"
    "Sapling_Data.Browsed) Is Null)) OR (((tbl_Sapling_Data.Sapling_Status)<>\"Remove"
    "d from study\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead\" And (tbl_Sapling_"
    "Data.Sapling_Status)<>\"Dead Fallen\" And (tbl_Sapling_Data.Sapling_Status)<>\"D"
    "ead Standing\" And (tbl_Sapling_Data.Sapling_Status)<>\"Downgraded to Non-Sample"
    "d\" And (tbl_Sapling_Data.Sapling_Status)<>\"Missing\") AND ((tbl_Sapling_Data.B"
    "rowsable) Is Null)) OR (((tbl_Tags.TSN) Is Null) AND ((tbl_Sapling_Data.Sapling_"
    "Status)<>\"Removed from study\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead\" "
    "And (tbl_Sapling_Data.Sapling_Status)<>\"Dead Fallen\" And (tbl_Sapling_Data.Sap"
    "ling_Status)<>\"Dead Standing\" And (tbl_Sapling_Data.Sapling_Status)<>\"Downgra"
    "ded to Non-Sampled\" And (tbl_Sapling_Data.Sapling_Status)<>\"Missing\")) OR ((("
    "tbl_Sapling_Data.Sapling_Status)<>\"Removed from study\" And (tbl_Sapling_Data.S"
    "apling_Status)<>\"Dead\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead Fallen\" "
    "And (tbl_Sapling_Data.Sapling_Status)<>\"Dead Standing\" And (tbl_Sapling_Data.S"
    "apling_Status)<>\"Downgraded to Non-Sampled\" And (tbl_Sapling_Data.Sapling_Stat"
    "us)<>\"Missing\") AND ((tbl_Sapling_Data.Habit) Is Null)) OR (((tbl_Tags.Tag) Is"
    " Null) AND ((tbl_Sapling_Data.Sapling_Status)<>\"Removed from study\" And (tbl_S"
    "apling_Data.Sapling_Status)<>\"Dead\" And (tbl_Sapling_Data.Sapling_Status)<>\"D"
    "ead Fallen\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead Standing\" And (tbl_S"
    "apling_Data.Sapling_Status)<>\"Downgraded to Non-Sampled\" And (tbl_Sapling_Data"
    ".Sapling_Status)<>\"Missing\")) OR (((tbl_Sapling_Data.Sapling_Status)<>\"Remove"
    "d from study\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead\" And (tbl_Sapling_"
    "Data.Sapling_Status)<>\"Dead Fallen\" And (tbl_Sapling_Data.Sapling_Status)<>\"D"
    "ead Standing\" And (tbl_Sapling_Data.Sapling_Status)<>\"Downgraded to Non-Sample"
    "d\" And (tbl_Sapling_Data.Sapling_Status)<>\"Missing\") AND ((tbl_Tags.Microplot"
    "_Number) Is Null)) OR (((tbl_Sapling_Data.Sapling_Status)<>\"Removed from study\""
    " And (tbl_Sapling_Data.Sapling_Status)<>\"Dead\" And (tbl_Sapling_Data.Sapling_S"
    "tatus)<>\"Dead Fallen\" And (tbl_Sapling_Data.Sapling_Status)<>\"Dead Standing\""
    " And (tbl_Sapling_Data.Sapling_Status)<>\"Downgraded to Non-Sampled\" And (tbl_S"
    "apling_Data.Sapling_Status)<>\"Missing\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tbl_Sapling_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Sapling_Data.Sapling_Status"
    Expression ="tbl_Sapling_Data.Browsed"
    Expression ="tbl_Sapling_Data.Browsable"
    Expression ="tbl_Sapling_Data.Habit"
    Expression ="tbl_Sapling_Data.DRC"
    Expression ="tbl_Sapling_Data.Sapling_Notes"
    Expression ="tbl_Sapling_Data.Sapling_Data_ID"
    Expression ="tbl_Tags.Microplot_Number"
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
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbText "Description" ="Sapling sampling record is incomplete"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1440"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="6300"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsable"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.DRC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-245
    Top =-19
    Right =1315
    Bottom =645
    Left =-1
    Top =-1
    Right =1528
    Bottom =236
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
        Right =286
        Bottom =177
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
