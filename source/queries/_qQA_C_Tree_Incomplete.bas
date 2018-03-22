Operation =1
Option =0
Where ="(((tbl_Tags.TSN) Is Null) AND ((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Tr"
    "ee_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_Status)<>\""
    "Dead\" And (tbl_Tree_Data.Tree_Status)<>\"And\"<>\"Downgraded to Non-Sampled\"))"
    " OR (((tbl_Tags.Tag) Is Null) AND ((tbl_Events.Event_Date) Is Not Null) AND ((tb"
    "l_Tree_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_Status)"
    "<>\"Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tree_Data.T"
    "ree_Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgraded to "
    "Non-Sampled\")) OR (((tbl_Tree_Data.Tree_Status) Is Null)) OR (((tbl_Tree_Data.T"
    "ree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_Status)<>\"Dead\" An"
    "d (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tree_Data.Tree_Status)<>"
    "\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgraded to Non-Sampled\""
    ") AND ((tbl_Tags.Azimuth) Is Null)) OR (((tbl_Tree_Data.Tree_Status)<>\"Removed "
    "from study\" And (tbl_Tree_Data.Tree_Status)<>\"Dead\" And (tbl_Tree_Data.Tree_S"
    "tatus)<>\"Dead Fallen\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Standing\" And ("
    "tbl_Tree_Data.Tree_Status)<>\"Downgraded to Non-Sampled\") AND ((tbl_Tags.Distan"
    "ce) Is Null)) OR (((tbl_Tree_Data.Tree_Status)<>\"Removed from study\" And (tbl_"
    "Tree_Data.Tree_Status)<>\"Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\""
    " And (tbl_Tree_Data.Tree_Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Stat"
    "us)<>\"Downgraded to Non-Sampled\") AND ((tbl_Tree_Data.Crown_Class) Is Null)) O"
    "R (((tbl_Tree_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_"
    "Status)<>\"Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tree"
    "_Data.Tree_Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgra"
    "ded to Non-Sampled\") AND ((tbl_Tree_Data.Foliage_Conditions_Checked) Is Null)) "
    "OR (((tbl_Tree_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree"
    "_Status)<>\"Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tre"
    "e_Data.Tree_Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgr"
    "aded to Non-Sampled\") AND ((tbl_Tree_Data.Vines_Checked) Is Null)) OR (((tbl_Tr"
    "ee_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_Status)<>\""
    "Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tree_Data.Tree_"
    "Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgraded to Non-"
    "Sampled\") AND ((tbl_Tree_Data.Conditions_Checked) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tbl_Tree_Data"
End
Begin OutputColumns
    Expression ="tbl_Tags.TSN"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="tbl_Locations.Panel"
    Alias ="EventTxt"
    Expression ="StringFromGUID([tbl_Tree_Data]![Event_ID])"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tree_Data.Crown_Class"
    Expression ="tbl_Tree_Data.Foliage_Conditions_Checked"
    Expression ="tbl_Tree_Data.Vines_Checked"
    Expression ="tbl_Tree_Data.Conditions_Checked"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =3
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
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
dbText "Description" ="Tree sampling record is incomplete"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventTxt"
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
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =33
    Top =331
    Right =1528
    Bottom =894
    Left =-1
    Top =-1
    Right =1463
    Bottom =128
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =807
        Top =8
        Right =951
        Bottom =152
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =434
        Top =91
        Right =578
        Bottom =235
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =615
        Top =9
        Right =759
        Bottom =153
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
End
