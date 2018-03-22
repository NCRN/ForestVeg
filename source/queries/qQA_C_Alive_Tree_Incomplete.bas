Operation =1
Option =0
Where ="(((tbl_Tags.TSN) Is Null) AND ((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Tr"
    "ee_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Fal"
    "len\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Leaning\" Or (tbl_Tree_Data.Tree_St"
    "atus)=\"Alive Standing\")) OR (((tbl_Tags.Tag) Is Null) AND ((tbl_Events.Event_D"
    "ate) Is Not Null) AND ((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tree"
    "_Data.Tree_Status)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Leani"
    "ng\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Standing\")) OR (((tbl_Tree_Data.Tre"
    "e_Status) Is Null)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tr"
    "ee_Data.Tree_Status)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Lea"
    "ning\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tags.Azimuth"
    ") Is Null)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tree_Data."
    "Tree_Status)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Leaning\" O"
    "r (tbl_Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tags.Distance) Is Nu"
    "ll)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tree_Data.Tree_St"
    "atus)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Leaning\" Or (tbl_"
    "Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tree_Data.Crown_Class) Is N"
    "ull)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tree_Data.Tree_S"
    "tatus)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Leaning\" Or (tbl"
    "_Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tree_Data.Foliage_Conditio"
    "ns_Checked)=False)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_Tr"
    "ee_Data.Tree_Status)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Lea"
    "ning\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tree_Data.Vi"
    "nes_Checked)=False)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or (tbl_T"
    "ree_Data.Tree_Status)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Le"
    "aning\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tree_Data.C"
    "onditions_Checked)=False)) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Broken\" Or "
    "(tbl_Tree_Data.Tree_Status)=\"Alive Fallen\" Or (tbl_Tree_Data.Tree_Status)=\"Al"
    "ive Leaning\" Or (tbl_Tree_Data.Tree_Status)=\"Alive Standing\") AND ((tbl_Tree_"
    "DBH.DBH) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_DBH"
End
Begin OutputColumns
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
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
    Expression ="tbl_Tree_DBH.DBH"
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
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_DBH"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID"
    Flag =2
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
dbText "Description" ="Tree sampling record is incomplete"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventTxt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5865"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Conditions_Checked"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-76
    Top =101
    Right =1479
    Bottom =1097
    Left =-1
    Top =-1
    Right =1523
    Bottom =261
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
        Left =852
        Top =102
        Right =996
        Bottom =246
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
        Right =801
        Bottom =241
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =999
        Top =12
        Right =1293
        Bottom =217
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
End
