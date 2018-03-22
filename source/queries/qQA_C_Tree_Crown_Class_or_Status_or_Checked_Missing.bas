Operation =1
Option =0
Where ="(((tbl_Tree_Data.Crown_Class) Is Null) AND ((tbl_Tree_Data.Tree_Status)=\"Alive "
    "Broken\" And (tbl_Tree_Data.Tree_Status)=\"Alive Fallen\" And (tbl_Tree_Data.Tre"
    "e_Status)=\"Alive Leaning\" And (tbl_Tree_Data.Tree_Status)=\"Alive Standing\"))"
    " OR (((tbl_Tree_Data.Tree_Status) Is Null)) OR (((tbl_Tags.Tag) Is Null) AND ((t"
    "bl_Tree_Data.Tree_Status)=\"Alive Broken\" And (tbl_Tree_Data.Tree_Status)=\"Ali"
    "ve Fallen\" And (tbl_Tree_Data.Tree_Status)=\"Alive Leaning\" And (tbl_Tree_Data"
    ".Tree_Status)=\"Alive Standing\")) OR (((tbl_Tree_Data.Tree_Status)=\"Alive Brok"
    "en\" And (tbl_Tree_Data.Tree_Status)=\"Alive Fallen\" And (tbl_Tree_Data.Tree_St"
    "atus)=\"Alive Leaning\" And (tbl_Tree_Data.Tree_Status)=\"Alive Standing\") AND "
    "((tbl_Tree_Data.Vines_Checked) Is Null)) OR (((tbl_Tree_Data.Tree_Status)=\"Aliv"
    "e Broken\" And (tbl_Tree_Data.Tree_Status)=\"Alive Fallen\" And (tbl_Tree_Data.T"
    "ree_Status)=\"Alive Leaning\" And (tbl_Tree_Data.Tree_Status)=\"Alive Standing\""
    ") AND ((tbl_Tree_Data.Conditions_Checked) Is Null)) OR (((tbl_Tree_Data.Tree_Sta"
    "tus)=\"Alive Broken\" And (tbl_Tree_Data.Tree_Status)=\"Alive Fallen\" And (tbl_"
    "Tree_Data.Tree_Status)=\"Alive Leaning\" And (tbl_Tree_Data.Tree_Status)=\"Alive"
    " Standing\") AND ((tbl_Tree_Data.Foliage_Conditions_Checked) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tree_Data.Crown_Class"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="tbl_Tree_Data.Vines_Checked"
    Expression ="tbl_Tree_Data.Conditions_Checked"
    Expression ="tbl_Tree_Data.Foliage_Conditions_Checked"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tree_Data.Tree_Data_ID"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
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
dbText "Description" ="Tree Crown Class, Status, or one of the \"parameters checked\" checkboxes are em"
    "pty"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
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
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =22
    Top =-6
    Right =1546
    Bottom =704
    Left =-1
    Top =-1
    Right =1492
    Bottom =321
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =46
        Top =23
        Right =190
        Bottom =167
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =348
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =439
        Top =13
        Right =583
        Bottom =249
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =635
        Top =169
        Right =779
        Bottom =465
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
