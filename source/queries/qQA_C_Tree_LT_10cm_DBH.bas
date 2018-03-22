Operation =1
Option =0
Where ="(((tbl_Tree_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_St"
    "atus)<>\"Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tree_D"
    "ata.Tree_Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgrade"
    "d to Non-Sampled\") AND ((qCalc_Basal_Area_per_Tree.SumBasalArea_cm2)<78.5)) OR "
    "(((tbl_Tree_Data.Tree_Status)<>\"Removed from study\" And (tbl_Tree_Data.Tree_St"
    "atus)<>\"Dead\" And (tbl_Tree_Data.Tree_Status)<>\"Dead Fallen\" And (tbl_Tree_D"
    "ata.Tree_Status)<>\"Dead Standing\" And (tbl_Tree_Data.Tree_Status)<>\"Downgrade"
    "d to Non-Sampled\") AND ((qCalc_Basal_Area_per_Tree.SumBasalArea_cm2) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
    Name ="qCalc_Basal_Area_per_Tree"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tree_Data.Crown_Class"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="qCalc_Basal_Area_per_Tree"
    Expression ="tbl_Tree_Data.Tree_Data_ID = qCalc_Basal_Area_per_Tree.Tree_Data_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
    Flag =0
    Expression ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
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
dbText "Description" ="Tree equivalent DBH is less than 10cm"
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
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
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
End
Begin
    State =0
    Left =8
    Top =233
    Right =1329
    Bottom =662
    Left =-1
    Top =-1
    Right =1289
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =610
        Top =235
        Right =754
        Bottom =420
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =18
        Top =10
        Right =162
        Bottom =411
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =192
        Top =14
        Right =336
        Bottom =238
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =388
        Top =106
        Right =532
        Bottom =350
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =760
        Top =25
        Right =904
        Bottom =169
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
End
