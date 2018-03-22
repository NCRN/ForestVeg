Operation =1
Option =0
Where ="(((qCalc_Basal_Area_per_Tree_and_Sapling.Sample_Year)>=2008 And (qCalc_Basal_Are"
    "a_per_Tree_and_Sapling.Sample_Year)<=2011) AND ((tlu_Plants.Genus)=\"Fraxinus\")"
    ")"
Begin InputTables
    Name ="qCalc_Basal_Area_per_Tree_and_Sapling"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Plot_Name"
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Sample_Year"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.Common"
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Sampled_As"
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Status"
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Stems"
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.SumBasalArea_cm2"
End
Begin Joins
    LeftTable ="qCalc_Basal_Area_per_Tree_and_Sapling"
    RightTable ="tbl_Tags"
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling.Plot_Name"
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
dbText "Description" ="Where were Fraxinus trees and saplings found from 2008 - 2011 and what was their"
    " basal area?"
Begin
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling.Sampled_As"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling.Status"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling.SumBasalArea_cm2"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =10
    Right =1451
    Bottom =947
    Left =-1
    Top =-1
    Right =1402
    Bottom =650
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =202
        Bottom =377
        Top =0
        Name ="qCalc_Basal_Area_per_Tree_and_Sapling"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =398
        Bottom =283
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =683
        Top =101
        Right =841
        Bottom =526
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =471
        Top =10
        Right =615
        Bottom =402
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
