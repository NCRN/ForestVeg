Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="qActive_Tree_Data"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_Conditions"
    Name ="tlu_Tree_Condition"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Location_ID"
    Expression ="qFiltered_Events.Event_ID"
    Expression ="tbl_Tree_Data.Tree_Data_ID"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Locations.Unit_Code"
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Locations.Panel"
    Expression ="qFiltered_Locations.Frame"
    Expression ="qFiltered_Events.Event_Date"
    Expression ="qActive_Tree_Data.Tag"
    Alias ="ConditionCount"
    Expression ="Sum(IIf(IsNull([Condition]),0,IIf([Pest],0,1)))"
    Alias ="PestCount"
    Expression ="Sum(IIf([Pest],1,0))"
    Alias ="ConditionPresentYN"
    Expression ="IIf([ConditionCount]=0,0,1)"
    Alias ="PestPresentYN"
    Expression ="IIf([PestCount]>0,1,0)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qActive_Tree_Data"
    Expression ="qFiltered_Events.Event_ID = qActive_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="tbl_Tree_Data"
    Expression ="qActive_Tree_Data.Tree_Data_ID = tbl_Tree_Data.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tree_Conditions"
    RightTable ="tlu_Tree_Condition"
    Expression ="tbl_Tree_Conditions.Condition = tlu_Tree_Condition.Description"
    Flag =2
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_Conditions"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Conditions.Tree_Data_ID"
    Flag =2
End
Begin Groups
    Expression ="qFiltered_Locations.Location_ID"
    GroupLevel =0
    Expression ="qFiltered_Events.Event_ID"
    GroupLevel =0
    Expression ="tbl_Tree_Data.Tree_Data_ID"
    GroupLevel =0
    Expression ="qFiltered_Locations.Plot_Name"
    GroupLevel =0
    Expression ="qFiltered_Locations.Unit_Code"
    GroupLevel =0
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qFiltered_Locations.Panel"
    GroupLevel =0
    Expression ="qFiltered_Locations.Frame"
    GroupLevel =0
    Expression ="qFiltered_Events.Event_Date"
    GroupLevel =0
    Expression ="qActive_Tree_Data.Tag"
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
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Frame"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Date"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PestCount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ConditionCount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ConditionPresentYN"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PestPresentYN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =34
    Top =37
    Right =1607
    Bottom =943
    Left =-1
    Top =-1
    Right =1541
    Bottom =538
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =8
        Top =10
        Right =188
        Bottom =341
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =218
        Top =10
        Right =462
        Bottom =386
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =492
        Top =16
        Right =700
        Bottom =399
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
    Begin
        Left =729
        Top =12
        Right =891
        Bottom =287
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =939
        Top =12
        Right =1083
        Bottom =156
        Top =0
        Name ="tbl_Tree_Conditions"
        Name =""
    End
    Begin
        Left =1131
        Top =12
        Right =1275
        Bottom =156
        Top =0
        Name ="tlu_Tree_Condition"
        Name =""
    End
End
