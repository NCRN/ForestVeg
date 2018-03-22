Operation =1
Option =0
Having ="(((tlu_Plants.Exotic)=True))"
Begin InputTables
    Name ="qActive_Trees_Shrubs_Herbs_Vines"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    Expression ="tlu_Plants.Exotic"
    Alias ="Count_of_Specimens"
    Expression ="Count(qActive_Trees_Shrubs_Herbs_Vines.TSN)"
End
Begin Joins
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tlu_Plants"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN=tlu_Plants.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
    Flag =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    Flag =0
End
Begin Groups
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
    GroupLevel =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    GroupLevel =0
    Expression ="tlu_Plants.Exotic"
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
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Specimens"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =104
    Top =52
    Right =1131
    Bottom =929
    Left =-1
    Top =-1
    Right =1003
    Bottom =547
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =72
        Top =33
        Right =448
        Bottom =295
        Top =0
        Name ="qActive_Trees_Shrubs_Herbs_Vines"
        Name =""
    End
    Begin
        Left =496
        Top =12
        Right =640
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
