Operation =1
Option =0
Begin InputTables
    Name ="qSum_Tagged_Items_w_Basal_Area_Current-Previous"
End
Begin OutputColumns
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Plot_Name"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Tag"
    Alias ="Status_Prev-Curr"
    Expression ="Nz([Status_Previous],\"-- \") & \"/\" & [Status_Current]"
    Alias ="Sampled_As_Prev-Curr"
    Expression ="Nz([Sampled_As_Previous],\"-- \") & \"/\" & [Sampled_As_Current]"
    Alias ="CrownClass_Prev_Curr"
    Expression ="Nz([CrownClassPrev],\"---\") & \"/\" & Nz([CrownClassCurrent],\"---\")"
    Alias ="Stems_Prev-Curr"
    Expression ="Nz([Stems_Previous],\"-- \") & \"/\" & [Stems_Current]"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].StemList_Previous"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].StemList_Current"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_Previous"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_Current"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_Change"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_PctChange"
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Latin_Name"
    Alias ="Notes"
    Expression ="\"PREVIOUS: \" & [Notes_Previous] & \" CURRENT: \" & [Notes_Current]"
End
Begin OrderBy
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Plot_Name"
    Flag =0
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Tag"
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
Begin
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].StemList_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stems_Prev-Curr"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sampled_As_Prev-Curr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status_Prev-Curr"
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_PctChange"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_Change"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].BA_cm2_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].StemList_Current"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qSum_Tagged_Items_w_Basal_Area_Current-Previous].Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbInteger "ColumnWidth" ="8850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrownClass_Prev_Curr"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =47
    Top =141
    Right =1506
    Bottom =867
    Left =-1
    Top =-1
    Right =1427
    Bottom =368
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =86
        Top =7
        Right =729
        Bottom =369
        Top =0
        Name ="qSum_Tagged_Items_w_Basal_Area_Current-Previous"
        Name =""
    End
End
