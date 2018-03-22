Operation =1
Option =0
Where ="((([qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed].BA_cm2_PctChange)>25 Or "
    "([qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed].BA_cm2_PctChange)<-25))"
Begin InputTables
    Name ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed"
End
Begin OutputColumns
    Expression ="[qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed].*"
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
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.Status_Prev-Curr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.Sampled_As_Prev-Curr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.Stems_Prev-Curr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].StemList_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].StemList_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].BA_cm2_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].BA_cm2_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].BA_cm2_Change"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].BA_cm2_PctChange"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.[qSum_Tagged_Items_w_Basal_Ar"
            "ea_Current-Previous].Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed.Notes"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =-11
    Right =1184
    Bottom =917
    Left =-1
    Top =-1
    Right =1046
    Bottom =599
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =32
        Top =24
        Right =393
        Bottom =283
        Top =0
        Name ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed"
        Name =""
    End
End
