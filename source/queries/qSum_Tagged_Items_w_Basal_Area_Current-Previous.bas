Operation =1
Option =0
Where ="(((tbl_Locations.Panel)=1) AND ((qSum_Tagged_Items_w_Basal_Area_Previous.Status)"
    " Is Not Null)) OR (((qSum_Tagged_Items_w_Basal_Area_Current.Status) Is Not Null)"
    ")"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Tags"
    Name ="qSum_Tagged_Items_w_Basal_Area_Previous"
    Name ="qSum_Tagged_Items_w_Basal_Area_Current"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.TSN"
    Alias ="Sampled_As_Previous"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.Sampled_As"
    Alias ="Status_Previous"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.Status"
    Alias ="CrownClassPrev"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.CrownClass"
    Alias ="Notes_Previous"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.Notes"
    Alias ="Stems_Previous"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.Stems"
    Alias ="BA_cm2_Previous"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.SumBasalArea_cm2"
    Alias ="Sampled_As_Current"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.Sampled_As"
    Alias ="Status_Current"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.Status"
    Alias ="CrownClassCurrent"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.CrownClass"
    Alias ="Notes_Current"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.Notes"
    Alias ="Stems_Current"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.Stems"
    Alias ="BA_cm2_Current"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.SumBasalArea_cm2"
    Alias ="Stem_Change"
    Expression ="[Stems_Current]-[Stems_Previous]"
    Alias ="BA_cm2_Change"
    Expression ="Round([BA_cm2_Current]-[BA_cm2_Previous],1)"
    Alias ="BA_cm2_PctChange"
    Expression ="Round(100*([ba_cm2_change]/[ba_cm2_Previous]),1)"
    Expression ="tbl_Tags.Tag_Status"
    Alias ="StemList_Previous"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Previous.StemList"
    Alias ="StemList_Current"
    Expression ="qSum_Tagged_Items_w_Basal_Area_Current.StemList"
    Expression ="tlu_Plants.Latin_Name"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="qSum_Tagged_Items_w_Basal_Area_Previous"
    Expression ="tbl_Tags.Tag_ID = qSum_Tagged_Items_w_Basal_Area_Previous.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="qSum_Tagged_Items_w_Basal_Area_Current"
    Expression ="tbl_Tags.Tag_ID = qSum_Tagged_Items_w_Basal_Area_Current.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BA_cm2_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StemList_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sampled_As_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sampled_As_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stems_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StemList_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BA_cm2_Current"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BA_cm2_PctChange"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stem_Change"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BA_cm2_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stems_Previous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrownClassCurrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrownClassPrev"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-10
    Top =196
    Right =1482
    Bottom =983
    Left =-1
    Top =-1
    Right =1460
    Bottom =401
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
        Left =240
        Top =12
        Right =381
        Bottom =361
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =978
        Top =11
        Right =1245
        Bottom =331
        Top =0
        Name ="qSum_Tagged_Items_w_Basal_Area_Previous"
        Name =""
    End
    Begin
        Left =433
        Top =65
        Right =781
        Bottom =212
        Top =0
        Name ="qSum_Tagged_Items_w_Basal_Area_Current"
        Name =""
    End
    Begin
        Left =417
        Top =230
        Right =561
        Bottom =374
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
