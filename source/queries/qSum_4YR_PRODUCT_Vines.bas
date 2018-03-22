Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species"
    Name ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species"
End
Begin OutputColumns
    Alias ="Latin Name"
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species.Latin_Name"
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species.TaxonCode"
    Alias ="Exotic YN"
    Expression ="IIf([qCalc_PRODUCT_Tree_Count_Per_Vine_Species]![Exotic]=0,\"No\",\"Yes\")"
    Alias ="% Plots w Species"
    Expression ="Round(([Plot_Count]*100)/DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for"
        "_4_Year_Cycle\"),2)"
    Alias ="% Trees with Vines"
    Expression ="Round(100*[Tree_Count]/DCount(\"[qCalc_PRODUCT_Tree_List_for_Cycle]![Tree_Count_"
        "for_4_Year_Cycle]\",\"qCalc_PRODUCT_Tree_List_for_Cycle\"),3)"
End
Begin Joins
    LeftTable ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species"
    RightTable ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species"
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species.TSN = qCalc_PRODUCT_Tree_Count_Per_Vin"
        "e_Species.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species.Latin_Name"
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
dbText "Description" ="Summarizes the vine species reported during a 4 year cycle.  The first year of t"
    "his cycle MUST be entered as the Year filter for this query to work correctly."
Begin
    Begin
        dbText "Name" ="% Plots w Species"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latin Name"
        dbInteger "ColumnWidth" ="2475"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic YN"
        dbInteger "ColumnWidth" ="1140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="% Trees with Vines"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =-1
    Top =-1
    Right =690
    Bottom =284
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =12
        Top =12
        Right =285
        Bottom =166
        Top =0
        Name ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species"
        Name =""
    End
    Begin
        Left =336
        Top =12
        Right =619
        Bottom =164
        Top =0
        Name ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species"
        Name =""
    End
End
