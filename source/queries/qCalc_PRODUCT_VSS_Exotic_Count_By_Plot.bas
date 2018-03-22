Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_All_Occurences"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Admin Unit Code]"
    Expression ="qSum_4YR_PRODUCT_All_Occurences.Plot_Name"
    Alias ="ExoticYN_Bin"
    Expression ="Sum(IIf([Exotic YN]=\"Yes\",1,0))"
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Admin Unit Code]"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_All_Occurences.Plot_Name"
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
        dbText "Name" ="qSum_4YR_PRODUCT_All_Occurences.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExoticYN_Bin"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_All_Occurences.[Admin Unit Code]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1518
    Bottom =947
    Left =-1
    Top =-1
    Right =1448
    Bottom =561
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =343
        Top =26
        Right =613
        Bottom =170
        Top =0
        Name ="qSum_4YR_PRODUCT_All_Occurences"
        Name =""
    End
End
