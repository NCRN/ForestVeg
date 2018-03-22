Operation =1
Option =0
Having ="(((qSum_4YR_PRODUCT_All_Occurences.[Exotic YN])=\"No\") AND ((IIf([Habit-Class] "
    "Like \"Shrub*\",\"Shrub\",\"Other\"))=\"Shrub\"))"
Begin InputTables
    Name ="qSum_4YR_PRODUCT_All_Occurences"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Admin Unit Code]"
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Latin Name]"
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Exotic YN]"
    Alias ="Habit"
    Expression ="IIf([Habit-Class] Like \"Shrub*\",\"Shrub\",\"Other\")"
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Admin Unit Code]"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Latin Name]"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_All_Occurences.[Exotic YN]"
    GroupLevel =0
    Expression ="IIf([Habit-Class] Like \"Shrub*\",\"Shrub\",\"Other\")"
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
        dbText "Name" ="qSum_4YR_PRODUCT_All_Occurences.[Admin Unit Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_All_Occurences.[Exotic YN]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_All_Occurences.[Latin Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1518
    Bottom =947
    Left =-1
    Top =-1
    Right =1410
    Bottom =515
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =324
        Bottom =198
        Top =0
        Name ="qSum_4YR_PRODUCT_All_Occurences"
        Name =""
    End
End
