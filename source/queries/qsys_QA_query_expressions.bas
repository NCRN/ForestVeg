Operation =1
Option =0
Where ="(((MSysObjects.Name) Like \"qQA*\") AND ((MSysQueries.Attribute)=8 Or (MSysQueri"
    "es.Attribute)=10) AND ((MSysQueries.Expression) Is Not Null))"
Begin InputTables
    Name ="MSysObjects"
    Name ="MSysQueries"
End
Begin OutputColumns
    Expression ="MSysObjects.Name"
    Expression ="MSysQueries.Attribute"
    Expression ="MSysQueries.Expression"
End
Begin Joins
    LeftTable ="MSysObjects"
    RightTable ="MSysQueries"
    Expression ="MSysObjects.Id=MSysQueries.ObjectId"
    Flag =2
End
Begin OrderBy
    Expression ="MSysObjects.Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="System query for showing the WHERE expression statements for quality assurance q"
    "ueries"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="MSysObjects.Name"
        dbInteger "ColumnWidth" ="4272"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSysQueries.Expression"
        dbInteger "ColumnWidth" ="17712"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSysQueries.Attribute"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =49
    Top =55
    Right =1031
    Bottom =403
    Left =-1
    Top =-1
    Right =958
    Bottom =146
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =168
        Bottom =114
        Top =0
        Name ="MSysObjects"
        Name =""
    End
    Begin
        Left =216
        Top =7
        Right =336
        Bottom =114
        Top =0
        Name ="MSysQueries"
        Name =""
    End
End
