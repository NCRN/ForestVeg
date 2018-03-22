Operation =1
Option =0
Having ="((([qry_Trees_2006-2009].TSN) Is Null))"
Begin InputTables
    Name ="qry_Trees_2006-2009"
    Name ="qry_Trees_2010-2013"
End
Begin OutputColumns
    Expression ="[qry_Trees_2010-2013].Unit_Code"
    Expression ="[qry_Trees_2010-2013].TaxonCode"
    Expression ="[qry_Trees_2010-2013].TSN"
    Expression ="[qry_Trees_2010-2013].Family"
    Expression ="[qry_Trees_2010-2013].Genus"
    Expression ="[qry_Trees_2010-2013].Species"
    Expression ="[qry_Trees_2010-2013].Subspecies"
    Expression ="[qry_Trees_2010-2013].Common"
    Expression ="[qry_Trees_2006-2009].TSN"
End
Begin Joins
    LeftTable ="qry_Trees_2010-2013"
    RightTable ="qry_Trees_2006-2009"
    Expression ="[qry_Trees_2010-2013].Unit_Code = [qry_Trees_2006-2009].Unit_Code"
    Flag =2
    LeftTable ="qry_Trees_2010-2013"
    RightTable ="qry_Trees_2006-2009"
    Expression ="[qry_Trees_2010-2013].TaxonCode = [qry_Trees_2006-2009].TaxonCode"
    Flag =2
End
Begin OrderBy
    Expression ="[qry_Trees_2010-2013].Family"
    Flag =0
    Expression ="[qry_Trees_2010-2013].Genus"
    Flag =0
    Expression ="[qry_Trees_2010-2013].Species"
    Flag =0
End
Begin Groups
    Expression ="[qry_Trees_2010-2013].Unit_Code"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].TaxonCode"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].TSN"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].Family"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].Genus"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].Species"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].Subspecies"
    GroupLevel =0
    Expression ="[qry_Trees_2010-2013].Common"
    GroupLevel =0
    Expression ="[qry_Trees_2006-2009].TSN"
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
        dbText "Name" ="[qry_Trees_2010-2013].Common"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qry_Trees_2006-2009].TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3390"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].Subspecies"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qry_Trees_2010-2013].TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =238
    Top =47
    Right =1351
    Bottom =1003
    Left =-1
    Top =-1
    Right =1081
    Bottom =623
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =137
        Top =105
        Right =289
        Bottom =333
        Top =0
        Name ="qry_Trees_2006-2009"
        Name =""
    End
    Begin
        Left =416
        Top =108
        Right =661
        Bottom =414
        Top =0
        Name ="qry_Trees_2010-2013"
        Name =""
    End
End
