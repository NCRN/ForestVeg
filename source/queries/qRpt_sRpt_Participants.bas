Operation =1
Option =0
Begin InputTables
    Name ="tbl_Events"
    Name ="tlu_Contacts"
    Name ="xref_Event_Contacts"
End
Begin OutputColumns
    Expression ="tbl_Events.Event_ID"
    Expression ="tlu_Contacts.First_Name"
    Expression ="tlu_Contacts.Last_Name"
    Expression ="xref_Event_Contacts.Contact_Role"
End
Begin Joins
    LeftTable ="tlu_Contacts"
    RightTable ="xref_Event_Contacts"
    Expression ="tlu_Contacts.Contact_ID=xref_Event_Contacts.Contact_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="xref_Event_Contacts"
    Expression ="tbl_Events.Event_ID=xref_Event_Contacts.Event_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tlu_Contacts.Last_Name"
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
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Contacts.Last_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Contacts.First_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xref_Event_Contacts.Contact_Role"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1010
    Bottom =675
    Left =-1
    Top =-1
    Right =872
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =432
        Top =12
        Right =720
        Bottom =225
        Top =0
        Name ="tlu_Contacts"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =208
        Bottom =159
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="xref_Event_Contacts"
        Name =""
    End
End
