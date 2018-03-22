Operation =1
Option =0
Where ="(((tbl_Tags.Tag_ID) Is Null))"
Begin InputTables
    Name ="tbl_Tags_History"
    Name ="tbl_Tags"
    Name ="tlu_Contacts"
End
Begin OutputColumns
    Expression ="tbl_Tags_History.Tags_History_ID"
    Alias ="Table"
    Expression ="tbl_Tags_History.Table_Name"
    Expression ="tbl_Tags_History.Record_ID_Field_Name"
    Expression ="tbl_Tags_History.Record_ID"
    Expression ="tbl_Tags.Tag"
    Alias ="Field"
    Expression ="tbl_Tags_History.Field_Name"
    Expression ="tbl_Tags_History.Value_New"
    Expression ="tbl_Tags_History.Value_Old"
    Expression ="tbl_Tags_History.Value_History_Notes"
    Expression ="tbl_Tags_History.Contact_ID"
    Expression ="tlu_Contacts.Last_Name"
    Expression ="tbl_Tags_History.Network_User_Name"
    Expression ="tbl_Tags_History.Change_Date"
    Expression ="tbl_Tags_History.Table_Name"
    Expression ="tbl_Tags_History.Updated_Date"
End
Begin Joins
    LeftTable ="tbl_Tags_History"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tags_History.Record_ID = tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="tbl_Tags_History"
    RightTable ="tlu_Contacts"
    Expression ="tbl_Tags_History.Contact_ID = tlu_Contacts.Contact_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Tags_History.Table_Name"
    Flag =0
    Expression ="tbl_Tags_History.Record_ID_Field_Name"
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
dbText "Description" ="Tag history record exists but is not matched to an existing tag"
Begin
    Begin
        dbText "Name" ="[tbl_Tags_History].[Tags_History_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Table_Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Record_ID_Field_Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Record_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Value_New]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Value_Old]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Value_History_Notes]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Contact_ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Network_User_Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Change_Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tbl_Tags_History].[Updated_Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Tags_History_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Table_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Record_ID_Field_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Record_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_New"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_Old"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Value_History_Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4005"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Contact_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4110"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Network_User_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Change_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Updated_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="tlu_Contacts.Last_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Table"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =149
    Right =1380
    Bottom =732
    Left =-1
    Top =-1
    Right =1309
    Bottom =281
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =168
        Top =18
        Right =369
        Bottom =282
        Top =0
        Name ="tbl_Tags_History"
        Name =""
    End
    Begin
        Left =383
        Top =16
        Right =559
        Bottom =282
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =712
        Top =108
        Right =918
        Bottom =282
        Top =0
        Name ="tlu_Contacts"
        Name =""
    End
End
