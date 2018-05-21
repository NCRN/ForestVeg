Version =20
VersionRequired =20
Begin Form
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =187
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7731
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =4545
    Top =810
    Right =12555
    Bottom =10965
    DatasheetGridlinesColor =12632256
    Filter =" [Delete_Table] =  False"
    RecSrcDt = Begin
        0xa4ccdf1a0fb2e340
    End
    RecordSource ="tsys_Import_Log"
    Caption ="Delete Imported Tables"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin FormHeader
            Height =1620
            BackColor =0
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =5160
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label2"
                    Caption ="Delete Imported Tables"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =600
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =240
                    Top =960
                    Width =2886
                    Height =478
                    ColumnOrder =0
                    Name ="optframe_SelectDelete"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =360
                            Top =780
                            Width =1020
                            Height =240
                            BackColor =0
                            ForeColor =16777215
                            Name ="Label8"
                            Caption ="Select....."
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =426
                            Top =1138
                            OptionValue =1
                            Name ="Check10"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =656
                                    Top =1110
                                    Width =720
                                    Height =240
                                    ForeColor =16777215
                                    Name ="Label11"
                                    Caption ="Select All"
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =1560
                            Top =1138
                            OptionValue =2
                            Name ="Check12"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1790
                                    Top =1110
                                    Width =930
                                    Height =240
                                    ForeColor =16777215
                                    Name ="Label13"
                                    Caption ="Select None"
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin Section
            Height =420
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =4980
                    ColumnWidth =5535
                    Name ="txt_Table_Name"
                    ControlSource ="Table_Name"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =960
                            Height =240
                            Name ="Label0"
                            Caption ="Table Name:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7260
                    Top =60
                    ColumnWidth =2070
                    TabIndex =1
                    Name ="chk_Delete_Table"
                    ControlSource ="Delete_Table"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =7260
                    LayoutCachedTop =60
                    LayoutCachedWidth =7520
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6120
                            Top =60
                            Width =1080
                            Height =240
                            Name ="Label1"
                            Caption ="Delete Table?"
                            LayoutCachedLeft =6120
                            LayoutCachedTop =60
                            LayoutCachedWidth =7200
                            LayoutCachedHeight =300
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =540
            BackColor =15527148
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =375
                    Top =120
                    Width =2580
                    FontWeight =700
                    Name ="cmd_Delete"
                    Caption ="Delete Selected Tables"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =375
                    LayoutCachedTop =120
                    LayoutCachedWidth =2955
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6075
                    Top =135
                    Width =840
                    FontWeight =700
                    TabIndex =1
                    Name ="cmd_Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6075
                    LayoutCachedTop =135
                    LayoutCachedWidth =6915
                    LayoutCachedHeight =495
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =3045
                    Top =135
                    Width =2925
                    FontWeight =700
                    TabIndex =2
                    Name ="cmd_Delete_and_Compact"
                    Caption ="Delete Selected and Compact"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3045
                    LayoutCachedTop =135
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =495
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub chk_Delete_Table_AfterUpdate()
If Me.chk_Delete_Table = True Then
    Me.cmd_Delete.Enabled = True
    Me.cmd_Delete_and_Compact.Enabled = True
Else
    Me.cmd_Delete.Enabled = False
    Me.cmd_Delete_and_Compact.Enabled = False
End If

End Sub

Private Sub cmd_Delete_and_Compact_Click()
Dim rs As DAO.Recordset
Set rs = Me.RecordsetClone

Dim strTable As String

'Populate the recordset
rs.MoveLast
rs.MoveFirst

'Cycle through the recordset
Do While Not rs.EOF

    strTable = rs![Table_Name]
       
'Check to see if the delete table check box is checked if not go to the next record.
    If rs![Delete_Table] = False Then
        
        GoTo NextRecord:
            
    ElseIf rs![Delete_Table] = True Then
'If the check box is checked then check to see if the table was deleted already by checking
'the delete date.
        If Not IsNull(rs![Delete_Date]) Then
            GoTo NextRecord:
'If a delete date exists then that table has already been removed.

'If a delete date does not exist and the delete check is checked then delete the table.
        
    ElseIf rs![Delete_Table] = True Then
        
        If IsNull(rs![Delete_Date]) Then
'Delete the table

            DoCmd.DeleteObject acTable, strTable
            
'Update the import log table with the delete date
            With rs
                .Edit
                rs![Delete_Date] = Date
                .Update
            End With
        End If
        End If
    End If
    
NextRecord:
    rs.MoveNext
Loop

Me.Requery

End Sub

Private Sub cmd_Delete_Click()

Dim rs As DAO.Recordset
Set rs = Me.RecordsetClone
Dim strTable As String

'Populate the recordset
rs.MoveLast
rs.MoveFirst

'Cycle through the recordset
Do While Not rs.EOF

    strTable = rs![Table_Name]
       
'Check to see if the delete table check box is checked if not go to the next record.
    If rs![Delete_Table] = False Then
        
        GoTo NextRecord:
            
    ElseIf rs![Delete_Table] = True Then
'If the check box is checked then check to see if the table was deleted already by checking
'the delete date.
        If Not IsNull(rs![Delete_Date]) Then
            GoTo NextRecord:
'If a delete date exists then that table has already been removed.

'If a delete date does not exist and the delete check is checked then delete the table.
        
    ElseIf rs![Delete_Table] = True Then
        
        If IsNull(rs![Delete_Date]) Then
'Delete the table

            DoCmd.DeleteObject acTable, strTable
            
'Update the import log table with the delete date
            With rs
                .Edit
                rs![Delete_Date] = Date
                .Update
            End With
        End If
        End If
    End If
    
NextRecord:
    rs.MoveNext
Loop

Me.Requery
'Need to find better code for compacting
'DBEngine.CompactDatabase

End Sub

Private Sub Form_Open(Cancel As Integer)

Me.filter = " [Delete_Table] =  " & False

Me.FilterOn = True

End Sub

Private Sub cmd_Close_Click()
On Error GoTo Err_cmd_Close_Click
     
        DoCmd.Close

Exit_cmd_Close_Click:
    Exit Sub
Err_cmd_Close_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Close_Click
End Sub

Private Sub optframe_SelectDelete_AfterUpdate()

Dim rsDelete As DAO.Recordset
Set rsDelete = Me.RecordsetClone

rsDelete.MoveFirst

Do Until rsDelete.EOF

If Me!optframe_SelectDelete.Value = 1 Then
     
    rsDelete.Edit
    rsDelete![Delete_Table] = True
    rsDelete.Update
    
ElseIf Me!optframe_SelectDelete.Value = 2 Then
    rsDelete.Edit
    rsDelete![Delete_Table] = False
    rsDelete.Update
    
Else: GoTo NextRecord:

End If

NextRecord:
rsDelete.MoveNext

Loop

  Me.cmd_Delete.Enabled = True

End Sub
