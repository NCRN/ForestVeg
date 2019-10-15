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
    Width =7800
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =1170
    Top =15
    Right =9510
    Bottom =5520
    DatasheetGridlinesColor =12632256
    Filter =" [Delete_Date] IS NULL"
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
                    Name ="lblTitle"
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
                    Name ="optgSelectDelete"
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
                            Name ="lblSelect"
                            Caption ="Select....."
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =426
                            Top =1138
                            OptionValue =1
                            Name ="chkALL"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =656
                                    Top =1110
                                    Width =720
                                    Height =240
                                    ForeColor =16777215
                                    Name ="lblALL"
                                    Caption ="Select All"
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =1560
                            Top =1138
                            TabIndex =1
                            OptionValue =2
                            Name ="chkNone"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1790
                                    Top =1110
                                    Width =930
                                    Height =240
                                    ForeColor =16777215
                                    Name ="lblNone"
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
                    Name ="tbxTableName"
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
                            Name ="lblTableName"
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
                    Name ="chkDeleteTable"
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
                            Name ="lblDeleteTable"
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
                    Name ="btnDelete"
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
                    Name ="btnClose"
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
                    Name ="btnDeleteAndCompact"
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
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =180
                    Width =420
                    Height =255
                    TabIndex =3
                    Name ="tbxRecordCount"

                    LayoutCachedLeft =6960
                    LayoutCachedTop =180
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =435
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7440
                    Top =180
                    Width =360
                    Height =255
                    TabIndex =4
                    Name ="tbxDeleteCount"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =180
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =435
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
Option Explicit

' =================================
' FORM:         frm_Append_Delete_Tables
' Level:        Form module
' Version:      1.01
'
' Description:  delete append tables related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 24, 2018
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 8/18/2018 - 1.01 - added documentation, error handling,
'                                          check for table existance before delete
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Events
' ----------------

' ----------------
'  Form
' ----------------
' ---------------------------------
' SUB:          Form_Open
' Description:  form open actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC    - 8/18/2018 - added documentation & error handling,
'                        adjusted filter to include all non-deleted tables
'                        selected for deletion but not yet deleted
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
   
   'use NULL Delete_Date vs. Delete_Table = False since tables could have been
   'selected for deletion, but no deletion completed
    Me.Filter = " [Delete_Date] IS NULL"   '" [Delete_Table] =  " & False
    
    Me.FilterOn = True
    
    'defaults
    btnDelete.Enabled = False   'check for table records first
    btnDeleteAndCompact.Enabled = False
    
   'populate the count
   Dim rs As DAO.Recordset
   Set rs = Me.RecordsetClone
   If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        tbxRecordCount = rs.RecordCount
   
        'clear all Delete checkboxes
        rs.MoveFirst
        Do Until rs.EOF
            If rs!Delete_Table = True And IsNull(rs!Delete_Date) = True Then
                rs.Edit
                rs!Delete_Table = False
                rs.Update
            End If
            
            rs.MoveNext
        Loop
   Else
        tbxRecordCount = 0
   End If
      
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDeleteAndCompact_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC    - 8/18/2018 - added documentation & error handling
' ---------------------------------
Private Sub btnDeleteAndCompact_Click()
On Error GoTo Err_Handler

    DeleteTables
'    Dim rs As DAO.Recordset
'    Set rs = Me.RecordsetClone
'
'    Dim strTable As String
'
'    'Populate the recordset
'    rs.MoveLast
'    rs.MoveFirst
'
'    'Cycle through the recordset
'    Do While Not rs.EOF
'
'        strTable = rs![Table_Name]
'
'    'Check to see if the delete table check box is checked if not go to the next record.
'        If rs![Delete_Table] = False Then
'
'            GoTo NextRecord:
'
'        ElseIf rs![Delete_Table] = True Then
'    'If check box checked check if table deleted already by checking delete date
'            If Not IsNull(rs![Delete_Date]) Then
'                GoTo NextRecord:
'    'If delete date exists then that table has already been removed.
'
'    'If a delete date does not exist and the delete check is checked then delete the table.
'
'        ElseIf rs![Delete_Table] = True Then
'
'            If IsNull(rs![Delete_Date]) Then
'    'Delete the table
'
'                DoCmd.DeleteObject acTable, strTable
'
'    'Update the import log table with the delete date
'                With rs
'                    .Edit
'                    rs![Delete_Date] = Date
'                    .Update
'                End With
'            End If
'            End If
'        End If
'
'NextRecord:
'        rs.MoveNext
'    Loop
'
'    Me.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDeleteAndCompact_Click[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDelete_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC    - 8/18/2018 - added documentation & error handling
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler

    DeleteTables
'    Dim rs As DAO.Recordset
'    Set rs = Me.RecordsetClone
'    Dim strTable As String
'
'    'Populate the recordset
'    rs.MoveLast
'    rs.MoveFirst
'
'    'Cycle through the recordset
'    Do While Not rs.EOF
'
'        strTable = rs![Table_Name]
'
'        'Check if delete table check box is checked if not go to next record
'        If rs![Delete_Table] = False Then
'
'            GoTo NextRecord:
'
'        ElseIf rs![Delete_Table] = True Then
'        'If check box checked then check if table was deleted already by checking
'        'delete date
'            If Not IsNull(rs![Delete_Date]) Then
'                GoTo NextRecord:
'                'If delete date exists --> table already removed
'
'                'If delete date does not exist & delete check checked --> delete table
'
'            ElseIf rs![Delete_Table] = True Then
'
'                If IsNull(rs![Delete_Date]) Then
'                    'Delete table if it exists
'                    If TableExists(strTable) Then _
'                        DoCmd.DeleteObject acTable, strTable
'
'                    'Update import log table with delete date
'                    '***************************************************************************
'                    ' NOTE: This flags all tables set as deleted including
'                    '       those deleted through the UI (outside this procedure)
'                    '       For tables deleted through the UI, delete date will not be accurate
'                    '***************************************************************************
'                    With rs
'                        .Edit
'                        rs![Delete_Date] = Date
'                        .Update
'                    End With
'                End If
'            End If
'        End If
'
'NextRecord:
'        rs.MoveNext
'    Loop
'
'    Me.Requery
'
'    'Need to find better code for compacting
'    'DBEngine.CompactDatabase

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC    - 8/18/2018 - added documentation & error handling
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler
     
    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkDeleteTable_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC    - 8/18/2018 - added documentation & error handling
' ---------------------------------
Private Sub chkDeleteTable_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me.chkDeleteTable = True Then
        Me.btnDelete.Enabled = True
        'Me.btnDeleteAndCompact.Enabled = True
        Me.tbxDeleteCount = Nz(Me.tbxDeleteCount, 0) + 1
    Else
        Me.btnDelete.Enabled = False
        Me.btnDeleteAndCompact.Enabled = False
        Me.tbxDeleteCount = Nz(Me.tbxDeleteCount, 0) - 1
    End If

    Me.Repaint
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkDeleteTable_AfterUpdate[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          optgSeleteDelete_AfterUpdate
' Description:  option group after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC    - 8/18/2018 - added documentation & error handling,
'                        check for EOF & BOF (no records)
' ---------------------------------
Private Sub optgSelectDelete_AfterUpdate()
On Error GoTo Err_Handler

    Dim rsDelete As DAO.Recordset
    Set rsDelete = Me.RecordsetClone
    
    'ensure there are records
    If Not (rsDelete.EOF And rsDelete.BOF) Then
    
        rsDelete.MoveFirst
        
        Do Until rsDelete.EOF
        
        If Me!optgSelectDelete.Value = 1 Then
             
            rsDelete.Edit
            rsDelete![Delete_Table] = True
            rsDelete.Update
            
        ElseIf Me!optgSelectDelete.Value = 2 Then
            rsDelete.Edit
            rsDelete![Delete_Table] = False
            rsDelete.Update
            
        Else: GoTo NextRecord:
        
        End If
        
NextRecord:
        rsDelete.MoveNext
        
        Loop
        
        Me.btnDelete.Enabled = True
        
    Else
        'no records to delete
        Me.btnDelete.Enabled = False

    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optgSelectDelete_AfterUpdate[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          DeleteTables
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, 8/18/2019
' Adapted:      -
' Revisions:
'   BLC    - 8/18/2018 - initial version, shifted from btnDelete & btnDeleteAndCompact click events
' ---------------------------------
Private Sub DeleteTables()
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Set rs = Me.RecordsetClone
    Dim strTable As String
    Dim i As Integer
    
    'Populate the recordset
    rs.MoveLast
    rs.MoveFirst
        
    'initialize i
    i = 0

'    SetProgress i, Me.tbxDeleteCount, "Deleting Append Tables..." 'rs.RecordCount, "Deleting Append Tables..."
    
    'Cycle through the recordset
    Do Until rs.EOF
    
    Debug.Print rs![Table_Name]
        
            i = i + 1
 Debug.Print "i=" & i
        
        strTable = rs![Table_Name]
        
        'Check if delete table check box is checked
        '   --> if not go to next record
        '   --> if so, check if table was deleted by checking delete date
        
        If rs![Delete_Table] = True Then
            
            'Check if table was deleted by checking delete date
            'If delete date exists         --> table already removed
            'If delete date does not exist --> delete table (if delete table checked)
            
            If IsNull(rs![Delete_Date]) Then
            
                'If delete date does not exist & delete check checked
                '  --> delete table

'                SetProgress i - 1, tbxDeleteCount, "Deleting Append Tables..."
                    
                'Delete table if it exists
                If TableExists(strTable) Then
                    DoCmd.DeleteObject acTable, strTable
                
                    Debug.Print "deleted"
                End If
                
                'Update import log table with delete date
                '***************************************************************************
                ' NOTE: This flags all tables set as deleted including
                '       those deleted through the UI (outside this procedure)
                '       For tables deleted through the UI, delete date will not be accurate
                '***************************************************************************
                With rs
                    .Edit
                    rs![Delete_Date] = Date
                    .Update
                    Me.Requery
                End With
                
            End If
        
        End If
        
        'go to next record
        rs.MoveNext
        
        'Me.Requery
        Debug.Print i
    Loop
    
'    SetProgress Me.tbxDeleteCount, Me.tbxDeleteCount, "Deleting Append Tables COMPLETE!"
    'update counts
    Me.tbxDeleteCount = 0
    Me.tbxRecordCount = Me.RecordsetClone.RecordCount
    
    Me.Requery
    
'    FormRefresh
    
    'Need to find better code for compacting
    'DBEngine.CompactDatabase

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeleteTables[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetProgress
' Description:  Set & update progress bar display
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, 8/18/2019
' Adapted:      -
' Revisions:
'   BLC    - 8/18/2018 - initial version
' ---------------------------------
Public Sub SetProgress(CurrentStep As Integer, TotalSteps As Integer, msg As String)
On Error GoTo Err_Handler

    'initialize steps if step is 0 or 1
    If CurrentStep = 0 Then _
        SysCmd acSysCmdInitMeter, msg, TotalSteps
    
    'update meter
    SysCmd acSysCmdUpdateMeter, CurrentStep
    DoEvents
    
    If CurrentStep = TotalSteps Then
        'Report finished & remove meter
        SysCmd acSysCmdRemoveMeter
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetProgress[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          FormRefresh
' Description:  Refreshes & recalculates form display
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, 8/18/2019
' Adapted:      -
' Revisions:
'   BLC    - 8/18/2018 - initial version
' ---------------------------------
Public Sub FormRefresh()
On Error GoTo Err_Handler

    Me.Requery
    Me.tbxDeleteCount = 0

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormRefresh[frm_Append_Delete_Tables])"
    End Select
    Resume Exit_Handler
End Sub
