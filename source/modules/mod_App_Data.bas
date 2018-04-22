Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.01
'
' Description:  application data related functions & procedures
'
' Source/date:  Bonnie Campbell, April 4, 2018
' Revisions:    BLC - 4/4/2018 - 1.00 - initial version
'               BLC - 4/9/2018 - 1.01 - added CheckTagStatus
'               BLC - 4/19/2018 - 1.02 - add CurrDb property (normally resides in framework)
'                                        added colors
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
Public Const lngWhite As Long = 16777215    '?RGB(255,255,255) #FFFFFF
Public Const lngYellow As Long = 65535      '?RGB(255,255,0) #FFFF00
Public Const lngLtYellow As Long = 14745599 '?RGB(255,255,224) #FFFFE0
Public Const lngGray As Long = 8224125      '?RGB(125, 125, 125)
Public Const lngLtGray As Long = 13882323   '?RGB(211, 211, 211)
Public Const lngGray50 As Long = 8355711    '?RGB(127,127,127) Text 1, Lighter 50% #7F7F7F Gray50
Public Const lngLime As Long = 6750105      '?RGB(153, 255, 102) #99FF66
Public Const lngBlue As Long = 16711680     '?RGB(0, 0, 255) #0000FF
Public Const lngBlack As Long = 0           '?RGB(0,0,0) #000000
Public Const lngRed As Long = 255           '?RGB(255,0,0) #FF0000
Public Const lngGreen As Long = 65280       '?RGB(0,255,0) #00FF00

Public Const pi As Single = 3.1415            'pi value

' ---------------------------------
'  Database-wide Properties
' ---------------------------------

' ---------------------------------
' PROPERTY:     CurrDb
' Description:  Gets a single instance of the current db to avoid multiple calls
'               to CurrentDb which can yield to Error 3048 "Cannot open any more databases" errors
'               due to multiple open db
' Parameters:   -
' Returns:      current database object
' Throws:       -
' References:   -
' Source/date:  Dirk Goldgar, MS Access MVP - May 22, 2013
'   http://social.msdn.microsoft.com/Forums/office/en-US/9993d229-8a00-4a59-a796-dfa2dad505bc/cannot-open-any-more-databases?forum=accessdev
'   Michael Kaplan via Darrel H. Burns, February 8, 2011
'   https://social.msdn.microsoft.com/Forums/office/en-US/7ea9506f-5e91-4896-80b9-6712762388ea/currentdbtabledefs-vs-dbtabledefs-object-invalid-or-not-set-error?forum=accessdev
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Revisions:    BLC, 7/23/2014 - initial version
'               BLC, 10/5/2017 - combined with DbCurrent from mod_SQL
' ---------------------------------
Private m_db As DAO.Database

Public Property Get CurrDb() As DAO.Database

    If (m_db Is Nothing) Then
        Set m_db = CurrentDb
    End If

    Set CurrDb = m_db

End Property

' ----------------
'  Methods
' ----------------
' ---------------------------------
' SUB:          CheckTagStatus
' Description:  compare tag status w/ tree/sapling status
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/9/2018 - initial version
' ---------------------------------
Public Sub CheckTagStatus(StatusType As String)
On Error GoTo Err_Handler

    Dim frm As Form, frmTag As Form
    Dim cbx As ComboBox
    Dim frmName As String, fsubName As String, fsubTagName As String
    Dim cbxName As String
    
    Select Case StatusType
        Case "Sapling"
        
        Case "Tree"
        
        
        Case Else
            GoTo Exit_Handler
    End Select
    
    frmName = "frm_Events"
    fsubName = "fsub_" & StatusType & "_Data"
    fsubTagName = "fsub_Tag_" & StatusType
    cbxName = "cbx" & StatusType & "Status"
     
    Set frm = Forms(frmName).Form.Controls(fsubName).Form
    Set frmTag = frm.Form.Controls(fsubTagName).Form
    Set cbx = frm.Controls(cbxName)
    
    'default
    frmTag.Controls("cbxTagStatus").BackColor = lngWhite
    
    'dead statuses only
    If Left(cbx, 4) = "Dead" Then
    
        'compare w/ tag status
        Select Case frmTag.Controls("cbxTagStatus")
        
            Case Is = Null
                'set value
                frmTag.Controls("cbxTagStatus") = "Retired (In Office)"
            Case "Retired (In Office)"
                'do nothing
            Case Else
                'highlight
                frmTag.Controls("cbxTagStatus").BackColor = lngYellow
        End Select
            
    End If
    
'    'Tree status = Dead* ?
'    ' --> trigger tag status = RIO (Retired (In Office))
'    If Left(cbxTreeStatus, 4) = "Dead" Then
'
''Debug.Print "tag status = " & Me.fsub_Tag_Tree.Controls("cbxTagStatus")

'        Select Case fsub_Tag_Tree.Controls("cbxTagStatus")
'         Case Is <> "Retired (In Office)"
'            Me.fsub_Tag_Tree.Controls("cbxTagStatus").BackColor = lngYellow
'
'         Case Is = Null
'Debug.Print "tag status = NULL " & Me.fsub_Tag_Tree.Controls("cbxTagStatus")
'                'set the value
'                fsub_Tag_Tree.Controls("cbxTagStatus") = "Retired (In Office)"
'         Case Else
'            'do nothing
'        End Select
'
'    Else
'
'        Me.fsub_Tag_Tree.Controls("cbxTagStatus").BackColor = lngWhite
'
'    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CheckTagStatus[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Functions
' ----------------

' ---------------------------------
' FUNCTION:     ValidDBH
' Description:  validates DBH value for invalid change
' Assumptions:  INVALID --> DBH value changes > +/- 4cm over prior year
'                       --> Sapling DBH values < 1cm
'                           (Minimum DBH for saplings = 1 cm)
' Parameters:   Habit - tree or sapling (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 4, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version
'   BLC - 4/18/2018 - renamed txtTag > tbxTag
'   BLC - 4/19/2018 - revised to use CurrDb & accept Sapling/Tree to determine value
' ---------------------------------
Public Function ValidDBH(Habit As String) As Boolean 'fsub_Sapling_DBH_Exit(Cancel As Integer)
On Error GoTo Err_Handler

    Dim IsValid As Boolean
    Dim strLocID As String
    Dim intTag As Integer
    Dim CurrentDBH As Variant
    Dim PastDBH As Variant
    Dim strSQL As String
    Dim strEquivDBHCalc As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim frmDataName As String
    Dim frmTagName As String
    Dim frmDBHName As String
    
    'default
    IsValid = True

'   Me.Refresh
    
    Set db = CurrDb
    
'    'Check to see if the temporary query exists and if it does delete it.
'
'    If fxnQueryExists("_qCOMPARE_DBH") Then
'        db.QueryDefs.Delete ("_qCOMPARE_DBH")
'    End If
        
    'fetch tree/sapling form names
    frmDataName = Replace("fsub_Habit_Data", "Habit", Habit)
    frmTagName = Replace("fsub_Tag_Habit", "Habit", Habit)
    
    'unhighlight DBH Double Checked as default
    With Forms!frm_Events.Form.Controls(frmDataName).Form
        .Controls("lblDBHCheck").ForeColor = lngBlack
        '.Controls("tbxHighlightChk").Visible = False
        .Controls("tbxComments").BackColor = lngWhite
    End With
    
    strLocID = Forms!frm_Events!txtLocation_ID
    
    'intTag = Forms!frm_Events!fsub_Sapling_Data!fsub_Tag_Sapling!tbxTag
    intTag = Forms!frm_Events.Form.Controls(frmDataName).Form.Controls(frmTagName).Controls("tbxTag")
    
    strEquivDBHCalc = "Round((((Sum(" & pi & "*((IIf([Live]=True,[DBH],0))/2)^2))*(1/" & pi & "))^0.5)*2,6)"
    
'    strSQL = "SELECT tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag, " _
'            & "Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2,6) AS EquivDBH " _
'            & "FROM ((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " _
'            & "INNER JOIN (tbl_Sapling_Data INNER JOIN tbl_Tags ON tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID) ON tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID) " _
'            & "INNER JOIN tbl_Sapling_DBH ON tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_DBH.Sapling_Data_ID " _
'            & "GROUP BY tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag " _
'            & "HAVING (((tbl_Locations.Location_ID) = """ & strLocID & """) And ((tbl_Tags.Tag) = " & intTag & ")) " _
'            & "ORDER BY tbl_Events.Event_Date;"

    'generic SQL
    strSQL = "SELECT l.Location_ID, e.Event_ID, l.Admin_Unit_Code, " _
            & "l.Subunit_Code, e.Event_Date, t.Tag, " _
            & strEquivDBHCalc & " AS EquivDBH " _
            & "FROM ((tbl_Locations l " _
            & "INNER JOIN tbl_Events e ON l.Location_ID = e.Location_ID) " _
            & "INNER JOIN (tbl_HABIT_Data sd " _
            & "INNER JOIN tbl_Tags t ON sd.Tag_ID = t.Tag_ID) " _
            & "ON e.Event_ID = sd.Event_ID) " _
            & "INNER JOIN tbl_HABIT_DBH sbh ON sd.HABIT_Data_ID = sbh.HABIT_Data_ID " _
            & "GROUP BY l.Location_ID, e.Event_ID, l.Admin_Unit_Code, " _
            & "l.Subunit_Code, e.Event_Date, t.Tag " _
            & "HAVING (((l.Location_ID) = """ & strLocID & """) " _
            & "AND ((t.Tag) = " & intTag & ")) " _
            & "ORDER BY e.Event_Date;"

    strSQL = Replace(strSQL, "HABIT", Habit)

Debug.Print "DBH_mod_App_Data: " & strSQL

    'Dim qdf As DAO.QueryDef
    'Set qdf = db.CreateQueryDef("_qCOMPARE_DBH", strSQL)
    
    'use usys_temp_qdf
    Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    'Set rs = db.OpenRecordset("_qCOMPARE_DBH")
    Set rs = db.OpenRecordset("usys_temp_qdf")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        
        'validate if there are DBH records
        If rs.RecordCount > 1 Then
        
            CurrentDBH = rs![EquivDBH]
            rs.MovePrevious
            PastDBH = rs![EquivDBH]
            
            ' +/- 4cm threshold check
            If CurrentDBH - PastDBH >= 4 Or CurrentDBH - PastDBH <= -4 Then
    
'                'highlight DBH Double Checked
'                With Forms!frm_Events.Form.Controls(frmDataName).Form
'                    '.Controls("lblDBHCheck").Visible = True
'                    '.Controls("chkDBHCheck").Visible = True
'                    .Controls("tbxHighlightChk").Visible = True
'                    '.Controls("lblDBHCheck").ForeColor = lngRed
'                    .Controls("tbxHighlightChk").BackColor = lngLtYellow
'                    .Controls("tbxComments").BackColor = lngYellow
'                End With
            
                 'exceeds +/- 4cm threshold
                MsgBox "Warning...change in DBH exceeds +/- 4cm. Please check value.", _
                    vbExclamation, "NCRN Vegetation Monitoring"
                
                IsValid = False
                
            End If
        End If
    End If
    
    
    Select Case Habit
    
        Case "Sapling"
            'saplings DBH > = 1 (minimum threshold)
            If Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!tbxEquivDBH < 1 Then
            

            
                MsgBox "Saplings must have a minimum DBH of 1.0. Please address the issue"
                Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!tbxDBH.SetFocus
                IsValid = False
            End If
        
        Case "Tree"
            'do nothing
        Case Else
            GoTo Exit_Handler
    End Select
    
    'set focus if not valid
    If Not (IsValid = True) And _
        Forms!frm_Events.Form.Controls(frmDataName).Form.Recordset.RecordCount > 0 Then
        
        'highlight DBH Double Checked
        With Forms!frm_Events.Form.Controls(frmDataName).Form
            '.Controls("lblDBHCheck").Visible = True
            '.Controls("chkDBHCheck").Visible = True
            .Controls("tbxHighlightChk").Visible = True
            .Controls("lblDBHCheck").ForeColor = lngRed
            .Controls("tbxHighlightChk").BackColor = lngLtYellow
            .Controls("tbxComments").BackColor = lngYellow
        End With
        
        'dbh form
        frmDBHName = Replace("fsub_HABIT_DBH", "HABIT", Habit)
    
        Forms!frm_Events.Form.Controls(frmDataName).Form.Controls(frmDBHName).Form.Controls("tbxDBH").SetFocus
    End If
    
    ValidDBH = IsValid

Exit_Handler:
    'cleanup
    'DoCmd.DeleteObject acQuery, "_qCOMPARE_DBH"
    Set CurrentDBH = Nothing
    Set PastDBH = Nothing
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ValidDBH[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     CalcEquivDBH
' Description:  calculate equivalent DBH
' Assumptions:  -
' Parameters:   IsLive - whether the tree is live or dead (boolean)
'               DBH - DBH measurement (double)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 4, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version (currently unused)
' ---------------------------------
Public Function CalcEquivDBH(IsLive As Boolean, DBH As Double) As Double
On Error GoTo Err_Handler
    
    'EquivDBH = Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2,6)
    
    'CalcEquivDBH = Round((((Sum(pi * ((IIf(IsLive = True, DBH, 0)) / 2) ^ 2)) * (1 / pi)) ^ 0.5) * 2, 6)
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CalcEquivDBH[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetPriorDBH
' Description:  retrieve the tag's previous DBH value
' Assumptions:  -
' Parameters:   DataID - tag identifier (string)
'               VegType - tree or sapling (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 16, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/16/2018 - initial version
'   BLC - 4/19/2018 - revise to use CurrDb
' ---------------------------------
Public Function GetPriorDBH(DataID As String, VegType As String) As Double
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim tblName As String
    Dim fldName As String
    
    tblName = "tbl_" & VegType & "_DBH"
    fldName = VegType & "_Data_ID"
    
'    fails --> aggregate Max in WHERE clause
'    strSQL = "SELECT DBH FROM " & tblName & _
'             "WHERE " & fldName & _
'             "= '" & DataID & _
'             "' AND Max(Updated_Date);"
    
    strSQL = "SELECT TOP 1 DBH FROM " & tblName & " " & _
             "WHERE " & fldName & _
             "= '" & DataID & "' " & _
             "ORDER BY Updated_Date;"
    
    'use usys_temp_qdf
    Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    Set rs = CurrDb.OpenRecordset("usys_temp_qdf")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
    
        If rs.RecordCount = 1 Then
            'valid
            GetPriorDBH = rs("DBH")
        Else
            'invalid
            GetPriorDBH = 99
            'GoTo Exit_Handler
        End If
    
    Else
        GetPriorDBH = 0
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetPriorDBH[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function