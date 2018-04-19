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
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
Public Const lngWhite As Long = 16777215    '?RGB(255,255,255) #FFFFFF
Public Const lngYellow As Long = 65535      '?RGB(255,255,0) #FFFF00

Public Const pi As Single = 3.1415            'pi value

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
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 4, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version
'   BLC - 4/18/2018 - renamed txtTag > tbxTag
' ---------------------------------
Public Function ValidDBH() As Boolean 'fsub_Sapling_DBH_Exit(Cancel As Integer)
On Error GoTo Err_Handler

    Dim IsValid As Boolean
    
    'default
    IsValid = True

'   Me.Refresh
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    'Check to see if the temporary query exists and if it does delete it.
    
    If fxnQueryExists("_qCOMPARE_DBH") Then
        db.QueryDefs.Delete ("_qCOMPARE_DBH")
    End If
    
    Dim strLocID As String
    strLocID = Forms!frm_Events!txtLocation_ID
    
    Dim intTag As Integer
    intTag = Forms!frm_Events!fsub_Sapling_Data!fsub_Tag_Sapling!tbxTag
    
    Dim CurrentDBH As Variant
    Dim PastDBH As Variant
    
    Dim strSQL As String
    strSQL = "SELECT tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag, " _
            & "Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2,6) AS EquivDBH " _
            & "FROM ((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " _
            & "INNER JOIN (tbl_Sapling_Data INNER JOIN tbl_Tags ON tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID) ON tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID) " _
            & "INNER JOIN tbl_Sapling_DBH ON tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_DBH.Sapling_Data_ID " _
            & "GROUP BY tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag " _
            & "HAVING (((tbl_Locations.Location_ID) = """ & strLocID & """) And ((tbl_Tags.Tag) = " & intTag & ")) " _
            & "ORDER BY tbl_Events.Event_Date;"
    
    Dim qdf As DAO.QueryDef
    Set qdf = db.CreateQueryDef("_qCOMPARE_DBH", strSQL)
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("_qCOMPARE_DBH")
    
    rs.MoveLast
    If rs.RecordCount > 1 Then
    
        CurrentDBH = rs![EquivDBH]
        rs.MovePrevious
        PastDBH = rs![EquivDBH]
        
        If CurrentDBH - PastDBH >= 4 Or CurrentDBH - PastDBH <= -4 Then
            'MsgBox "Warning...change in DBH exceeds threshold. Please check value.", vbExclamation, "NCRN Vegetation Monitoring"
            'exceeds +/- 4cm threshold
            IsValid = False
        End If
    End If
    
    
    If Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!txtEquivDBH < 1 Then
        MsgBox "Saplings must have a minimum DBH of 1.0. Please address the issue"
        Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!txtDBH.SetFocus
        IsValid = False
    End If
    
    ValidDBH = IsValid

Exit_Handler:
    DoCmd.DeleteObject acQuery, "_qCOMPARE_DBH"
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
    Set qdf = CurrentDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    Set rs = CurrentDb.OpenRecordset("usys_temp_qdf")
    
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