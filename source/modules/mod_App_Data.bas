Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.08
'
' Description:  application data related functions & procedures
'
' Source/date:  Bonnie Campbell, April 4, 2018
' Revisions:    BLC - 4/4/2018 - 1.00 - initial version
'               BLC - 4/9/2018 - 1.01 - added CheckTagStatus
'               BLC - 4/19/2018 - 1.02 - add CurrDb property (normally resides in framework)
'                                        added colors
'               BLC - 4/21/2018  - 1.03 - revised VaidateDBH condense logic
'               BLC - 5/24/2018 - 1.04 - removed CurrDb property (added framework
'                                        mod_Db module where it normally resides)
'                                        added DB_SYS_TABLES, APP_SYS_TABLES (normally in framework)
'               BLC - 8/27/2019 - 1.05 - added lngLtBlue, enabled lngLtGray
'               BLC - 6/29/2020 - 1.06 - ValidDBH() changed sapling IF statement to include equivalent DBH check vs sapling habit alone
'                                        Added GetEquivDBH() for avoiding Sapling DBH popups when subforms don't populate before check
'                                        Added TruncateNumber() - shift to framework module later
'               BLC - 8/2/2020  - 1.07 - ValidDBH() adjusted to use ABS() for DBH comparison instead of OR
'               BLC - 8/6/2020  - 1.08 - ValidDBH() - accommodate "N/A" values when no prior sampling DBH values exist
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
    'Public Const lngWhite As Long = 16777215    '?RGB(255,255,255) #FFFFFF
    'Public Const lngYellow As Long = 65535      '?RGB(255,255,0) #FFFF00
    Public Const lngLtYellow As Long = 14745599 '?RGB(255,255,224) #FFFFE0
    'Public Const lngGray As Long = 8224125      '?RGB(125, 125, 125)
'    Public Const lngLtGray As Long = 13882323   '?RGB(211, 211, 211)
'    Public Const lngGray50 As Long = 8355711    '?RGB(127,127,127) Text 1, Lighter 50% #7F7F7F Gray50
    'Public Const lngLime As Long = 6750105      '?RGB(153, 255, 102) #99FF66
'    Public Const lngBlue As Long = 16711680     '?RGB(0, 0, 255) #0000FF
'    Public Const lngBlack As Long = 0           '?RGB(0,0,0) #000000
'    Public Const lngRed As Long = 255           '?RGB(255,0,0) #FF0000
'    Public Const lngGreen As Long = 65280       '?RGB(0,255,0) #00FF00
'    Public Const lngLtBlue As Long = 16777164   '?RGB(204,255,255) #CCFFFF
    Public Const lngPink As Long = 10582263     '?RGB(247,120,161) #F778A1 carnation red
'    Public Const lngLime As Long = 6750105      '?RGB(153, 255, 102) #99FF66
'    Public Const lngLtLime As Long = 6750156    '?RGB(204,255,102) #CCFF66

'-----------------'-----------------------------
' RETIRED - 7/1/2020 - covered in fw_mod_Math
'----------------------------------------------
'Public Const PI As Single = 3.1415            'pi value

'normally in framework
'----------------------------------------------
' RETIRED - 7/1/2020 - covered in fw_mod_App_Settings
'----------------------------------------------
'Public DB_SYS_TABLES As Variant
'Public APP_SYS_TABLES As Variant

Public SWITCHBOARD As Form

' ---------------------------------
'  Database-wide Properties
' ---------------------------------

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
    frmTag.Controls("cbxTagStatus").backcolor = lngWhite
    
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
                frmTag.Controls("cbxTagStatus").backcolor = lngYellow
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
'               EventDate - current sampling event date (date)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 4, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version
'   BLC - 4/18/2018 - renamed txtTag > tbxTag
'   BLC - 4/19/2018 - revised to use CurrDb & accept Sapling/Tree to determine value
'   BLC - 4/21/2018 - revise & condense logic
'   BLC - 6/29/2020 - changed sapling IF statement to include equivalent DBH check vs sapling habit alone
'   BLC - 8/2/2020  - adjusted to use ABS() for DBH comparison instead of OR
'   BLC - 8/6/2020  - accommodate "N/A" values when no prior sampling DBH values exist
'   BLC - 8/7/2020  - adjust to include current sampling event date to assure comparison w/ prior sampling
'                     event date's DBH values
' ---------------------------------
Public Function ValidDBH(Habit As String, EventDate As Date) As Boolean 'fsub_Sapling_DBH_Exit(Cancel As Integer)
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

    Select Case Habit
        Case "Sapling", "Tree"
            'proceed
        Case Else
            GoTo Exit_Handler
    End Select

'   Me.Refresh
    
    Set db = CurrDb
        
    'fetch tree/sapling form names
    frmDataName = Replace("fsub_Habit_Data", "Habit", Habit)
    frmTagName = Replace("fsub_Tag_Habit", "Habit", Habit)
    'dbh form
    frmDBHName = Replace("fsub_HABIT_DBH", "HABIT", Habit)
    
    'unhighlight DBH Double Checked as default
    With Forms!frm_Events.Form.Controls(frmDataName).Form
        .Controls("lblDBHCheck").forecolor = lngBlack
        '.Controls("tbxHighlightChk").Visible = False
        '.Controls("tbxComments").BackColor = lngWhite
    End With
    
    strLocID = Forms!frm_Events!txtLocation_ID
    
    'skip validation if value is "N/A"
    If Forms!frm_Events.Form.Controls(frmDataName).Form.Controls(frmTagName).Controls("tbxTag") = "N/A" Then
        Exit Function
    End If
    
    'intTag = Forms!frm_Events!fsub_Sapling_Data!fsub_Tag_Sapling!tbxTag
    intTag = Forms!frm_Events.Form.Controls(frmDataName).Form.Controls(frmTagName).Controls("tbxTag")
    
    strEquivDBHCalc = "Round((((Sum(" & PI & "*((IIf([Live]=True,[DBH],0))/2)^2))*(1/" & PI & "))^0.5)*2,6)"
    
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
            & "FROM ((((tbl_Locations l " _
            & "INNER JOIN tbl_Events e ON l.Location_ID = e.Location_ID) " _
            & "INNER JOIN tbl_HABIT_Data sd ON e.Event_ID = sd.Event_ID) " _
            & "INNER JOIN tbl_Tags t ON sd.Tag_ID = t.Tag_ID) " _
            & "INNER JOIN tbl_HABIT_DBH sbh ON sd.HABIT_Data_ID = sbh.HABIT_Data_ID) " _
            & "WHERE e.Event_Date <=#" & EventDate & "# " _
            & "GROUP BY l.Location_ID, e.Event_ID, l.Admin_Unit_Code, " _
            & "l.Subunit_Code, e.Event_Date, t.Tag " _
            & "HAVING l.Location_ID = """ & strLocID & """ " _
            & "AND t.Tag = " & intTag & " " _
            & "ORDER BY e.Event_Date;"

    strSQL = Replace(strSQL, "HABIT", Habit)

Debug.Print "DBH_mod_App_Data: " & strSQL
    
    'use usys_temp_qdf
    Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    Set rs = db.OpenRecordset("usys_temp_qdf")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        
        ' ------------------------------------
        '  validate if there are DBH records
        ' ------------------------------------
        'no records (Saplings or Trees) --> validate OK since there
        If rs.RecordCount = 0 Then
            ValidDBH = True
            Exit Function
        End If
        
        'one record means only the current event has DBH values so there is nothing to validate against
        
        'more than one record (Trees)
        If rs.RecordCount > 1 And Habit = "Tree" Then
        
            CurrentDBH = rs![EquivDBH]
            rs.MovePrevious
            PastDBH = rs![EquivDBH]
                        
            ' +/- 4cm threshold check
            If Abs(CurrentDBH - PastDBH) >= 4 Then
            
                 'exceeds +/- 4cm threshold
                MsgBox "Warning...change in DBH exceeds +/- 4cm. " _
                    & vbCrLf & "Please check DBH values.", _
                    vbExclamation, "NCRN Vegetation Monitoring > Suspect DBH"
                
                IsValid = False
                
            End If
        
        'any records (Saplings)
        ElseIf rs.RecordCount >= 1 Then
                
            'refresh to update tbxEquivDBH BEFORE ValidDBH check
            '--> results in Error #2455 - invalid reference to property form/report
            'Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH.Requery
                
            'saplings DBH > = 1 (minimum threshold) check
            'avoid error #2455 - entered an expression w/ an invalid reference to the property form/report
            '   by checking if fsub_Sapling_DBH is loaded first, if not bypass by setting 1 < 1 which
            '   is not true so the section is skipped
'            If Habit = "Sapling" And _
'                IIf(IsLoaded("fsub_Sapling_DBH") = True, Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!tbxEquivDBH, 1) < 1 Then

                'Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!tbxEquivDBH < 1 Then

            If Habit = "Sapling" Then
            
                'nest IF since Tree doesn't have fsub_Sapling_DBH
                'avoids error #2455 - you have entered an invalid reference to the property Form/Report
                If Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!tbxEquivDBH < 1 Then
                        
                    MsgBox "Saplings must have a minimum DBH of 1.0. " _
                            & "Please check sapling DBH values.", _
                        vbExclamation, "NCRN Vegetation Monitoring > Invalid DBH"
                        
                    IsValid = False
                End If
                
            End If
        End If
        
        'highlight & set focus if not valid & DBH records exist
        If IsValid = False And rs.RecordCount > 0 Then
            
            'highlight DBH Double Checked
            With Forms!frm_Events.Form.Controls(frmDataName).Form
                '.Controls("lblDBHCheck").Visible = True
                '.Controls("chkDBHCheck").Visible = True
                .Controls("tbxHighlightChk").visible = True
                .Controls("lblDBHCheck").forecolor = lngRed
                .Controls("tbxHighlightChk").backcolor = lngLtYellow
                .Controls("tbxComments").backcolor = lngYellow
            
                'set focus
                .Controls(frmDBHName).Form.Controls("tbxDBH").SetFocus
            
            End With
        
            'set focus
'            Forms!frm_Events.Form.Controls(frmDataName).Form.Controls(frmDBHName).Form.Controls("tbxDBH").SetFocus
        
        End If
    End If
        
    ValidDBH = IsValid

Exit_Handler:
    'cleanup
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
' SUB:          GetEquivDBH
' Description:  retrieve equivalent DBH value for a given sapling
' Assumptions:  -
' Parameters:   DataID - sapling identifier (typically GUID, set as variant)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 30, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/30/2020 - initial version
' ---------------------------------
Public Function GetEquivDBH(DataID As Variant) As Double
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim qdf As QueryDef
    Dim rs As DAO.Recordset
    
    'default
    GetEquivDBH = 0
    
    'ensure data ID is not empty
    If Len(DataID) > 0 Then
        strSQL = "SELECT sdbh.Sapling_Data_ID, sdbh.Live, " _
                    & "(" _
                    & "(SUM(3.1415*(sdbh.DBH/2)^2)" _
                    & "*(1/3.1415))^.5" _
                    & ")*2 AS EquivDBH " _
                    & "FROM tbl_Sapling_DBH sdbh " _
                    & "WHERE sdbh.Sapling_Data_ID = " & DataID & " " _
                    & "GROUP BY sdbh.Sapling_Data_ID,sdbh.Live"
    Debug.Print strSQL
    
        'use usys_temp_qdf
        Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
        qdf.SQL = strSQL
        
        Set rs = CurrDb.OpenRecordset("usys_temp_qdf")
        
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveLast
        
            If rs.RecordCount = 1 Then
                'valid
                GetEquivDBH = rs("EquivDBH")
            End If
        End If
    End If
    
Debug.Print GetEquivDBH

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetEquivDBH[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetPriorDBH
' Description:  retrieve the tag's previous DBH value
'               (previous DBH value is for ALL stems and for the sampling event BEFORE the current one)
' Assumptions:  -
' Parameters:   DataID - tree or sapling data identifier (string)
'               VegType - tree or sapling (string)
'               TagID - tag identifier (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 16, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/16/2018 - initial version
'   BLC - 4/19/2018 - revise to use CurrDb
' ---------------------------------
Public Function GetPriorDBH(DataID As String, VegType As String, TagID As String) As Double
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim tblName As String
    Dim fldName As String
    Dim tblDBHName As String
    
    tblName = "tbl_" & VegType & "_Data"
    fldName = VegType & "_Data_ID"
    tblDBHName = "tbl_" & VegType & "_DBH"
    
'    fails --> aggregate Max in WHERE clause
'    strSQL = "SELECT DBH FROM " & tblName & _
'             "WHERE " & fldName & _
'             "= '" & DataID & _
'             "' AND Max(Updated_Date);"
    
'    fails --> only gives last DBH not the sum of all stem dbhs and only for last sampling event
'    strSQL = "SELECT TOP 1 DBH FROM " & tblName & " " & _
'             "WHERE " & fldName & _
'             "= '" & DataID & "' " & _
'             "ORDER BY Updated_Date;"
    
'    strSQL = "SELECT SUM(DBH) AS SumDBH" & _
'        "FROM tbl_Tree_DBH dbh" & _
'        "WHERE" & _
'        "Tree_Data_ID IN (" & _
'        "SELECT Tree_Data_ID FROM tbl_Tree_Data WHERE Tag_ID = TagID)" & _
'        "AND" & _
'        "Tree_Data_ID <> DataID" & _
'        "GROUP BY Tree_Data_ID, DBH" & _
'        "ORDER BY Updated_Date DESC;"

'genericized
'    strSQL = "SELECT SUM(DBH) AS SumDBH " & _
'        "FROM " & tblDBHName & " dbh " & _
'        "WHERE " & _
'        fldName & " IN (" & _
'        "SELECT " & fldName & " FROM " & tblName & " WHERE Tag_ID = " & TagID & ") " & _
'        "AND " & _
'        fldName & " <> " & DataID & " " & _
'        "GROUP BY " & fldName & ", DBH " & _
'        "ORDER BY Updated_Date DESC;"
    
    strSQL = "SELECT SumDBH FROM " & _
        "(SELECT TOP 1 e.Event_Date, dbh." & fldName & ", SUM(dbh.DBH) AS SumDBH " & _
        "FROM ((tbl_Events e " & _
        "INNER JOIN " & tblName & " d ON d.Event_ID = e.Event_ID) " & _
        "INNER JOIN " & tblDBHName & " dbh ON dbh." & fldName & " = d." & fldName & ") " & _
        "WHERE " & _
        "d." & fldName & " <> " & DataID & " " & _
        "AND " & _
        "d.Tag_ID = " & TagID & " " & _
"AND YEAR(e.Event_Date) < " & _
"(" & _
"SELECT YEAR(ee.Event_Date) " & _
"FROM " & _
"(tbl_Events ee " & _
"INNER JOIN " & tblName & " dd ON dd.Event_ID = ee.Event_ID) " & _
"WHERE dd." & fldName & " = " & DataID & " " & _
") " & _
        "GROUP BY dbh." & fldName & ", e.Event_Date " & " " & _
        "ORDER BY e.Event_Date DESC);"
    
 Debug.Print strSQL
    
    'use usys_temp_qdf
    Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    Set rs = CurrDb.OpenRecordset("usys_temp_qdf")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
    
        If rs.RecordCount = 1 Then
            'valid
            GetPriorDBH = rs("SumDBH") 'rs("DBH")
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

' ---------------------------------
' FUNCTION:     GetPriorDBHStems
' Description:  retrieve the tag's previous event stems
'               (previous DBH stems returns ALL stems for the sampling event BEFORE the current one)
' Assumptions:  -
' Parameters:   DataID - tree or sapling data identifier (string)
'               VegType - tree or sapling (string)
'               TagID - tag identifier (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 8/2/2020 - initial version
' ---------------------------------
Public Function GetPriorDBHStems(DataID As String, VegType As String, TagID As String) As DAO.Recordset
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim tblName As String
    Dim fldName As String
    Dim tblDBHName As String
    
    tblName = "tbl_" & VegType & "_Data"
    fldName = VegType & "_Data_ID"
    tblDBHName = "tbl_" & VegType & "_DBH"
    
    strSQL = "SELECT * FROM " & _
        "(SELECT e.Event_Date, dbh." & fldName & ", dbh.DBH " & _
        "FROM ((tbl_Events e " & _
        "INNER JOIN " & tblName & " d ON d.Event_ID = e.Event_ID) " & _
        "INNER JOIN " & tblDBHName & " dbh ON dbh." & fldName & " = d." & fldName & ") " & _
        "WHERE " & _
        "d." & fldName & " <> " & DataID & " " & _
        "AND " & _
        "d.Tag_ID = " & TagID & " " & _
"AND YEAR(e.Event_Date) < " & _
"(" & _
"SELECT YEAR(ee.Event_Date) " & _
"FROM " & _
"(tbl_Events ee " & _
"INNER JOIN " & tblName & " dd ON dd.Event_ID = ee.Event_ID) " & _
"WHERE dd." & fldName & " = " & DataID & " " & _
") " & _
        "GROUP BY dbh." & fldName & ", e.Event_Date, dbh.DBH " & " " & _
        "ORDER BY e.Event_Date DESC);"
    
 Debug.Print strSQL
    
    'use usys_temp_qdf
    Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    Set rs = CurrDb.OpenRecordset("usys_temp_qdf")
    
    With rs
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveLast
        
            If rs.RecordCount > 0 Then
                'valid
                Set GetPriorDBHStems = rs
            Else
                'invalid
                'GetPriorDBHStems = 99
                'GoTo Exit_Handler
            End If
        
        Else
            'add a default
            Dim rsDefault As DAO.Recordset
            'Set rsDefault = CurrDb.OpenRecordset("rsDefault")
            strSQL = "SELECT DISTINCT Now() AS Event_Date, '' AS " & fldName & ", 'NoDBH' AS DBH " & _
                   "FROM tbl_Events;"
            
            Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
            qdf.SQL = strSQL
            
            Set rsDefault = CurrDb.OpenRecordset("usys_temp_qdf")
            
            Debug.Print strSQL
            Debug.Print rsDefault.RecordCount
            
            Set GetPriorDBHStems = rsDefault
        End If
    End With
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetPriorDBHStems[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function


' ---------------------------------
' FUNCTION:     GetDBHCheck
' Description:  retrieve DBH value
' Assumptions:  -
' Parameters:   DataID - tag identifier (string)
'               Habit - tree or sapling (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/16/2018 - initial version
'   BLC - 4/19/2018 - revise to use CurrDb
'   BLC - 8/6/2020 - revise to point to tbl_HABIT_Data vs tbl_HABIT_DBH for DBH_Check
' ---------------------------------
Public Function GetDBHCheck(DataID As String, Habit As String) As Byte
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim tblName As String
    Dim fldName As String
    
    tblName = "tbl_" & Habit & "_Data"
    fldName = Habit & "_Data_ID"
    
    strSQL = "SELECT TOP 1 DBH_Check FROM " & tblName & " " & _
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
            GetDBHCheck = rs("DBH_Check")
        Else
            'invalid
            GetDBHCheck = 0
            'GoTo Exit_Handler
        End If
    
    Else
        GetDBHCheck = 0
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDBHCheck[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetDBHCheck
' Description:  retrieve DBH value
' Assumptions:  DBH_Check is a byte field (0 - false, 1 - true)
'               so checkbox values must be converted from -1 = true to 1 = true
' Parameters:   DataID - tag identifier (string)
'               Habit - tree or sapling (string)
'               chk - checkbox value (boolean)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/16/2018 - initial version
'   BLC - 4/19/2018 - revise to use CurrDb
' ---------------------------------
Public Function SetDBHCheck(DataID As String, Habit As String, chk As Boolean) As Byte
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim strSQL As String
    Dim tblName As String
    Dim fldName As String
    
    tblName = "tbl_" & Habit & "_Data"
    fldName = Habit & "_Data_ID"
    
    strSQL = "SELECT * FROM " & tblName & " " & _
             "WHERE " & fldName & _
             "= '" & DataID & "';"
    
    'use usys_temp_qdf
    Set qdf = CurrDb.QueryDefs("usys_temp_qdf")
    qdf.SQL = strSQL
    
    Set rs = CurrDb.OpenRecordset("usys_temp_qdf")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
    
        If rs.RecordCount = 1 Then
            
            'update the record
            With rs
                .Edit
                !DBH_Check = Abs(chk)
                !Updated_Date = Now
                .Update
                
                SetDBHCheck = True
            End With
            
        Else
            SetDBHCheck = False
            GoTo Exit_Handler
        End If
    
    Else
        SetDBHCheck = False
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetDBHCheck[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     MakeTreeStemList
' Description:  Collapse all tree stems into a single field (for event report)
' Assumptions:  -
' Parameters:   EventID - event identifier (string)
'               TreeDataID - tree identifier (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman, August 21, 2006
' Adapted:      Bonnie Campbell, May 14, 2018
' Revisions:
'   MEL - 8/21/2006 - initial version
'   BLC - 5/14/2018 - move from mod_AppSpecific > mod_App_Data
' ---------------------------------
Public Function MakeTreeStemList(EventID As String, TreeDataID As String) As String
On Error GoTo Err_Handler

    'Collapse all tree stems into a single field   mel 8/21/06
    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim strStemList As String
    Dim strStemListLive As String
    Dim strStemListDead As String
    
    strSQL = "SELECT d.DBH, d.Live, td.Event_ID, td.Tree_Data_ID " _
            & "FROM tbl_Tree_Data td " _
            & "INNER JOIN tbl_Tree_DBH d ON td.Tree_Data_ID = d.Tree_Data_ID " _
            & "WHERE td.Event_ID= """ & EventID & """ " _
            & "AND td.Tree_Data_ID= """ & TreeDataID & """;"

    Set rs = CurrentDb.OpenRecordset(strSQL)

    Do Until rs.EOF
        If rs!Live = True Then
            strStemListLive = strStemListLive & ", " & Format(rs!DBH, "#0.0")
        Else
            strStemListDead = strStemListDead & ", " & Format(rs!DBH, "#0.0")
        End If
        rs.MoveNext
    Loop

    strStemListLive = mid(strStemListLive, 3)
    strStemListDead = mid(strStemListDead, 3)
    strStemList = "L: " & strStemListLive & " D: " & strStemListDead
    
    MakeTreeStemList = strStemList

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MakeTreeStemList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     MakeSaplingStemList
' Description:  Collapse all sapling stems into a single field (for event report)
' Assumptions:  -
' Parameters:   EventID - event identifier (string)
'               SaplingDataID - sapling identifier (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman, August 21, 2006
' Adapted:      Bonnie Campbell, May 14, 2018
' Revisions:
'   MEL - 8/21/2006 - initial version
'   BLC - 5/14/2018 - move from mod_AppSpecific > mod_App_Data
' ---------------------------------
Public Function MakeSaplingStemList(EventID As String, SaplingDataID As String) As String
On Error GoTo Err_Handler

    'Collapse all sapling stems into a single field   mel 8/21/06
    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim strStemList As String
    Dim strStemListLive As String
    Dim strStemListDead As String
    
    strSQL = "SELECT d.DBH, d.Live, sd.Event_ID, sd.Sapling_Data_ID " _
        & "FROM tbl_Sapling_Data sd " _
        & "INNER JOIN tbl_Sapling_DBH d ON sd.Sapling_Data_ID = d.Sapling_Data_ID " _
        & "WHERE sd.Event_ID= """ & EventID & """ " _
        & "AND sd.Sapling_Data_ID= """ & SaplingDataID & """;"
    
    Set rs = CurrentDb.OpenRecordset(strSQL)

    Do Until rs.EOF
        If rs!Live = True Then
            strStemListLive = strStemListLive & ", " & Format(rs!DBH, "#0.0")
        Else
            strStemListDead = strStemListDead & ", " & Format(rs!DBH, "#0.0")
        End If
        rs.MoveNext
    Loop

    strStemListLive = mid(strStemListLive, 3)
    strStemListDead = mid(strStemListDead, 3)
    strStemList = "L: " & strStemListLive & " D: " & strStemListDead
    
    MakeSaplingStemList = strStemList

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MakeSaplingStemList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     MakeStemList
' Description:  Collapse all tree/sapling stems into a single field (for event report)
' Assumptions:  -
' Parameters:   Mode - tree or stem (string)
'               EventID - event identifier (string)
'               SaplingDataID - sapling identifier (string)
' Returns:      -
' Throws:       none
' References:   -
' Used in:      Event report - tree & sapling subforms
' Source/date:  Mark Lehman, August 21, 2006
' Adapted:      Bonnie Campbell, May 14, 2018 from MakeTreeStemList/MakeSaplingStemList
' Revisions:
'   MEL - 8/21/2006 - initial version (MakeTreeStemList/MakeSaplingStemList)
'   BLC - 5/14/2018 - move from mod_AppSpecific > mod_App_Data & revise to accommodate both
'                     trees & saplings
' ---------------------------------
Public Function MakeStemList(mode As String, EventID As String, DataID As String) As String
On Error GoTo Err_Handler

    'Collapse all sapling stems into a single field   mel 8/21/06
    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim strStemList As String
    Dim strStemListLive As String
    Dim strStemListDead As String
    
    Select Case mode
        Case "Sapling"
            strSQL = "SELECT d.DBH, d.Live, sd.Event_ID, sd.Sapling_Data_ID " _
                & "FROM tbl_Sapling_Data sd " _
                & "INNER JOIN tbl_Sapling_DBH d ON sd.Sapling_Data_ID = d.Sapling_Data_ID " _
                & "WHERE sd.Event_ID= """ & EventID & """ " _
                & "AND sd.Sapling_Data_ID= """ & DataID & """;"
        
        Case "Tree"
            strSQL = "SELECT d.DBH, d.Live, td.Event_ID, td.Tree_Data_ID " _
                    & "FROM tbl_Tree_Data td " _
                    & "INNER JOIN tbl_Tree_DBH d ON td.Tree_Data_ID = d.Tree_Data_ID " _
                    & "WHERE td.Event_ID= """ & EventID & """ " _
                    & "AND td.Tree_Data_ID= """ & DataID & """;"
    End Select
    
    Set rs = CurrentDb.OpenRecordset(strSQL)

    Do Until rs.EOF
        If rs!Live = True Then
            strStemListLive = strStemListLive & ", " & Format(rs!DBH, "#0.0")
        Else
            strStemListDead = strStemListDead & ", " & Format(rs!DBH, "#0.0")
        End If
        rs.MoveNext
    Loop

    strStemListLive = mid(strStemListLive, 3)
    strStemListDead = mid(strStemListDead, 3)
    strStemList = "L: " & strStemListLive & " D: " & strStemListDead
    
    MakeStemList = strStemList

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MakeStemList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     MakeLiveFlag
' Description:  Collapse all tree/sapling stem live/dead flags into a single value (for event report)
' Assumptions:  Live/Dead flags should match with sapling status
'
' Parameters:   Mode - tree or sapling (string)
'               EventID - event identifier (string)
'               SaplingDataID - sapling identifier (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 14, 2018
' Adapted:      -
' Revisions:
'   BLC - 5/14/2018 - initial version
' ---------------------------------
Public Function MakeLiveFlag(mode As String, EventID As String, DataID As String) As String
On Error GoTo Err_Handler

    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim iStemsLive As Integer
    'Dim iStemsDead As Integer
    
    Select Case mode
        Case "Sapling"
            strSQL = "SELECT d.DBH, d.Live, sd.Event_ID, sd.Sapling_Data_ID " _
                & "FROM tbl_Sapling_Data sd " _
                & "INNER JOIN tbl_Sapling_DBH d ON sd.Sapling_Data_ID = d.Sapling_Data_ID " _
                & "WHERE sd.Event_ID= """ & EventID & """ " _
                & "AND sd.Sapling_Data_ID= """ & DataID & """;"
        
        Case "Tree"
            strSQL = "SELECT d.DBH, d.Live, td.Event_ID, td.Tree_Data_ID " _
                    & "FROM tbl_Tree_Data td " _
                    & "INNER JOIN tbl_Tree_DBH d ON td.Tree_Data_ID = d.Tree_Data_ID " _
                    & "WHERE td.Event_ID= """ & EventID & """ " _
                    & "AND td.Tree_Data_ID= """ & DataID & """;"
    End Select
    
    Set rs = CurrentDb.OpenRecordset(strSQL)

    Do Until rs.EOF
    
        iStemsLive = iStemsLive + Abs(CInt(rs!Live))
    
'        If rs!Live = True Then
'            iStemsLive = iStemsLive + 1
'        Else
'            iStemsDead = iStemsDead + 1
'        End If
        rs.MoveNext
    
    Loop
    
    MakeLiveFlag = iStemsLive

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MakeLiveFlag[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     TruncateNumber
' Description:  Truncates a number to a desired # of decimal places
' Assumptions:  -
' Parameters:   InputNumber - number to be truncated (double)
'               DecimalPlaces - number of decimal places to return (int)
' Returns:      InputNumber truncated to # of DecimalPlaces desired (double)
' Throws:       none
' References:
'   Makah, March 23, 2015
'   https://stackoverflow.com/questions/11347704/truncating-double-with-vba-in-excel
' Source/date:  Bonnie Campbell, June 30, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/30/2020 - initial version
' ---------------------------------
Public Function TruncateNumber(ByVal InputNumber As Double, ByVal DecimalPlaces As Integer) As Double
On Error GoTo Err_Handler

    TruncateNumber = Int(InputNumber * (10 ^ DecimalPlaces)) / (10 ^ DecimalPlaces)

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TruncateNumber[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function