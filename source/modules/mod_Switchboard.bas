' =================================
' MODULE:       basSwitchboard
' Description:  Standard module for functions associated with the application switchboard
'
'   The functions in this module require that the database contain the following two tables:
'
'   tsys_Link_Files:  Link_type (txt 50), Link_file_name (txt 100), Link_file_path (txt 255);
'       optional fields:  Link_description (txt 255).  [Link_type] should be 'Back-end data'
'       for the primary back-end database (in case of multiple back-ends).
'
'   tsys_Link_Tables:  Link_type (txt 50), Link_table (txt 100), Table_type (txt 50),
'       Description_text (txt 255).
'
' Source/date:  John R. Boetsch, May 2005; based on a similar implementation by Susan Huse,
'                   MonitoringSM.mdb v 7/28/2004
' Revisions:    JRB, May 2006

Option Compare Database
Option Explicit

' =================================
' FUNCTION:     fxnOpenDbChecks
' Description:  Checks the status of back-end connection and creates a backup upon
'               opening the database; triggered by the AutoExec macro
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   fxnFileExists, fxnSwitchboardIsOpen, fxnVerifyLinks, fxnMakeBackup
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
' Revisions:    JRB, May 24, 2006 - updated documentation, error traps, modified backup
'               strategy and added verification of individual table links
' =================================

Public Function fxnOpenDbChecks()
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSysTable As String
    Dim strDataFileName As String
    Dim strMissingFile As String
    Dim strErrorMsg As String
    Dim varConnected As Variant

    Set db = CurrentDb
    strSysTable = "tsys_Link_Files"     ' System table listing linked tables

    ' Verify that each linked database file is where it should be.
    '   Loops through multiple back-end files in case there is more than one

    ' Set the recordset to the system table
    Set rst = db.OpenRecordset(strSysTable, dbOpenTable, dbReadOnly)

    Do Until rst.EOF
        strDataFileName = rst![Link_file_path]
        If strDataFileName <> "" Then
            ' Set the connection status variable to TRUE if the file exists, otherwise FALSE
            varConnected = fxnFileExists(strDataFileName)
            ' If not connected, set the missing file string.  Note: if looping through
            '   multiple back-end files, the user will only be notified of one broken link
            If varConnected = False Then
                strMissingFile = strDataFileName
                ' Initialize the error message with the missing file string
                strErrorMsg = "Back-end database file(s) missing: " & vbCrLf & vbCrLf _
                    & strMissingFile & vbCrLf & vbCrLf & "You must update the data table " & _
                    "connections by selecting " & vbCrLf & "'Connect Data Tables' from " & _
                    "the menu before using the database." & vbCrLf & vbCrLf & _
                    "Would you like to fix the connection now?"
                ' Skip the routine for testing the individual table links
                GoTo Update_Routine
            End If
        End If
        rst.MoveNext
    Loop

    ' Check the status of individual table links, depending on application settings
    If fxnSwitchboardIsOpen Then
        If Forms!frm_Switchboard!chkVerifyOnStartup Then
            If fxnVerifyLinks = False Then
                varConnected = False
                ' Initialize the error message
                strErrorMsg = "You must update the data table connections by " _
                    & vbCrLf & "selecting 'Connect Data Tables' from the menu " _
                    & "before using the database." & vbCrLf & vbCrLf & _
                    "Would you like to fix the connection now?"
            End If
        End If
    End If

' --------------------------
Update_Routine:
    If varConnected = False Then
        If MsgBox(strErrorMsg, vbCritical + vbYesNo, "Update Data Table Connections") = vbYes Then
            ' Open the form to reconnect back-end tables
            DoCmd.OpenForm "frm_Connect_Tables"
        Else: Exit Function
        End If
    Else:
        ' Prompt for database backup, depending on system default settings
        '   and whether or not the back-end is properly connected
        If fxnSwitchboardIsOpen And varConnected Then
            If Forms!frm_Switchboard!chkBackupOnStartup Then fxnMakeBackup
        End If
    End If

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3011   ' System table not found
             MsgBox "Error #" & Err.Number & ":  Missing the following system table:" & _
                vbCrLf & vbCrLf & strSysTable & vbCrLf & vbCrLf & _
                "Please notify the database administrator before using this application.", _
                vbCritical, "System table error (fxnOpenDbChecks)"
        Case 3265   ' Field name in the tsys_Link_Files improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (fxnOpenDbChecks)"
        Case 94    ' Missing information in the tsys_Link_Files systems table
            MsgBox "Error #" & Err.Number & ":  Missing database path." & vbCrLf & _
                "Please notify the database administrator before using this application.", _
                vbCritical, "System table error (fxnOpenDbChecks)"
        Case 3061   ' Bad parameters for the SQL string
            MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
                "database administrator before using this application.", vbCritical, _
                "SQL String Error (fxnOpenDbChecks)"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnOpenDbChecks)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnSwitchboardIsOpen
' Description:  Indicates whether or not the switchboard form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnSwitchboardIsOpen() As Boolean
    On Error GoTo Err_Handler

    fxnSwitchboardIsOpen = False    ' Default in case of error

    Dim strSwitchboardName As String

    strSwitchboardName = "frm_Switchboard"

    If CurrentProject.AllForms(strSwitchboardName).IsLoaded = True Then
        If CurrentProject.AllForms(strSwitchboardName).CurrentView = 1 Then
            fxnSwitchboardIsOpen = True
        End If
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnSwitchboardIsOpen)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnFileExists
' Description:  Indicates whether or not the indicated file exists
' Parameters:   strPath as a string
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 8, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnFileExists(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnFileExists = False    ' Default in case of error

    Dim FS As Variant

    Set FS = CreateObject("Scripting.FileSystemObject")
    If FS.FileExists(strPath) Then fxnFileExists = True

Exit_Procedure:
    On Error Resume Next
    Set FS = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnFileExists)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnMakeBackup
' Description:  Creates a backup of the primary back-end data file
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   fxnSaveFile
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
' Revisions:    JRB, May 16, 2006 - updated documentation, error traps, modified date/time stamp
'               to be appended to the database name, changed strCopyFileName to a Variant to
'               accommodate nulls from the procedure call, changed overall backup strategy
'               SDK, Feb 26, 2007 - updated to loop through linked file table and backup every file
'               that has backup checkbox checked
' =================================

Public Function fxnMakeBackup()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim FS As Variant
Dim strSysTable As String
Dim strSourceFileName As String
Dim strCopyFileName As Variant
Dim strFileName As String
Dim strBackupDate As String

On Error GoTo Err_Handler

' Prompt the user to verify before backing up
If MsgBox("Would you like to make a backup copy of the data?", vbYesNo, _
    "Create Backup?") = vbNo Then
    GoTo Exit_Procedure
Else
    Set db = CurrentDb
    strSysTable = "tsys_Link_Files"     ' System table listing linked tables
    
    Set rst = db.OpenRecordset("SELECT Link_file_path FROM tsys_Link_Files WHERE Backup;", dbOpenForwardOnly)
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    strBackupDate = Format$(Now, "YYYYMMDD_HHNN")
    
    ' Loops through multiple back-end files in case there is more than one
    Do Until rst.EOF
        ' Verify that an .mdb is being specified
        If Right(rst!Link_file_path, 2) <> "db" Then
            MsgBox "Expected *.*db file, found " & strSourceFileName & vbCrLf & _
                    "File NOT copied.", vbCritical, "Error creating data backup"
            GoTo Exit_Procedure
        Else
            strFileName = Left(rst!Link_file_path, Len(rst!Link_file_path) - 6) & _
                "_" & strBackupDate & "_" & Forms![frm_Switchboard]![Entry_Role] & ".accdb"
            ' Open the save file dialog and update to the actual name given by the user
            strCopyFileName = fxnSaveFile(strFileName, "Access (*.*db)", "*.*db")
    
            If IsNull(strCopyFileName) Then
                ' User canceled save operation
                MsgBox "No Backup Made", vbOKOnly
            Else
                ' Perform the actual file copy
                FS.CopyFile rst!Link_file_path, strCopyFileName
                MsgBox "Backup file successfully created: " & vbCrLf & vbCrLf & _
                    strCopyFileName, vbOKOnly
            End If
        End If
    
        rst.MoveNext
    Loop


End If  ' End of initial user msgbox prompt

Exit_Procedure:
    On Error Resume Next
    Set FS = Nothing
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3011   ' System table not found
             MsgBox "Error #" & Err.Number & ":  Missing the following system table:" & _
                vbCrLf & vbCrLf & strSysTable & vbCrLf & vbCrLf & _
                "Please notify the database administrator before using this application.", _
                vbCritical, "System table error (fxnMakeBackup)"
        Case 3265   ' Field name in the tsys_Link_Files improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (fxnMakeBackup)"
        Case 94    ' Missing information in the tsys_Link_Files systems table
            MsgBox "Error #" & Err.Number & ":  Missing database path." & vbCrLf & _
                "Please notify the database administrator before using this application.", _
                vbCritical, "System table error (fxnMakeBackup)"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnMakeBackup)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnSaveFile
' Description:  Opens the save file dialog and returns the saved file name
' Parameters:   strFileName, strFileType, strFileExt - file name, type and extension
' Returns:      name of the saved file, or Null if user cancels
' Throws:       none
' References:   adhAddFilterItem, adhCommonFileOpenSave
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
' Revisions:    JRB, May 16, 2006 - updated documentation, error traps
' =================================

Public Function fxnSaveFile(strFileName As String, strFileType As String, _
    strFileExt As String) As Variant

    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim lngFlags As Long

    ' Use the save file dialog to interactively select the save file name and location
    strFilter = adhAddFilterItem(strFilter, strFileType, strFileExt)

    lngFlags = adhOFN_HIDEREADONLY Or adhOFN_OVERWRITEPROMPT Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR

    fxnSaveFile = adhCommonFileOpenSave( _
        OpenFile:=False, _
        Filter:=strFilter, _
        Flags:=lngFlags, _
        DialogTitle:="Save As", _
        fileName:=strFileName)

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnSaveFile)"
    End Select
    Resume Exit_Procedure
    
End Function

' =================================
' FUNCTION:     fxnDeleteFile
' Description:  Deletes the specified file; this is preferred over the Kill command
'               because it works for hidden files and read-only files
' Parameters:   strPath - the path and file name to be deleted
' Returns:      True if deleted, or False if error
' Throws:       none
' References:   fxnFileExists
' Source/date:  John R. Boetsch, May 19, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnDeleteFile(ByVal strPath As String) As Boolean
    On Error GoTo Err_Handler

    fxnDeleteFile = False    ' Default in case of error

    Dim FS As Variant

    Set FS = CreateObject("Scripting.FileSystemObject")
    If fxnFileExists(strPath) Then
        FS.DeleteFile strPath, True
        fxnDeleteFile = True
    Else
        MsgBox "Unable to delete the specified file", vbCritical, _
            "File delete error (fxnDeleteFile)"
    End If

Exit_Procedure:
    On Error Resume Next
    Set FS = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnDeleteFile)"
            Resume Exit_Procedure
    End Select

End Function