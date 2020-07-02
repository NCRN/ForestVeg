Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_CSV
' Level:        Framework module
' Version:      1.02
' Description:  Framework-wide CSV related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, September 30, 2016 for NCPN tools
' Revisions:    BLC, 9/30/2016 - 1.00 - initial version
'               BLC, 10/6/2017 - 1.01 - added UploadCSVFile() from mod_App_Data
'               BLC, 5/16/2019 - 1.02 - added fw_ module prefix
' =================================

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, September 30, 2016 for NCPN tools
' Adapted:      -
' Revisions:    BLC, 9/30/2016 - initial version
' ---------------------------------

'-----------------------------------------------------------------------
' Constants
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' Declarations
'-----------------------------------------------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

'-----------------------------------------------------------------------
' Functions
'-----------------------------------------------------------------------

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' SUB:          ImportCSV
' Description:  CSV import actions
' Assumptions:  If DeleteExistingTable is False, data will append to table if possible
' Parameters:   strPath - CSV full file path (string)
'               strTable - table to insert data into (string)
'               HasHeaders - whether CSV first row is a header row
'                            (boolean, optional, default = true)
'               DeleteExistingTable - whether table should be deleted first
'                                     (boolean, optional, default = true)
' Returns:      -
' Throws:       none
' References:   -
'
'
' Source/date:  Bonnie Campbell, September 30, 2016 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 9/30/2015 - initial version
' ---------------------------------
Public Sub ImportCSV(strPath As String, strTable As String, _
                    Optional hasHeaders As Boolean = True, _
                    Optional DeleteExistingTable As Boolean = True)
On Error GoTo Err_Handler

    'remove existing table --> otherwise append
    If DeleteExistingTable Then
        If TableExists(strTable) Then _
            DoCmd.DeleteObject acTable, strTable
    End If

    DoCmd.TransferText acImportDelim, , strTable, strPath, hasHeaders

Exit_Handler:
    'cleanup
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ImportCSV[fw_mod_CSV])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          UploadCSVFile
' Description:  Uploads data into database from CSV file
' Assumptions:  -
' Parameters:   strFilename - name of file being uploaded (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/1/2016 - initial version
'   BLC - 10/19/2016 - renamed to UploadCSVFile from UploadSurveyFile to genericize
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'                   - un-comment out
' --------------------------------------------------------------------
' ---------------------------------
Public Sub UploadCSVFile(strFileName As String)
On Error GoTo Err_Handler

    'import to table
    ImportCSV strFileName, "usys_temp_csv", True, True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UploadCSVFile[fw_mod_CSV])"
    End Select
    Resume Exit_Handler
End Sub