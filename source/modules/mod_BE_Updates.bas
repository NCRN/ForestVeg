Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_BE_Updates
' Level:        Application module
' Version:      1.00
'
' Description:  application backend update related functions & procedures
'
' Source/date:  Bonnie Campbell, April 4, 2018
' Revisions:    BLC - 4/4/2018 - 1.00 - initial version
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Methods
' ----------------

' ---------------------------------
' SUB:          AlterBE
' Description:  alters backend database by running SQL
' Assumptions:
'  ConnectionStrings:
'    old    "Provider=Microsoft.Jet.OLEDB.4.0; Persist Security Info=False;Data Source=" & DbName
'    2007   "Provider=Microsoft.Jet.OLEDB.12.0; Persist Security Info=False;Data Source=" & DbName
'    2010   "Provider=Microsoft.Jet.OLEDB.14.0; Persist Security Info=False;Data Source=" & DbName
'    2016   "Provider=Microsoft.ACE.OLEDB.16.0; Persist Security Info=False;Data Source=" & DbName
' Notes:
'           -------------------------------------
'           Types
'           -------------------------------------
'           bit --> comes out as Yes/No
'           -------------------------------------
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   HansUp, May 20, 2011
'   David W Fenton, May 22, 2011
'   https://stackoverflow.com/questions/47535/sql-to-add-column-with-default-value-access-2003
' Source/date:  Bonnie Campbell, April 4, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version (currently unused)
' ---------------------------------
Public Sub AlterBE()
On Error GoTo Err_Handler
    
    'defaults
    Dim DbName As String
    Dim strSQL As String
    
    DbName = CurrentDb.Properties("Name") '""
    
    'use ADO (SQL 92 mode)
    'open connection
    Dim conn As ADODB.Connection
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.16.0;" & _
        "Data Source=" & DbName
    
    conn.Open
        
    'update tables
    
    '-----------------------------------------------------------------
    '2018 Pre-Season Updates:
    '-----------------------------------------------------------------
    'add DBH double check -> tbl_Sapling_Data, tbl_Tree_Data
    'table:         tbl_Sapling_Data
    'column:        DBH_Check
    'type:          bit --> comes out as
    'default:       0
    'description:   'DBH double check flag (0-not double checked, 1-double checked)'
    strSQL = "ALTER TABLE tbl_Sapling_Data_DUPE " _
           & "ADD COLUMN DBH_Check2 BYTE DEFAULT 0 " _
           & " 'DBH double check flag (0-not double checked, 1-double checked)'"
    
    conn.Execute strSQL
    
Exit_Handler:
    conn.Close
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AlterBE[mod_BE_Updates])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     RunChanges
' Description:  runs AlterBE to make BE changes
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, April 4, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version (currently unused)
' ---------------------------------
Public Function RunChanges()
On Error GoTo Err_Handler
    
   AlterBE
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunChanges[mod_BE_Updates])"
    End Select
    Resume Exit_Handler
End Function