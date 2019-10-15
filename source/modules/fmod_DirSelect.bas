' =================================
' MODULE:       fmod_DirSelect
' Level:        Framework module
' Version:      1.00
'
' Description:  directory selection related functions & procedures
'
' Source/date:  Bonnie Campbell, October 2, 2019
' Adapted:      -
' Revisions:    BLC - 10/2/2019 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

' ----------------
'  Events
' ----------------

' ---------------------------------
' Sub:          SelectFolder
' Description:  directory selection actions
' Assumptions:  Microsoft Office 14.0 (or current version) Object Library is required
'                   Microsoft Office 14.0 Object Library in Access 2010
'                   Microsoft Office 15.0 Object Library in Access 2013
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   sxschech, August 14, 2016
'   https://www.tek-tips.com/viewthread.cfm?qid=1768657
' Source/date:  Bonnie Campbell, October 2019
' Adapted:
'   http://answers.microsoft.com/en-us/office/forum/office_2003-customize/vba-example-select-a-directory/f1c57e80-8185-48de-8c03-8bc52770a44e
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Public Function SelectFolder()
On Error GoTo Err_Handler
    Dim fd As FileDialog
    Dim FolderName As String
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.AllowMultiSelect = False
    fd.Title = "Choose the directory you would like to save the file in"
    If fd.Show = True Then
        FolderName = fd.SelectedItems(1)
    End If
        
    'Return Folder name and path
    SelectFolder = FolderName

Exit_Handler:
    'clear file dialog
    Set fd = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SelectFolder[fmod_DirSelect])"
    End Select
    Resume Exit_Handler
End Function