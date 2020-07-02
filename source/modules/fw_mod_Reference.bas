Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_Reference
' Level:        Framework module
' Version:      1.00
' Description:  Framework-wide Reference related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, May 16, 2019
' Revisions:    BLC, 5/16/2019 - 1.00 - initial version
' =================================

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
' SUB:          VerifyReferences
' Description:  verify reference actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 16, 2019
' Adapted:  -
' Revisions:
'   BLC - 5/16/2019 - initial version
' ---------------------------------
Public Function VerifyReferences()
On Error GoTo Err_Handler
    
    'references
    Dim ref As Reference

    'iterate through references
    For Each ref In References
        'check if broken
        If ref.IsBroken = False Then
            Debug.Print "Name: ", ref.Name
            Debug.Print "FullPath: ", ref.FullPath
            Debug.Print "Version: ", ref.Major & "." & ref.Minor
        Else
            Debug.Print "GUIDs of broken references:"
            Debug.Print ref.GUID
        End If
    Next

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyReferences[fw_mod_Reference])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          AddReference
' Description:  add reference from file actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 16, 2019
' Adapted:  -
' Revisions:
'   BLC - 5/16/2019 - initial version
' ---------------------------------
Public Function AddReference(FullPath As String)
On Error GoTo Err_Handler

    Dim RefFile As String
    
    RefFile = ParseFileName(FullPath)
    
    If FileExists(FullPath) Then
    
        References.AddFromFile FullPath
    End If
    
    If References.Item(RefFile).IsBroken = False Then
        MsgBox RefFile & " reference successfully added!", vbInformation, "Add Reference Results..."
    Else
        MsgBox RefFile & " reference addition failed!", vbCritical, "Add Reference Results..."
    End If
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case 32813 'Name conflicts w/ existing module, project, or object library
        Dim resp As Boolean
        resp = MsgBox("Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
            "Click 'YES' to replace the existing version " & vbCrLf & vbCrLf & _
            "or 'NO' to leave the existing version as is.", vbYesNo, _
            "Error encountered (#" & Err.Number & " - AddReference[fw_mod_Reference])")
        
        If resp = True Then ReplaceReference (RefFile)
        
        Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddReference[fw_mod_Reference])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          RemoveReference
' Description:  remove reference from file actions
' Assumptions:  -
' Parameters:   refFile - full path name of reference (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 16, 2019
' Adapted:  -
' Revisions:
'   BLC - 5/16/2019 - initial version
' ---------------------------------
Public Function RemoveReference(RefFile As String)
On Error GoTo Err_Handler
   
    Dim ref As Reference, oRef As Reference
    Dim refName As String
   
    refName = GetFileNameOnly(RefFile)
    Set ref = GetReferenceFromPath(RefFile)
    
    For Each oRef In References
        If oRef.Name = refName Then
            If References.Item(refName).IsBroken = False Then
        
            References.Remove ref
                
            End If
            
            If References.Item(ref).IsBroken = True Then
                MsgBox RefFile & " reference successfully removed!", vbInformation, "Remove Reference Results..."
            Else
                MsgBox RefFile & " reference removal failed!", vbCritical, "Remove Reference Results..."
            End If
        End If
    Next
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case 32813 'Name conflicts w/ existing module, project, or object library
        Dim resp As Boolean
        resp = MsgBox("Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
            "Click 'YES' to replace the existing version " & vbCrLf & vbCrLf & _
            "or 'NO' to leave the existing version as is.", vbYesNo, _
            "Error encountered (#" & Err.Number & " - RemoveReference[fw_mod_Reference])")
        
        If resp = True Then ReplaceReference (RefFile)
        
        Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveReference[fw_mod_Reference])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ReplaceReference
' Description:  replace reference from file actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 16, 2019
' Adapted:  -
' Revisions:
'   BLC - 5/16/2019 - initial version
' ---------------------------------
Public Function ReplaceReference(FullPath As String)
On Error GoTo Err_Handler

    Dim RefFile As String
    
    RefFile = ParseFileName(FullPath)
    
    If FileExists(FullPath) Then
    
        References.AddFromFile FullPath
    End If
    
    If References.Item(RefFile).IsBroken = False Then
        MsgBox RefFile & " reference successfully added!", vbInformation, "Add Reference Results..."
    Else
        MsgBox RefFile & " reference addition failed!", vbCritical, "Add Reference Results..."
    End If
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case 32813 'Name conflicts w/ existing module, project, or object library
        Dim resp As Boolean
        resp = MsgBox("Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
            "Click 'YES' to replace the existing version " & vbCrLf & vbCrLf & _
            "or 'NO' to leave the existing version as is.", vbYesNo, _
            "Error encountered (#" & Err.Number & " - AddReference[fw_mod_Reference])")
        
        If resp = True Then ReplaceReference (RefFile)
        
        Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReplaceReference[fw_mod_Reference])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          GetReferenceFromPath
' Description:  retrieve a reference object from a file full path
' Assumptions:  -
' Parameters:   refFile - full path name to reference file (string)
' Returns:      -
' Throws:       none
' References:
'   PaulG, February 2, 2017
'   https://stackoverflow.com/questions/41886667/remove-a-vba-project-reference
' Source/date:  Bonnie Campbell, May 16, 2019
' Adapted:  -
' Revisions:
'   BLC - 5/16/2019 - initial version
' ---------------------------------
Public Function GetReferenceFromPath(ByVal RefFile As String) As Object
On Error GoTo Err_Handler

    Dim oFile As Object, oReferences As Object, oReference As Object
    Dim FileName As String, RefFileName As String
    
    Set oFile = Interaction.CreateObject("Scripting.FileSystemObject")
    
    Set oReferences = References
    
    FileName = ParseFileName(RefFile)
    
    For Each oReference In oReferences
        RefFileName = oFile.GetFileName(oReference.FullPath)
        If StrComp(FileName, RefFileName, vbTextCompare) = 0 Then
            Set GetReferenceFromPath = oReference
        End If
    Next

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetReferenceFromPath[fw_mod_Reference])"
    End Select
    Resume Exit_Handler
End Function