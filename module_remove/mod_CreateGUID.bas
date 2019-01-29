Attribute VB_Name = "mod_CreateGUID"
' =================================
' MODULE:       basUtilities
' Description:  Standard module for creating a GUID value on demand
' Source/date:  Ben Baird, http://vbthunder.com/, downloaded 12/22/2005
' Revisions:    John R. Boetsch, May 26, 2006 - documentation and minimal edits

Option Compare Database
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long

Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" _
    (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

' =================================
' FUNCTION:     fxnGUIDGen
' Description:  Generates 16-byte globally-unique identifiers (GUIDs)
' Parameters:   none
' Returns:      a formatted guid as a string type, which can be saved directly
'               to either a string or a guid field
' Throws:       none
' References:   CoCreateGuid API to generate guid, StringFromGUID2 API to
'               format as a string
' Source/date:  Ben Baird, http://vbthunder.com/, downloaded 12/22/2005
' Revisions:    John R. Boetsch, May 26, 2006 - documentation and minimal edits
' =================================

Public Function fxnGUIDGen() As String
    On Error GoTo Err_Handler

    Dim uGUID As GUID       ' the structured guid
    Dim sGUID As String     ' for storing the results
    Dim bGUID() As Byte     ' the formatted string
    Dim lLen As Long
    Dim RetVal As Long
    lLen = 40
    bGUID = String(lLen, 0)

    ' use the API to generate the guid
    CoCreateGuid uGUID

    ' use the API to format as string
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID
    If (Asc(mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
    ' truncate the string
    fxnGUIDGen = Left$(sGUID, RetVal)

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnGUIDGen)"
            Resume Exit_Procedure
    End Select

End Function
