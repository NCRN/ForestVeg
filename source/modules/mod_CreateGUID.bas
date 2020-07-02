' =================================
' MODULE:       basUtilities
' Description:  Standard module for creating a GUID value on demand
' Source/date:  Ben Baird, http://vbthunder.com/, downloaded 12/22/2005
' Revisions:    John R. Boetsch, May 26, 2006 - documentation and minimal edits
'               BLC - 3/10/2020 - 64-bit OS updates

Option Compare Database
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As LongPtr
    
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" _
        (rguid As Any, ByVal lpstrClsId As LongPtr, ByVal cbMax As LongPtr) As LongPtr
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As guid) As Long
    
    Private Declare Function StringFromGUID2 Lib "ole32.dll" _
        (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
#End If

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
'               BLC - 3/10/2020 - 64-bit OS updates
' =================================

Public Function fxnGUIDGen() As String
    On Error GoTo Err_Handler

    Dim uGUID As GUID       ' the structured guid
    Dim sGUID As String     ' for storing the results
    Dim bGUID() As Byte     ' the formatted string
    Dim lLen As LongPtr
    Dim RetVal As LongPtr
    lLen = 40
    bGUID = String(CInt(lLen), 0)

    ' use the API to generate the guid
    CoCreateGuid uGUID

    ' use the API to format as string
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID
    If (Asc(mid$(sGUID, CLng(RetVal), 1)) = 0) Then RetVal = RetVal - 1
    ' truncate the string
    fxnGUIDGen = Left$(sGUID, CLng(RetVal))

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