Option Compare Database
Option Explicit

' =================================
' MODULE:       fmod_GUID
' Level:        Application module
' Version:      1.00
'
' Description:  Application GUID related functions & subroutines
'
' Source/date:  Bonnie Campbell, August 28, 2019
' Revisions:    BLC, 8/28/2019  - 1.00 - initial version
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
' -- Constants --

' -- Values --

' -- Functions --


' ---------------------------------
'  Methods
' ---------------------------------

' *********************************
'    Common
' *********************************

' ---------------------------------
' SUB:          FormatIsGUID
' Description:  checks a GUID string to determine if it is in GUID format
' Assumptions:
'   matches GUIDs with the following format:    #(14)-(1)#(9).(1)#(6)
'   for example:                                {20190503115602-533424019.813538}
' Parameters:   -
' Returns:      True or False depending on whether strGUID has proper GUID format
' Throws:       none
' References:
'   BenV, September 22, 2010
'   https://stackoverflow.com/questions/3770672/regular-expressions-in-ms-access-vba
'   Vernon W. Hui, Microsoft, May 10, 1999
'   https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ms974570(v=msdn.10)#scripting05_topic2
' Source/date:  Bonnie Campbell, August 28, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/28/2019 - initial version
' ---------------------------------
Public Function FormatIsGUID(strGUID As Variant)
On Error GoTo Err_Handler
        
    Dim RegEx As New RegExp
    
'    regex.Pattern = "^(\{){0,1}[0-9a-fA-F]{8}\-" & _
'                         "[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-" & _
'                         "[0-9a-fA-F]{12}(\}){0,1}$"

    RegEx.Pattern = "^\{\d{14}-\d{9}.\d{6}\}$"
    
    FormatIsGUID = RegEx.Test(strGUID)
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormatIsGUID[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          String2GUID
' Description:  creates a GUID string from a string value
' Assumptions:
'               Access database GUIDs created as Number > Replication ID values
'               are *string* values with brackets & hyphens
'               This functions returns those features to the GUID string
'               Usage:
'                   rst_BillTo.Fields("LineItemGUID") = strGUID
'               Includes check for double {{ or }} to avoid adding extras
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   kmoloney, January 28, 2004
'   https://www.experts-exchange.com/questions/20866254/Can't-Generate-GUID-from-String-ActiveX-component-for-'GUIDFromString-'.html
' Source/date:  Bonnie Campbell, August 28, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/28/2019 - initial version
' ---------------------------------
Public Function String2GUID(strGUID As Variant)
On Error GoTo Err_Handler
        
'    strGUID = "{" & _
'    Left(strGUID, 8) & "-" & _
'    Mid(strGUID, 9, 4) & "-" & _
'    Mid(strGUID, 13, 4) & "-" & _
'    Mid(strGUID, 17, 4) & "-" & _
'    Right(strGUID, 12) & "}"
        
    strGUID = "{" & strGUID & "}"
    
    'ensure there are no double {{ or }}
    Do Until (Len(Replace(strGUID, "{{", "")) = Len(strGUID) And _
              Len(Replace(strGUID, "}}", "")) = Len(strGUID))
        
        strGUID = Replace(Replace(strGUID, "{{", "{"), "}}", "}")
    
    Loop
    
    String2GUID = strGUID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - String2GUID[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function