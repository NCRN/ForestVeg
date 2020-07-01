Option Compare Database
Option Explicit

' =================================
' MODULE:       fmod_User
' Level:        Application module
' Version:      1.00
'
' Description:  Application user related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 3, 2020
' Revisions:    BLC, 4/3/2020  - 1.00 - initial version
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
' SUB:          GetUsersName
' Description:  returns current user's display name
' Assumptions:
'               user is logged into a system that will return display name using LDAP
' Parameters:   -
' Returns:      user's full name last, first MI (string) or NULL if no name is returned
' Throws:       none
' References:
'   BenV, August 13, 2019
'   https://stackoverflow.com/questions/57486049/how-to-get-email-address-with-vba-based-on-windows-login-name
' Source/date:  Bonnie Campbell, April 3, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/3/2020 - initial version
' ---------------------------------
Public Function GetUsersName(Optional strUserName As String) As Variant
On Error GoTo Err_Handler

    Dim sysInfo         As Object
    Dim oUser           As Object
    Dim DisplayName     As String

    If strUserName = "" Then
        ' No name was passed in.  Get it for the current user.
        strUserName = Environ("USERNAME")
    End If

    Set sysInfo = CreateObject("ADSystemInfo")
    Set oUser = GetObject("LDAP://" & sysInfo.UserName & "")    'requires connection to server (either direct or VPN)

'    Debug.Print "Display Name: "; Tab(20); oUser.Get("DisplayName")
'    Debug.Print "Email Address: "; Tab(20); oUser.Get("mail")
'    Debug.Print "Computer Name: "; Tab(20); sysInfo.ComputerName
'    Debug.Print "Site Name: "; Tab(20); sysInfo.SiteName
'    Debug.Print "Domain DNS Name: "; Tab(20); sysInfo.DomainDNSName

'    GetEmailAddress = oUser.Get("mail")

    DisplayName = oUser.Get("DisplayName")
    
    'parse user's name to first & last
    If Len(DisplayName) > 0 Then
        Dim LastName As String
        Dim MI As String
        Dim FirstName As String
        Dim FullName As String
        Dim aryName As Variant
        
        LastName = Left(DisplayName, InStr(DisplayName, ",") - 1)
        MI = Right(DisplayName, 1)
        FirstName = Replace(Replace(DisplayName, LastName & ", ", ""), " " & MI, "")
    
        FullName = FirstName & "," & MI & "," & LastName
        aryName = Split(FullName, ",")
    End If
        
    GetUsersName = aryName
    
Exit_Handler:
    Set sysInfo = Nothing
    Set oUser = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case -2147023541 'no direct connection to server so no LDAP info -> ask to select user
        DoCmd.OpenForm "frm_SelectUser", acNormal, , , acFormEdit, acDialog
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetUsersName[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function