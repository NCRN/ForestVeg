Option Compare Database
Option Explicit

' =================================
' MODULE:       fmod_MsgBox
' Level:        Framework module
' Version:      1.00
'
' Description:  custom message box related functions & procedures
' References:
'   Dean Kinnear, June 8, 2017
'   http://shutupdean.com/blog/2014/08/01/vba-msgbox-custom-button-text/
' Source/date:  Bonnie Campbell, October 2, 2019
' Adapted:      -
' Revisions:    BLC - 10/2/2019 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private Const HCBT_ACTIVATE = 5

Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7

#If Win64 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
    Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As LongPtr, ByVal lpString As String) As LongPtr
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
#End If

Private m_hWnd As Long

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------
Public Property Get hWndApplication() As Long
    If m_hWnd = 0 Then
        If Application.Name = "Microsoft Access" Then
             m_hWnd = FindWindow("OMain", vbNullString)
        ElseIf Application.Name = "Microsoft Word" Then
            m_hWnd = FindWindow("OpusApp", vbNullString)
        ElseIf Application.Name = "Microsoft Excel" Then
            m_hWnd = FindWindow("XLMAIN", vbNullString)
        End If
    End If
    hWndApplication = m_hWnd
End Property

#If Win64 Then
    Private Property Get ObjectFromPtr(ByVal lPtr As LongPtr) As Object
        Dim obj As Object
        
        CopyMemory obj, lPtr, 4
        Set ObjectFromPtr = obj
        CopyMemory obj, 0&, 4
    End Property
#Else
    Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
        Dim obj As Object
        
        CopyMemory obj, lPtr, 4
        Set ObjectFromPtr = obj
        CopyMemory obj, 0&, 4
    End Property
#End If

' ----------------
'  Methods
' ----------------
' ---------------------------------
' Sub:          MsgBoxHookProc
' Description:  message hook procedures
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Dean Kinnear, June 8, 2017
'   http://shutupdean.com/blog/2014/08/01/vba-msgbox-custom-button-text/
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Public Function MsgBoxHookProc(ByVal uMsg As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long
On Error GoTo Err_Handler

    #If Win64 Then
        Dim lPtr As LongPtr
        Dim lProcHook As LongPtr
    #Else
        Dim lPtr As Long
        Dim lProcHook As Long
    #End If
    
    Dim cM As clsMsgBox
    
    Select Case uMsg
        Case HCBT_ACTIVATE
            lPtr = GetProp(hWndApplication, "ObjPtr")
            If (lPtr <> 0) Then
                Set cM = ObjectFromPtr(lPtr)
                If Not cM Is Nothing Then
                    If Len(cM.ButtonText1) > 0 And Len(cM.ButtonText2) > 0 And Len(cM.ButtonText3) > 0 Then
                        If cM.UseCancel Then
                            SetDlgItemText wParam, IDYES, cM.ButtonText1
                            SetDlgItemText wParam, IDNO, cM.ButtonText2
                            SetDlgItemText wParam, IDCANCEL, cM.ButtonText3
                        Else
                            SetDlgItemText wParam, IDABORT, cM.ButtonText1
                            SetDlgItemText wParam, IDRETRY, cM.ButtonText2
                            SetDlgItemText wParam, IDIGNORE, cM.ButtonText3
                        End If
                        
                    ElseIf Len(cM.ButtonText1) > 0 And Len(cM.ButtonText2) Then
                        If cM.UseCancel Then
                            SetDlgItemText wParam, IDOK, cM.ButtonText1
                            SetDlgItemText wParam, IDCANCEL, cM.ButtonText2
                        Else
                            SetDlgItemText wParam, IDYES, cM.ButtonText1
                            SetDlgItemText wParam, IDNO, cM.ButtonText2
                        End If
                    Else
                        SetDlgItemText wParam, IDOK, cM.ButtonText1
                    End If
                    lProcHook = cM.ProcHook
                End If
            End If
            RemovePropPointer
            If lProcHook <> 0 Then UnhookWindowsHookEx lProcHook
    End Select
    
    MsgBoxHookProc = False

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MsgBoxHookProc[fmod_MsgBox])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          RemovePropPointer
' Description:  remove property pointer
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Dean Kinnear, June 8, 2017
'   http://shutupdean.com/blog/2014/08/01/vba-msgbox-custom-button-text/
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Public Sub RemovePropPointer()
    #If Win64 Then
        Dim lPtr As LongPtr
    #Else
        Dim lPtr As Long
    #End If
    
    lPtr = GetProp(hWndApplication, "ObjPtr")
    If lPtr <> 0 Then RemoveProp hWndApplication, "ObjPtr"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MsgBoxHookProc[fmod_MsgBox])"
    End Select
    Resume Exit_Handler
End Sub