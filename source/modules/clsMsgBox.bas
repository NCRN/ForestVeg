Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       clsMsgBox
' Level:        Framework class
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
Public Enum MessageBoxIcon
    NoIcon = 0
    Critical = &H10
    Question = &H20
    Exclamation = &H30
    Information = &H40
    DefaultButton1 = 0
    DefaultButton2 = &H100
    DefaultButton3 = &H200
End Enum

Public Enum MessageBoxReturn
    Unknown
    Button1
    Button2
    Button3
End Enum

Private Const m_sSource As String = "clsMsgbox"
Private Const GWL_HINSTANCE As Long = (-6)
Private Const WH_CBT = 5
Private Const MB_TASKMODAL = &H2000&

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As LongPtr
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As LongPtr, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As LongPtr) As LongPtr
    Private Declare PtrSafe Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
    Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
#End If

Private m_sButtonText1 As String
Private m_sButtonText2 As String
Private m_sButtonText3 As String
Private m_sPrompt As String
Private m_sTitle As String
Private m_eIcon As MessageBoxIcon
Private m_bUseCancel As Boolean

#If Win64 Then
    Private m_hInstance As LongPtr
    Private m_hThreadID As LongPtr
    Private m_lProcHook As LongPtr
#Else
    Private m_hInstance As Long
    Private m_hThreadID As Long
    Private m_lProcHook As Long
#End If

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------
#If Win64 Then
    Public Property Get ProcHook() As LongPtr
        ProcHook = m_lProcHook
    End Property
#Else
    Public Property Get ProcHook() As Long
        ProcHook = m_lProcHook
    End Property
#End If
Public Property Get UseCancel() As Boolean
    UseCancel = m_bUseCancel
End Property
Public Property Let UseCancel(ByVal NewValue As Boolean)
    m_bUseCancel = NewValue
End Property
Public Property Let Prompt(ByVal NewValue As String)
    m_sPrompt = NewValue
End Property
Public Property Get Prompt() As String
    Prompt = m_sPrompt
End Property
Public Property Let title(ByVal NewValue As String)
    m_sTitle = NewValue
End Property
Public Property Get title() As String
    title = m_sTitle
End Property
Public Property Let Icon(ByVal NewValue As MessageBoxIcon)
    m_eIcon = NewValue
End Property
Public Property Get Icon() As MessageBoxIcon
    Icon = m_eIcon
End Property
Public Property Let ButtonText1(ByVal NewValue As String)
    m_sButtonText1 = NewValue
End Property
Public Property Get ButtonText1() As String
    ButtonText1 = m_sButtonText1
End Property
Public Property Let ButtonText2(ByVal NewValue As String)
    m_sButtonText2 = NewValue
End Property
Public Property Get ButtonText2() As String
    ButtonText2 = m_sButtonText2
End Property
Public Property Let ButtonText3(ByVal NewValue As String)
    m_sButtonText3 = NewValue
End Property
Public Property Get ButtonText3() As String
    ButtonText3 = m_sButtonText3
End Property
#If Win64 Then
    Private Property Get PtrFromObject(ByRef obj As Object) As LongPtr
        PtrFromObject = ObjPtr(obj)
    End Property
#Else
    Private Property Get PtrFromObject(ByRef obj As Object) As Long
        PtrFromObject = ObjPtr(obj)
    End Property
#End If

' ----------------
'  Instantiation
' ----------------
' ---------------------------------
' Sub:          Class_Initialize
' Description:  initialize custom message box class
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
Private Sub Class_Initialize()
On Error GoTo Err_Handler

    #If Win64 Then
        m_hInstance = GetWindowLong(CLngPtr(hWndApplication), GWL_HINSTANCE)
    #Else
        m_hInstance = GetWindowLong(hWndApplication, GWL_HINSTANCE)
    #End If
    
    m_hThreadID = GetCurrentThreadId()
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[clsMsgBox])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  terminate custom message box class
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
Private Sub Class_Terminate()
On Error GoTo Err_Handler

   RemovePropPointer
   
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[clsMsgBox])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Functions
' ----------------
' ---------------------------------
' Function:     MessageBox
' Description:  generates message box
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
Public Function MessageBox() As MessageBoxReturn
On Error GoTo Err_Handler

    Dim bCancel As Boolean
    
    #If Win64 Then
        Dim lR As LongPtr
        Dim lType As LongPtr
    #Else
        Dim lR As Long
        Dim lType As Long
    #End If

    
    If Not m_hInstance > 0 Then Err.Raise vbObjectError + 1, m_sSource, "Instance handle not found"
    If Not m_hThreadID > 0 Then Err.Raise vbObjectError + 2, m_sSource, "Thread id not found"

    If Len(Me.title) = 0 Then Me.title = "Microsoft Excel"
    bCancel = Me.UseCancel
    
    If Len(Me.ButtonText1) > 0 And Len(Me.ButtonText2) > 0 And Len(Me.ButtonText3) > 0 Then
        lType = Me.Icon Or IIf(bCancel, vbYesNoCancel, vbAbortRetryIgnore)
        
    ElseIf Len(Me.ButtonText1) > 0 And Len(Me.ButtonText2) Then
        lType = Me.Icon Or IIf(bCancel, vbOKCancel, vbYesNo)
    Else
        If Len(Me.ButtonText1) = 0 Then Me.ButtonText1 = "OK"
        lType = Me.Icon Or vbOKOnly
    End If
    
    m_lProcHook = SetWindowsHookEx(WH_CBT, _
                                   AddressOf MsgBoxHookProc, _
                                   m_hInstance, _
                                   m_hThreadID)
                                      
                                      
    'Private Declare PtrSafe Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As LongPtr) As LongPtr

    SetProp hWndApplication, "ObjPtr", PtrFromObject(Me)
    lR = MessageBoxA(hWndApplication, Me.Prompt, Me.title, lType Or MB_TASKMODAL)
    
    If Len(Me.ButtonText1) > 0 And Len(Me.ButtonText2) > 0 And Len(Me.ButtonText3) > 0 Then
        If lR = IIf(bCancel, vbYes, vbAbort) Then
            MessageBox = Button1
        ElseIf lR = IIf(bCancel, vbNo, vbRetry) Then
            MessageBox = Button2
        ElseIf lR = IIf(bCancel, vbCancel, vbIgnore) Then
            MessageBox = Button3
        End If
    ElseIf Len(Me.ButtonText1) > 0 And Len(Me.ButtonText2) Then
        If lR = IIf(bCancel, vbOK, vbYes) Then
            MessageBox = Button1
        ElseIf lR = IIf(bCancel, vbCancel, vbNo) Then
            MessageBox = Button2
        End If
    Else
        If lR = vbOK Then
            MessageBox = Button1
        End If
    End If

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MessageBox[clsMsgBox])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     MessageBoxEx
' Description:  terminates message box
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
Public Function MessageBoxEx(ByVal Prompt As String, _
                             Optional Icon As MessageBoxIcon, _
                             Optional ByVal title As String, _
                             Optional ByVal ButtonText1 As String, _
                             Optional ByVal ButtonText2 As String, _
                             Optional ByVal ButtonText3 As String) As MessageBoxReturn
On Error GoTo Err_Handler

    Me.Prompt = Prompt
    Me.Icon = Icon
    Me.title = title
    Me.ButtonText1 = ButtonText1
    Me.ButtonText2 = ButtonText2
    Me.ButtonText3 = ButtonText3
    
    MessageBoxEx = MessageBox

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MessageBoxEx[clsMsgBox])"
    End Select
    Resume Exit_Handler
End Function