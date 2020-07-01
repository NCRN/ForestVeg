Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5040
    DatasheetFontHeight =9
    ItemSuffix =35
    Left =7605
    Top =2145
    Right =12645
    Bottom =6030
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xa818e7379372e540
    End
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ToggleButton
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =0
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedColor =0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeColor =0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =0
            PressedForeThemeColorIndex =1
        End
        Begin FormHeader
            Height =3000
            BackColor =14277338
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =60
                    Top =600
                    Width =4920
                    Height =1080
                    BackColor =13434879
                    BorderColor =10921638
                    Name ="rctEdit"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =600
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =1680
                    BackThemeColorIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =5040
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =275078
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Select Current User"
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =540
                    BackThemeColorIndex =5
                    BackShade =50.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1800
                    Width =4800
                    Height =480
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblDirections"
                    Caption ="Select the current user (i.e. you!). "
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1800
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =2280
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4260
                    Top =180
                    Width =720
                    Height =285
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxDevMode"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4260
                    LayoutCachedTop =180
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =465
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1260
                    Top =960
                    Width =3600
                    Height =315
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =2
                    ForeColor =8355711
                    Name ="tbxID"
                    ControlTipText ="Selected user's ID"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =960
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1275
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =120
                            Top =960
                            Width =1020
                            Height =315
                            FontWeight =500
                            Name ="lblID"
                            Caption ="ID"
                            LayoutCachedLeft =120
                            LayoutCachedTop =960
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =1275
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1260
                    Top =645
                    Width =3600
                    Height =315
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =3
                    ForeColor =8355711
                    Name ="tbxSelectedUser"
                    ControlTipText ="Selected user's name"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =645
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =645
                    Width =1020
                    Height =315
                    FontWeight =500
                    Name ="lblSelectedUser"
                    Caption ="USER"
                    LayoutCachedLeft =120
                    LayoutCachedTop =645
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =960
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =720
                    Height =285
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =4
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxUsername"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =840
                    LayoutCachedHeight =405
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeThemeColorIndex =1
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =12
                    Left =1020
                    Top =2400
                    Width =3600
                    Height =435
                    ColumnOrder =1
                    FontWeight =700
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cbxUsers"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) & (\" \"+[Mi"
                        "ddle_Init]) AS FullName, tlu_Contacts.Organization, tlu_Contacts.Position_title "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Last_Name, tlu_Contacts.First_Name;"
                    ColumnWidths ="0;2160;720;720"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""
                    OnChange ="[Event Procedure]"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =2400
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =2835
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =60
                            Top =2400
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblUsers"
                            Caption ="Users"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2400
                            LayoutCachedWidth =930
                            LayoutCachedHeight =2915
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
        End
        Begin FormFooter
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =900
            BackColor =15921906
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =204
                    Left =705
                    Top =60
                    Width =2325
                    Height =720
                    FontSize =14
                    ForeColor =0
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save new date"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =705
                    LayoutCachedTop =60
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =780
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =204
                    Left =3105
                    Top =60
                    Width =1020
                    Height =720
                    FontSize =14
                    TabIndex =1
                    ForeColor =0
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close this form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3105
                    LayoutCachedTop =60
                    LayoutCachedWidth =4125
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =255
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       frm_Select_User
' Level:        Application module
' Version:      1.01
'
' Description:  add event related functions & procedures
'
' Source/date:  Bonnie Campbell, February, 2019
' Adapted:      -
' Revisions:    BLC - 4/2/2020 - 1.00 - initial version
'               BLC - 4/17/2019 - 1.01 - update open pseudo event
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

Private m_OrigDate As Date
Private m_EditByID As Long
Private m_EditTable As String
Private m_EditField As String
Private m_EditIDField As String
Private m_EditID As String
Private m_Edit As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(value As String)
Public Event InvalidDirections(value As String)
Public Event InvalidCallingForm(value As String)

'---------------------
' Properties
'---------------------
Public Property Let title(value As String)
    If Len(value) > 0 Then
        m_Title = value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        'Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(value)
    End If
End Property

Public Property Get title() As String
    title = m_Title
End Property

Public Property Let Directions(value As String)
    If Len(value) > 0 Then
        m_Directions = value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(value As String)
        m_CallingForm = value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'editing info

' ----------------
'  Events
' ----------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "frm_Switchboard"
'
    If Len(Me.OpenArgs) > 0 Then Me.CallingForm = Me.OpenArgs
'
'    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'dev mode
    tbxDevMode = DEV_MODE
                
    title = "Select Current User"
    'lblTitle.Caption = "" 'clear header title
    Directions = "Sorry, I can't tell who you are. Please let me know and select the current user (i.e. you!). "
    
    'defaults
    lblDirections.ForeColor = lngBlue
    btnSave.HoverColor = lngGreen
    btnCancel.HoverColor = lngRed
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Select_User form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
  
'    Me.tbxSelectedUser = cbxUsers.SelText
'    Me.tbxID = cbxUsers
'    Me.tbxUsername = Environ("USERNAME")

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Current
' Description:  current form actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxUsers_Change
' Description:  combobox change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 30, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/30/2020 - initial version
' ---------------------------------
Private Sub cbxUsers_Change()
On Error GoTo Err_Handler
    
'    Me.tbxSelectedUser = cbxUsers.Value '.Column(1)
'    Me.tbxID = cbxUsers.Column(0)
'
'    Debug.Print "ID " & cbxUsers.Column(0)
'
'    Me.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUsers_Change[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxUsers_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub cbxUsers_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim ary() As String
    ary = Split(cbxUsers.Column(1), ", ")
    
    Me.tbxSelectedUser = cbxUsers.Column(1)
    Me.tbxID = cbxUsers.Column(0)

    SetTempVar "UserID", cbxUsers.Column(0)

    Me.tbxUsername = TempVars("UserID")

    Debug.Print ary(1) & " " & ary(0)

    Debug.Print "ID " & cbxUsers.Column(0) & " " & cbxUsers.Column(1)
    
    Me.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUsers_AfterUpdate[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnSave_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler

    'open form not in new event record mode (acFormAdd) so it accepts form defaults
'    DoCmd.OpenForm "EventAdd", acNormal, , , , acWindowNormal, Me.cbxParkCode
'    DoCmd.OpenForm "EventAdd2", acNormal, , , , acWindowNormal, Me.cbxParkCode

    'return to the prior form
    'maximize calling form
    ToggleForm Me.CallingForm, 0
    
    DoCmd.Close acForm, Me.Name
    

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCancel_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
' ---------------------------------
Private Sub btnCancel_Click()
On Error GoTo Err_Handler
   
    DoCmd.Close
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[frm_Select_User])"
    End Select
    Resume Exit_Handler
End Sub

Public Sub GetFullNameX()
    Dim oAD As Object
    Dim oUser As Object
    Dim strDisplayName As String
    Dim givenName As String
    Dim surname As String
    Dim ComputerName As String
    
    'Getting computer name
ComputerName = Environ("computername")
    
    Set oAD = CreateObject("ADSystemInfo")
    Set oUser = GetObject("LDAP://" & oAD.UserName)
    
    strDisplayName = oUser.DisplayName
    givenName = oUser.givenName
    surname = oUser.sn
    
    Debug.Print strDisplayName
    Debug.Print ComputerName
    Debug.Print givenName
    Debug.Print surname
    
End Sub

Public Function TestItX()
    Dim i As Integer
    Dim stEnviron As String
    
    For i = 1 To 50
    ' get the environment variable
        stEnviron = Environ(i)
    ' see if there is a variable set
        If Len(stEnviron) > 0 Then
            Debug.Print i, Environ(i)
        Else
            Exit For
        End If
    Next
    
End Function
