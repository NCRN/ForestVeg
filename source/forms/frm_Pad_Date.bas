Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DataEntry = NotDefault
    AllowUpdating =2
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
    OrderBy ="Plot_Name"
    RecSrcDt = Begin
        0xe29f1894d742e540
    End
    RecordSource ="SELECT tbl_Events.*, tbl_Locations.Plot_Name, tbl_Locations.Unit_Code, tbl_Event"
        "s.PseudoEvent, tbl_Events.Event_ID FROM tbl_Locations INNER JOIN tbl_Events ON t"
        "bl_Locations.Location_ID = tbl_Events.Location_ID WHERE (((tbl_Events.PseudoEven"
        "t)=1));"
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
    FilterOnLoad =255
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
                    Caption ="Edit Date"
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
                    Caption ="Choose the desired date. Click Save to save the new date. Changing the date WILL"
                        " change underlying data."
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
                    ColumnOrder =6
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
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1020
                    Top =2400
                    Width =3480
                    Height =510
                    ColumnOrder =5
                    FontSize =16
                    FontWeight =700
                    Name ="tbxDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =2400
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =2910
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
                            Name ="lblDate"
                            Caption ="Date"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2400
                            LayoutCachedWidth =930
                            LayoutCachedHeight =2915
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1260
                    Top =960
                    Width =3600
                    Height =315
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =2
                    ForeColor =8355711
                    Name ="tbxID"
                    ControlTipText ="Record ID for the date being edited"

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
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1260
                    Top =645
                    Width =3600
                    Height =315
                    ColumnOrder =4
                    FontSize =9
                    TabIndex =3
                    ForeColor =8355711
                    Name ="tbxTable"
                    ControlTipText ="Record ID for the date being edited"

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
                    Name ="lblEditing"
                    Caption ="Event_ID"
                    LayoutCachedLeft =120
                    LayoutCachedTop =645
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =960
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =1275
                    Width =2160
                    Height =315
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =4
                    ForeColor =8355711
                    Name ="tbxOrigDateTIme"
                    ControlTipText ="Record ID for the date being edited"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1275
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =1590
                    ForeThemeColorIndex =0
                    ForeTint =50.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =120
                            Top =1275
                            Width =2460
                            Height =315
                            FontWeight =500
                            Name ="lblOrigDateTime"
                            Caption ="Original Event_Date"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1275
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1590
                        End
                    End
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
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxUpdateByID"
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
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =900
                    Top =120
                    Width =720
                    Height =285
                    ColumnOrder =1
                    FontSize =9
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxLastUpdate"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =120
                    LayoutCachedWidth =1620
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
' MODULE:       frm_Pad_Date
' Level:        Application module
' Version:      1.00
'
' Description:  date popup related functions & procedures
'
' Source/date:  Bonnie Campbell, May 1, 2020
' Adapted:      -
' Revisions:    BLC - 5/1/2020 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

Private m_OriginalDate As Date
Private m_EditByID As String
Private m_EditTable As String
Private m_EditField As String
Private m_EditIDField As String
Private m_EditID As String
Private m_Edit As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        'Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'editing info
Public Property Let EditByID(Value As String)
    If Not IsNull(Value) Then
        m_EditByID = Value
    Else
        m_EditByID = XML_Read("UpdateByID", Me.OpenArgs)
    End If
    Me.tbxUpdateByID = m_EditByID
End Property

Public Property Get EditByID() As String
    EditTable = m_EditTable
End Property

Public Property Let OriginalDate(Value As Date)
    If Not IsNull(Value) Then
        m_OriginalDate = Value
    Else
        m_OriginalDate = XML_Read("ControlValue", Me.OpenArgs)
    End If
    Me.tbxOrigDateTIme = m_OriginalDate
End Property

Public Property Get OriginalDate() As Date
    OriginalDate = m_OriginalDate
End Property

Public Property Let EditTable(Value As String)
    If Not IsNull(Value) Then
        m_EditTable = Value
    Else
        m_EditTable = XML_Read("EditTable", Me.OpenArgs)
    End If
    Me.tbxTable = m_EditTable
End Property

Public Property Get EditTable() As String
    EditTable = m_EditTable
End Property

Public Property Let EditField(Value As String)
    If Not IsNull(Value) Then
        m_EditField = Value
    Else
        m_EditField = XML_Read("EditField", Me.OpenArgs)
    End If
    Me.lblOrigDateTime.Caption = "Original " & m_EditField
End Property

Public Property Get EditField() As String
    EditField = m_EditField
End Property

Public Property Let EditIDField(Value As String)
    If Not IsNull(Value) Then
        m_EditIDField = Value
    Else
        m_EditIDField = XML_Read("EditIDField", Me.OpenArgs)
    End If
    Me.lblEditing.Caption = m_EditIDField
End Property

Public Property Get EditIDField() As String
    EditIDField = m_EditIDField
End Property

Public Property Let EditID(Value As String)
    If Not IsNull(Value) Then
        m_EditID = Value
    Else
        m_EditID = XML_Read("EditID", Me.OpenArgs)
    End If
    Me.tbxID = m_EditID
End Property

Public Property Get EditID() As String
    EditID = m_EditID
End Property

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
'    Me.CallingForm = "Main"
'
'    If Len(Me.OpenArgs) > 0 Then Me.CallingForm = Me.OpenArgs
'
'    'minimize calling form
'    ToggleForm Me.CallingForm, -1
    
    'dev mode
    tbxDevMode = DEV_MODE
                
    Title = "Edit Date"
    'lblTitle.Caption = "" 'clear header title
    Directions = "Choose the desired date. " _
              & "Click Save to save the new date. " _
              & "Changing the date WILL change underlying data."
    
    'defaults
    lblDirections.forecolor = lngBlue
    btnSave.hoverColor = lngGreen
    btnCancel.hoverColor = lngRed
       
    'set values
    EditTable = XML_Read("EditTable", Me.OpenArgs)
    EditField = XML_Read("EditField", Me.OpenArgs)
    EditIDField = XML_Read("EditIDField", Me.OpenArgs)
    EditByID = XML_Read("UpdateByID", Me.OpenArgs)
    EditID = XML_Read("EditID", Me.OpenArgs)
    OriginalDate = XML_Read("ControlValue", Me.OpenArgs)
Debug.Print Me.OpenArgs
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Pad_Date form])"
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

    'Me.RecordSource = Me.EditTable
    'Me.tbxDate.ControlSource = Me.EditField
    
    Debug.Print Me.EditField
    
    'Me.FilterOnLoad = True
    'Me.Filter = Me.EditIDField & "=" & Me.EditID
    'Me.FilterOn = True
        
'    Me.tbxUpdateByID = Me.EditByID
'    Me.tbxTable = Me.EditTable
'    Me.lblID.Caption = Me.EditIDField
'    Me.tbxID = Me.EditID
'    Me.lblOrigDateTime.Caption = "Original " & Me.EditField
'    Me.tbxOrigDateTIme = Me.OriginalDate
    
    'save the time from the date field
    Dim strOrigDateTime As String
    strOrigDateTime = XML_Read("ControlValue", Me.OpenArgs)

    'set new date based on old
    Me.tbxDate = strOrigDateTime

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[frm_Pad_Date])"
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
            "Error encountered (#" & Err.Number & " - Form_Current[frm_Pad_Date])"
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
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[frm_Pad_Date])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDate_GotFocus
' Description:  textbox actions when the box is clicked
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2019
' Adapted:      -
' Revisions:
'   BLC - 3/18/2019 - initial version
' ---------------------------------
Private Sub tbxDate_GotFocus()
On Error GoTo Err_Handler

    tbxDate.forecolor = lngBlack
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDate_GotFocus[frm_Pad_Date])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDate_AfterUpdate
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
Private Sub tbxDate_AfterUpdate()
On Error GoTo Err_Handler
    

    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDate_AfterUpdate[frm_Pad_Date])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxPlotName_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
'   BLC - 4/2/2019 - revised to use ConvertEvent
' ---------------------------------
Private Sub tbxPlotName_Click()
On Error GoTo Err_Handler
    
'    ConvertEvent
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPlotName_Click[frm_Pad_Date])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxEventDate_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version
'   BLC - 4/2/2019 - revised to use ConvertEvent
' ---------------------------------
Private Sub tbxEventDate_Click()
On Error GoTo Err_Handler
    
    'ConvertEvent
    
'    Dim strCriteriaLoc As String
'    Dim strCriteriaEvent As String

    'Record what the current record is so we can go back to that record on return
    'WriteRecordCriteria
    
    'NCRN NOTE: For this database, we will not create new events through this mechanism.
    'It is unclear to me how to use this mechanism to create a second event for a location (mel).
    
    'If there is not an event id, add a new data entry record
    'If IsNull(Me!txtEvent_ID) Then
    '            DoCmd.OpenForm "frm_Events", , , , acFormAdd, , "New record"
    '    If Not IsNull(Me!txtLocation_ID) Then
    '        ' Fill in Location
    '        Forms!frm_Events!cboLocation_ID = Me!txtLocation_ID
    '        Forms!frm_Events.Update_Loc_Info
    '    End If
    'if there is an event id, bring up the selected data entry record
    'Else
        'strCriteriaLoc = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
'        strCriteriaEvent = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "tbxEventID")
        ' Filter by location and event
'        DoCmd.OpenForm "frm_Events", , , strCriteriaEvent, , , "(Browsing)"
    'End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEventDate_Click[frm_Pad_Date])"
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

    Dim SQL As String

    'save history
    SaveHistory Me.tbxTable, Me.EditField, Me.lblEditing.Caption, Me.tbxID, _
                Me.tbxOrigDateTIme, Me.tbxDate, Me.tbxUpdateByID, "", Date, Date, "Event date change"

    'identify update being made
    Select Case Me.EditTable
        Case "tbl_Events"
            SQL = "UPDATE table_Events SET " & Me.EditField & " = #" & Me.tbxDate & "#, Updated_By = " & Me.tbxUpdateByID _
             & ", Updated_Date = #" & Date & "# WHERE " & Me.EditIDField & "= " & Me.EditID & ";"
    End Select
    
    Debug.Print SQL
    
    'update event date
    'DoCmd.RunSQL ""

    DoCmd.Close acForm, Me.Name
    

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[frm_Pad_Date])"
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

    'Close the Create Event form without creating a record

'    If Me.Dirty Then Me.Undo
'    If Not Me.NewRecord Then
'        DoCmd.RunCommand acCmdDeleteRecord
'    End If
Debug.Print "Dirty = " & Me.Dirty
Debug.Print "NewRec = " & Me.NewRecord

    'remove new record if created
    If Me.Dirty Then Me.Undo
    If Not Me.NewRecord = True Then
        DoCmd.RunCommand acCmdDelete
    End If
    
    DoCmd.Close
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[frm_Pad_Date])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SaveHistory
' Description:  saves relevant changes to tbl_History
' Assumptions:
'       tbl_History includes the following fields
'                   Table_Name (short text) - Changed record's table name
'                   Record_ID_Field_Name (text) - Changed record's ID field name
'                   Record_ID (text) - Changed record's ID
'                   Field_Name (text) - Changed field's name
'                   Value_New (text) - Changed field's new value
'                   Value_Old (text) - Changed field's old value
'                   Value_History_Notes (text) - Comments about the change
'                   Contact_ID (text) - Contact ID for person doing the change
'                   Network_User_Name (text) - Network username for person doing the change
'                   Change_Date (datetime) - Date the field value was changed
'                   LastUpdate (datetime, default Now) - Date of record creation or update
'       Values to populate these record fields are either passed in, stored in TempVars,
'       or otherwise accessible to this subroutine.
' Parameters:
'       EditTable (string) - Changed record's table name
'       RecordIDField (string) - Changed record's ID field name
'       RecordID (string) - Changed record's ID
'       FieldName (string) - Changed field's name
'       Value_New (string) - Changed field's new value
'       Value_Old (string) - Changed field's old value
'       Value_History_Notes (string) - Comments about the change
'       UpdatedByID (string) - Contact ID for person doing the change
'       NetworkUserName (string, default Environ("username")) - Network username for person doing the change
'       Change_Date (datetime) - Date the field value was changed
'       LastUpdate (datetime, default Now) - Date of record creation or update
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 1, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/1/2020 - initial version
' ---------------------------------
Private Sub SaveHistory(EditTable As String, EditField As String, _
                        RecordIDField As String, RecordID As String, OldValue As String, NewValue As String, _
                        UpdatedByID As String, NetworkUsername As String, _
                        ChangeDate As Date, LastUpdate As Date, _
                        Optional UpdateNotes As String = "")
On Error GoTo Err_Handler

    Dim SQL As String
    Dim strVals As String
    
    'set defaults
    NetworkUsername = IIf(Len(NetworkUsername) = 0, Nz(Environ("Username"), ""), NetworkUsername)
    ChangeDate = IIf(IsDate(ChangeDate) = False, Date, ChangeDate)
    LastUpdate = IIf(IsDate(LastUpdate) = False, Date, LastUpdate)
    
    'populate fields
    strVals = "'" & EditTable & "', '" & RecordIDField & "', '" & RecordID & "', " _
               & "'" & EditField & "', '" & NewValue & "', '" & OldValue & "', " _
               & "'" & UpdateNotes & "'," _
               & "'" & UpdatedByID & "', '" & NetworkUsername & "', " _
               & "#" & ChangeDate & "#, #" & LastUpdate & "#"
    SQL = "INSERT INTO tbl_History(Table_Name, Record_ID_Field_Name, Record_ID, Field_Name, Value_New, Value_Old, " _
          & "Value_History_Notes, Contact_ID, Network_User_Name, Change_Date, LastUpdate) VALUES (" & strVals & ");"
    
    Debug.Print SQL
    
    DoCmd.RunSQL SQL
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SaveHistory[frm_Pad_Date])"
    End Select
    Resume Exit_Handler
End Sub
