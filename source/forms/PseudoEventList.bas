Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowUpdating =2
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4440
    DatasheetFontHeight =9
    ItemSuffix =24
    Left =7905
    Top =1170
    Right =12345
    Bottom =7995
    DatasheetGridlinesColor =15062992
    Filter ="[Unit_Code]='ANTI' AND Year([Event_Date]) > 2017"
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
            Height =2700
            BackColor =14277338
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =4440
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =275078
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Create New Event"
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =540
                    BackThemeColorIndex =5
                    BackShade =50.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =120
                    Top =600
                    Width =4200
                    Height =960
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="dirs"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1560
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3720
                    Top =360
                    Width =420
                    Height =300
                    ColumnOrder =0
                    FontSize =9
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

                    LayoutCachedLeft =3720
                    LayoutCachedTop =360
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =660
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2220
                    Top =2400
                    Width =1665
                    Height =300
                    FontSize =12
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="lblEventDate"
                    Caption ="Sample Date*"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =2400
                    LayoutCachedWidth =3885
                    LayoutCachedHeight =2700
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =300
                    Top =2400
                    Width =1485
                    Height =300
                    FontSize =12
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="lblPlotName"
                    Caption ="Plot Name*"
                    OnDblClick ="[Event Procedure]"
                    LayoutCachedLeft =300
                    LayoutCachedTop =2400
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =2700
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =1365
                    Top =1740
                    Width =2475
                    Height =510
                    ColumnOrder =1
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxParkCode"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Unit Code\")) ORDER BY tlu_Enumerations.Enum_Code;"
                    ColumnWidths ="2160"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =1365
                    LayoutCachedTop =1740
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2250
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =1740
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPark"
                            Caption ="Park"
                            LayoutCachedLeft =420
                            LayoutCachedTop =1740
                            LayoutCachedWidth =1290
                            LayoutCachedHeight =2255
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Width =4440
                    Height =360
                    FontSize =13
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    Name ="tbxRecord"
                    ControlSource ="tbl_Events.PseudoEvent"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0xffccff00ffcdcd00000000000200000003000000050000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000300000000000
                    End

                    LayoutCachedWidth =4440
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x010002000000000000000200000001000000ffccff00ffcdcd00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2160
                    Width =1740
                    Height =300
                    FontSize =13
                    TabIndex =1
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =16711680
                    Name ="tbxEventDate"
                    ControlSource ="Event_Date"
                    Format ="dd-mmm-yyyy"
                    StatusBarText ="Start date of the sampling event"
                    OnClick ="[Event Procedure]"
                    ShowDatePicker =0

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Width =1500
                    Height =300
                    FontSize =13
                    TabIndex =2
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =16711680
                    Name ="tbxPlotName"
                    ControlSource ="Plot_Name"
                    StatusBarText ="Name of the location"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =420
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3300
                    Width =960
                    Height =300
                    FontSize =13
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =16711680
                    Name ="tbxEventID"
                    ControlSource ="tbl_Events.Event_ID"

                    LayoutCachedLeft =3300
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =300
                End
            End
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
                    Left =420
                    Top =60
                    Width =2325
                    Height =720
                    FontSize =14
                    ForeColor =0
                    Name ="btnCreate"
                    Caption ="Add New Event"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create a new event"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =2745
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =9226162
                    HoverThemeColorIndex =7
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
                    Left =2820
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

                    LayoutCachedLeft =2820
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =7775995
                    HoverThemeColorIndex =5
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
' MODULE:       PseudoEventList
' Level:        Application module
' Version:      1.01
'
' Description:  add event related functions & procedures
'
' Source/date:  Bonnie Campbell, February, 2019
' Adapted:      -
' Revisions:    BLC - 2/20/2019 - 1.00 - initial version
'               BLC - 4/17/2019 - 1.01 - update open pseudo event
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
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

                
    title = "Create New Event"
    'lblTitle.Caption = "" 'clear header title
    Directions = "Choose the park for your event. " _
              & "If a pseudo-event (rehab) was done in the last year " _
              & "for the desired plot, choose it and convert it to an actual sampling event. " _
              & "Click Add New Event to create a new event."
    
    'defaults
    lblDirections.ForeColor = lngBlack
'    rctPseudoEvent.BackColor = lngLtTan
    btnCreate.HoverColor = lngGreen
    btnCancel.HoverColor = lngRed
    
    cbxParkCode = "Choose park"
    cbxParkCode.ForeColor = lngLtGray
    
    'set the default filter (i.e. no pseudoevents)
    Me.FilterOnLoad = True
    FilterRecords ""
'    Me.Refresh
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PseudoEventList form])"
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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[PseudoEventList])"
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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    NewRecordMark Me.Form

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[PseudoEventList])"
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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblPlotName_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Private Sub lblPlotName_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Plot_Name")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPlotName_DblClick[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblEventDate_DblClick
' Description:  label double click actions
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
Private Sub lblEventDate_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Event_Date")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblEventDate_DblClick[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxParkFilter_GotFocus
' Description:  combobox actions when the box is clicked
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
Private Sub cbxParkCode_GotFocus()
On Error GoTo Err_Handler

    cbxParkCode.ForeColor = lngBlack
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkCode_GotFocus[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxParkFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Private Sub cbxParkCode_AfterUpdate()
On Error GoTo Err_Handler
    
'    Dim strFilter
'
'    'add park filter to filter string
'    strFilter = "[Unit_Code]='" & Me!cbxParkCode _
'             & "' AND Year([Event_Date]) > " & Year(Now()) - 2
''    strFilter = "[Unit_Code]='" & Me!cbxParkCode _
''             & "' AND Year([Event_Date]) > " & Year(Now()) - 2
'Debug.Print strFilter
'
'    Me.Filter = strFilter
'    'Me.Requery
'Debug.Print Me.Filter
'    DoCmd.ApplyFilter , strFilter

    FilterRecords Me!cbxParkCode
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkCode_AfterUpdate[PseudoEventList])"
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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
'   BLC - 4/2/2019 - revised to use ConvertEvent
' ---------------------------------
Private Sub tbxPlotName_Click()
On Error GoTo Err_Handler
    
    ConvertEvent
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPlotName_Click[PseudoEventList])"
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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
'   BLC - 4/2/2019 - revised to use ConvertEvent
' ---------------------------------
Private Sub tbxEventDate_Click()
On Error GoTo Err_Handler
    
    ConvertEvent
    
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
            "Error encountered (#" & Err.Number & " - tbxEventDate_Click[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCreate_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
' ---------------------------------
Private Sub btnCreate_Click()
On Error GoTo Err_Handler

    'open form not in new event record mode (acFormAdd) so it accepts form defaults
'    DoCmd.OpenForm "EventAdd", acNormal, , , , acWindowNormal, Me.cbxParkCode
    DoCmd.OpenForm "EventAdd2", acNormal, , , , acWindowNormal, Me.cbxParkCode
    DoCmd.Close acForm, Me.Name
    
'    'Save the new event if all of the needed information is provided, and open the Event form
'
'    Dim strDocName As String
'    Dim strLinkCriteria As String
'
'    If IsNull(Me!cbxLocationID) Then
'        MsgBox "You must select a location before you can enter record details!", _
'            vbExclamation, "Enter Location First"
'        Me!cbxLocationID.SetFocus
'    Else
'        If IsNull(Me!tbxEventDate) Then
'            MsgBox "You must enter a date before you can enter record details!", _
'                vbExclamation, "Enter Start Date"
'            Me!tbxEventDate.SetFocus
'        Else
'            DoCmd.RunCommand acCmdSaveRecord
'
'            'retrieve the EventID
''Debug.Print "eid = " & Me.tbxEID 'tbxEventID
'
'            strDocName = "frm_Events"
'            strLinkCriteria = "[Event_ID]=" & "'" & Me![tbxEventID] & "'"
''            DoCmd.OpenForm strDocName, , , strLinkCriteria, , , "(Creating)"
''            DoCmd.Close acForm, "frm_Event_Add"
'        End If
'    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCreate_Click[PseudoEventList])"
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
' Source/date:  Bonnie Campbell, February 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
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
            "Error encountered (#" & Err.Number & " - btnCancel_Click[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Sub


Sub NewRecordMark(frm As Form)
    Dim intnewrec As Integer
 
    intnewrec = frm.NewRecord
    If intnewrec = True Then
    MsgBox "You're in a new record." _
        & "@Do you want to add new data?" _
        & "@If not, move to an existing record."
    End If
End Sub

' ---------------------------------
' FUNCTION:     SortRecords
' Description:  sorts records by desired field
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  John R. Boetsch, May 5, 2006
'               Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling,
'                     renamed from fxnSortRecords
' ---------------------------------
Private Function SortRecords(ByVal strFieldName As String, _
    Optional ByVal strField2Name As String)
On Error GoTo Err_Handler
    
    Dim strSortField As String
    Dim strSortOrder As String
    Dim strOrderBy As String
    Dim strSortFieldLabel As String
        
    ' If already sorting in ascending order by this field, sort descending
    If strFieldName = strSortField And strSortOrder = "" Then
        strSortOrder = " DESC"
    Else: strSortOrder = ""
    End If
    
    ' Create the order by string and activate the filter
    strOrderBy = strFieldName & strSortOrder
    If strField2Name <> "" Then
        strOrderBy = strField2Name & " DESC, " & strOrderBy
    End If
    strSortField = strFieldName
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Change the label format to indicate the sorted field
    strSortFieldLabel = "lbl" & Replace(strFieldName, "_", "")
    With Me.Controls.item(strSortFieldLabel)
        .FontItalic = IIf(.FontItalic = False, True, False)
        .FontBold = IIf(.FontBold = False, True, False)
    
'        .FontItalic = False
'        .fontBold = False

'        .FontItalic = True
'        .fontBold = True
    End With
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortRecords[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FilterRecords
' Description:  filter records by park code
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
Public Function FilterRecords(ParkCode As String)
On Error GoTo Err_Handler
    
    Dim strFilter As String
        
    'set ParkCode to NONE to return no records if no park code passed
    ParkCode = IIf(IsEmpty(ParkCode), "NONE", ParkCode)
    
    'add park filter to filter string
    strFilter = "[Unit_Code]='" & ParkCode _
             & "' AND Year([Event_Date]) > " & year(Now()) - 2
'    strFilter = "[Unit_Code]='" & Me!cbxParkCode _
'             & "' AND Year([Event_Date]) > " & Year(Now()) - 2
Debug.Print strFilter

    Me.Filter = strFilter
    'Me.Requery
Debug.Print Me.Filter
    DoCmd.ApplyFilter , strFilter
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterRecords[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ConvertEvent
' Description:  convert pseudoevent to normal event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2019
' Adapted:      -
' Revisions:
'   BLC - 4/2/2019 - initial version
' ---------------------------------
Public Function ConvertEvent()
On Error GoTo Err_Handler
    
    'temporarily set UserID
    SetTempVar "UserID", 12345
    
    Dim ConvertIt As Integer
    ConvertIt = MsgBox("Click OK to confirm you would like to convert this PseudoEvent (rehab) to " _
      & "a regular event with today's date.", vbOKCancel, "Confirm PseudoEvent Conversion")
        
    Dim strCriteria As String
    strCriteria = "[Event_ID]='" & tbxEventID & "'"
            
    Select Case ConvertIt
        
        Case vbOK '1
            'manage conversion
           
            Dim varReturn As Variant
            Dim strSQL As String
            
            'status
            DoCmd.Hourglass True
            varReturn = SysCmd(acSysCmdSetStatus, "Converting pseudoevent...")
            
            'do conversion
            strSQL = "UPDATE tbl_Events e " _
                    & "SET e.PseudoEvent = 0, " _
                    & "e.Event_Date = Now(), " _
                    & "e.Updated_Date = Now(), " _
                    & "e.Updated_By = " & TempVars("UserID") & ", " _
                    & "e.Event_Notes = e.Event_Notes & CHR(13) & CHR(10) & CHR(13) & CHR(10) & 'Converted ' & e.Event_Date & ' rehab (pseudoevent)' " _
                    & "WHERE " _
                    & "e.Event_ID = '" & tbxEventID & "' " _
                    & "AND e.PseudoEvent = 1;"
                    
            DoCmd.RunSQL strSQL
            
            'status
            DoCmd.Hourglass False
            varReturn = SysCmd(acSysCmdSetStatus, "Pseudoevent converted...")
            varReturn = SysCmd(acSysCmdSetStatus, "Opening event...")
            varReturn = SysCmd(acSysCmdClearStatus)
        
        Case vbCancel '2
            
            'status
            varReturn = SysCmd(acSysCmdSetStatus, "Opening event...")
            varReturn = SysCmd(acSysCmdClearStatus)

    End Select
            
    'open the pseudoevent and close PseudoEventList form
    DoCmd.OpenForm "frm_Events", , , strCriteria, , , "Filter by event"
    DoCmd.Close acForm, "PseudoEventList", acSaveNo

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConvertEvent[PseudoEventList])"
    End Select
    Resume Exit_Handler
End Function
