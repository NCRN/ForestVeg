Version =21
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7620
    DatasheetFontHeight =11
    ItemSuffix =51
    Left =6105
    Top =3315
    Right =13980
    Bottom =7380
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x287f6d81dc7be540
    End
    RecordSource ="TargetSpecies"
    Caption ="Add Target Species"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
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
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =120
                    Width =7260
                    Height =720
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Directions"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3540
                    Top =960
                    Width =825
                    Height =345
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSpecies"
                    Caption ="Species"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =3540
                    LayoutCachedTop =960
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =1305
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =960
                    Width =1065
                    Height =345
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTargetList"
                    Caption ="Target List"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =960
                    LayoutCachedWidth =2265
                    LayoutCachedHeight =1305
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1230
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =45
                    LayoutCachedWidth =480
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =45
                    LayoutCachedWidth =960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3240
                    Top =540
                    Width =1020
                    Height =315
                    FontSize =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxEffectiveDate"
                    ControlSource ="EstablishDate"
                    Format ="Short Date"
                    FontName ="Franklin Gothic Book"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000b0000000020000000100000000000000000000002300000001000000 ,
                        0x00000000fff20000000000000300000024000000270000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028004c0065006e00280022007400620078004500660066006500 ,
                        0x630074006900760065004400610074006500220029003d0030002c0031002c00 ,
                        0x30002900000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =540
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =855
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000220000004900 ,
                        0x4900660028004c0065006e002800220074006200780045006600660065006300 ,
                        0x74006900760065004400610074006500220029003d0030002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000000000000300000001 ,
                        0x00000000000000ffffff00020000002200220000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6780
                    Top =60
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Save record"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0687050c06860ffb05850ffa05050ffa05050ff ,
                        0xa05050ff904850ff904840ff904840ff804040ff803840ff803840ff703840ff ,
                        0x703830ff0000000000000000d06870fff09090ffe08080ffb04820ff403020ff ,
                        0xc0b8b0ffc0b8b0ffd0c0c0ffd0c8c0ff505050ffa04030ffa04030ffa03830ff ,
                        0x703840ff0000000000000000d07070ffff98a0fff08880ffe08080ff705850ff ,
                        0x404030ff907870fff0e0e0fff0e8e0ff908070ffa04030ffa04040ffa04030ff ,
                        0x803840ff0000000000000000d07870ffffa0a0fff09090fff08880ff705850ff ,
                        0x000000ff404030fff0d8d0fff0e0d0ff807860ffb04840ffb04840ffa04040ff ,
                        0x804040ff0000000000000000d07880ffffa8b0ffffa0a0fff09090ff705850ff ,
                        0x705850ff705850ff705850ff706050ff806860ffc05850ffb05050ffb04840ff ,
                        0x804040ff0000000000000000e08080ffffb0b0ffffb0b0ffffa0a0fff09090ff ,
                        0xf08880ffe08080ffe07880ffd07070ffd06870ffc06060ffc05850ffb05050ff ,
                        0x904840ff0000000000000000e08890ffffb8c0ffffb8b0ffd06060ffc06050ff ,
                        0xc05850ffc05040ffb05030ffb04830ffa04020ffa03810ffc06060ffc05850ff ,
                        0x904840ff0000000000000000e09090ffffc0c0ffd06860ffffffffffffffffff ,
                        0xfff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffe0c8c0ffa03810ffc06060ff ,
                        0x904850ff0000000000000000e098a0ffffc0c0ffd07070ffffffffffffffffff ,
                        0xfffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffa04020ffd06860ff ,
                        0xa05050ff0000000000000000f0a0a0ffffc0c0ffe07870ffffffffffffffffff ,
                        0xfffffffffffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffb04830ffd07070ff ,
                        0xa05050ff0000000000000000f0a8a0ffffc0c0ffe08080ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffff8f0fff0f0f0fff0e8e0ffb05030ffe07880ff ,
                        0xa05050ff0000000000000000f0b0b0ffffc0c0fff08890ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff8f0fff0f0f0ffc05040ff603030ff ,
                        0xb05850ff0000000000000000f0b0b0ffffc0c0ffff9090ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff8f0ffc05850ffb05860ff ,
                        0xb05860ff0000000000000000f0b8b0fff0b8b0fff0b0b0fff0b0b0fff0a8b0ff ,
                        0xf0a0a0ffe098a0ffe09090ffe09090ffe08890ffe08080ffd07880ffd07870ff ,
                        0xd07070ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =60
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =420
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =540
                    Width =1920
                    Height =300
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblEffectiveDate"
                    Caption ="Effective Date"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =540
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =840
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =10
                    Left =3600
                    Top =60
                    Width =3060
                    Height =285
                    FontSize =9
                    TabIndex =4
                    BoundColumn =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =2171426
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";"
                        "\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"48\""
                    ConditionalFormat = Begin
                        0x0100000094000000020000000100000000000000000000001500000001000000 ,
                        0x00000000fff20000000000000300000016000000190000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x4c0065006e0028002200740062007800540065006d0070006c00610074006500 ,
                        0x220029003d003000000000002200220000000000
                    End
                    Name ="cbxSpecies"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Plants"
                    ColumnWidths ="0;0;0;0;0;0;0;0;0;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =345
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000140000004c00 ,
                        0x65006e0028002200740062007800540065006d0070006c006100740065002200 ,
                        0x29003d0030000000000000000000000000000000000000000000000000000003 ,
                        0x0000000100000000000000ffffff000200000022002200000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =480
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTSN"
                    ControlSource ="TSN"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =480
                    LayoutCachedWidth =960
                    LayoutCachedHeight =795
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =2280
                    Height =285
                    FontSize =9
                    TabIndex =6
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =2171426
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x0100000094000000020000000100000000000000000000001500000001000000 ,
                        0x00000000fff20000000000000300000016000000190000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x4c0065006e0028002200740062007800540065006d0070006c00610074006500 ,
                        0x220029003d003000000000002200220000000000
                    End
                    Name ="cbxTargetList"
                    RowSourceType ="Table/Query"
                    RowSource ="s_target_list"
                    ColumnWidths ="1440"
                    FontName ="Franklin Gothic Book"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =345
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000140000004c00 ,
                        0x65006e0028002200740062007800540065006d0070006c006100740065002200 ,
                        0x29003d0030000000000000000000000000000000000000000000000000000003 ,
                        0x0000000100000000000000ffffff000200000022002200000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =915
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTargetList"
                    ControlSource ="TargetList"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =915
                    LayoutCachedWidth =960
                    LayoutCachedHeight =1230
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Form:         TargetSpeciesAdd
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  Add record form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, June 16, 2020
' References:   -
' Revisions:    BLC - 10/4/2016 - 1.00 - initial version
'               BLC - 1/31/2017 - 1.01 - adjusted to set ID value = 0 to signal template
'                                        add vs. update
'               BLC - 2/1/2017  - 1.02 - adjusted to set hidden context value,
'                                        revised to unbound form (avoids template double save)
'                                        adjust to retrieve valid syntaxes from AppEnum
'                                        added CallingForm property
'               BLC - 2/2/2017 - 1.03 - truncate template SQL to 255 chars (others need to be added to table directly)
'                                       add character counter & tbxTemplate_Change()
'                                       added textbox & combobox change events
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

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
        Me.Caption = m_Title
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
    If Len(value) > 0 Then
        m_CallingForm = value
    Else
        RaiseEvent InvalidCallingForm(value)
    End If
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Microsoft, unknown
'   https://msdn.microsoft.com/en-us/library/office/aa223974(v=office.11).aspx
' Source/date:  Bonnie Campbell, October 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/4/2016 - initial version
'   BLC - 2/1/2017 - revised to unbound form (avoids double save of template)
'                    adjust to retrieve valid syntaxes from AppEnum
'                    added CallingForm default, adjusted directions
'   BLC - 2/2/2017 - handle Template memo truncation since > 255 yields
'                    error 3271 SetRecord mod_App_Data  Invalid property value.
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "TargetLists"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize Calling Form
    ToggleForm Me.CallingForm, -1

    Me.Caption = "Add Target Species"
    lblTitle.Caption = ""
    lblDirections.Caption = "Select the target list, species, and enter" _
                                & "the date target species monitoring began. " _
                                & "Then click save to add the target species to the list."
    tbxIcon.value = StringFromCodepoint(uLocked)
    tbxIcon.ForeColor = lngDkGreen
    lblDirections.ForeColor = lngLtBlue
    
    'set hover
    btnSave.HoverColor = lngGreen
    
    'set syntax values
    'SetTempVar "EnumType", "SyntaxType"
'    Set cbxTargetList.Recordset = GetRecords("s_app_enum_list")
'    cbxTargetList.ColumnCount = 1
'    cbxTargetList.BoundColumn = 1
'    cbxTargetList.ColumnWidths = "1;"
'    cbxTargetList.Value = ""
    
    'defaults
    btnSave.Enabled = False
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/4/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 16, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/16/2020 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxTargetList_Change
' Description:  Combobox actions on change event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 16, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/16/2020 - initial version
' ---------------------------------
Private Sub cbxTargetList_Change()
On Error GoTo Err_Handler

    ReadyForSave
    Me.tbxTargetList = cbxTargetList

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTargetList_Change[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSpecies_Change
' Description:  Textbox actions on change event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/2/2017 - initial version
' ---------------------------------
Private Sub cbxSpecies_Change()
On Error GoTo Err_Handler

    ReadyForSave

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSpecies_Change[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxEffectiveDate_Change
' Description:  Textbox actions on change event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 2, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 2/2/2017 - initial version
' ---------------------------------
Private Sub tbxEffectiveDate_Change()
On Error GoTo Err_Handler

    ReadyForSave

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEffectiveDate_Change[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSpecies_AfterUpdate
' Description:  combobox actions after update event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 16, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/16/2020 - initial version
' ---------------------------------
Private Sub cbxSpecies_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    Me.tbxTSN = Me.cbxSpecies

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSpecies_AfterUpdate[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSave_Click
' Description:  Save button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 1/31/2017 - adjusted to set ID value = 0 to signal template add vs. update
'   BLC - 2/1/2017  - adjusted to set hidden context value, removed Me.IsSupported
'                     value is defaulted to 1 for new templates in UpsertRecord
'                     toggle list form (calling form)
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'default IsSupported is 1 (set in UpsertRecord)
       
    'clear tbxID so templates are considered adds, not updates
'    Me.tbxID.Value = 0
'
'    UpsertRecord Me

    
'FIX --> adjust so that when form goes to calling form (TemplateList)
'        that list isn't calling same routine Form_Current
'        multiple times when transitioning back (get focus?)

    'minimize form (calling form will close it)
    ToggleForm Me.Name, -1
    
    'restore list & refresh
    ToggleForm Me.CallingForm, 0
    
    Forms(Me.CallingForm).Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 16, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/16/2020 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore Target Species List
    ToggleForm "TargetLists", 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyForSave
' Description:  Check if form values are ready to save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: site code & name (directions & description optional)
    If Len(Nz(cbxTargetList.value, "")) > 0 _
        And Len(Nz(cbxSpecies.value, "")) > 0 _
        And Len(Nz(tbxEffectiveDate.value, "")) > 0 _
        Then
        isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    'Me.Requery
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[TargetSpeciesAdd form])"
    End Select
    Resume Exit_Handler
End Sub
