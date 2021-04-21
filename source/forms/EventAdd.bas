Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =9
    ItemSuffix =29
    Left =5745
    Top =165
    Right =10065
    Bottom =5550
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xbfc854d06f46e540
    End
    RecordSource ="tbl_Events"
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
        Begin Section
            Height =5400
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =240
                    Top =2820
                    Width =3840
                    Height =1260
                    BackColor =13754087
                    BorderColor =10921638
                    Name ="rctPseudoEvent"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =2820
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =4080
                    BackThemeColorIndex =-1
                    BackTint =40.0
                End
                Begin ComboBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =1485
                    Top =1620
                    Width =2475
                    Height =510
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cbxLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, tbl_Locations.Panel, "
                        "tbl_Locations.Frame, tbl_Locations.Unit_Code FROM tbl_Locations WHERE (((tbl_Loc"
                        "ations.Panel)=[Forms]![frm_Switchboard]![Panel])) ORDER BY tbl_Locations.Plot_Na"
                        "me;"
                    ColumnWidths ="0;2160"
                    AfterUpdate ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1485
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2130
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =1620
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPlot"
                            Caption ="Plot"
                            LayoutCachedLeft =540
                            LayoutCachedTop =1620
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =2135
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1500
                    Top =2220
                    Width =2460
                    Height =510
                    FontSize =18
                    FontWeight =700
                    TabIndex =2
                    Name ="tbxEventDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=Date()"
                    ControlTipText ="Click in this field & use the date picker that appears to set the date"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2220
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2730
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =2220
                            Width =885
                            Height =510
                            FontSize =18
                            FontWeight =700
                            Name ="lblEventDate"
                            Caption ="Date"
                            LayoutCachedLeft =540
                            LayoutCachedTop =2220
                            LayoutCachedWidth =1425
                            LayoutCachedHeight =2730
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =4320
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =275078
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Create New Event"
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =540
                    BackThemeColorIndex =5
                    BackShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1275
                    Top =600
                    Width =2595
                    Height =210
                    ColumnWidth =1320
                    FontSize =8
                    TabIndex =5
                    Name ="tbxEventID"
                    StatusBarText ="M. Event identifier (Event_ID)"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =600
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =810
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =240
                            Top =600
                            Width =690
                            Height =240
                            FontSize =8
                            Name ="lblEventID"
                            Caption ="Event ID:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =600
                            LayoutCachedWidth =930
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    Left =420
                    Top =4200
                    Width =2325
                    Height =1080
                    FontSize =14
                    TabIndex =3
                    ForeColor =0
                    Name ="btnCreate"
                    Caption ="Create Event"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =420
                    LayoutCachedTop =4200
                    LayoutCachedWidth =2745
                    LayoutCachedHeight =5280
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
                    OverlapFlags =95
                    TextFontCharSet =204
                    Left =2820
                    Top =4200
                    Width =1020
                    Height =1080
                    FontSize =14
                    TabIndex =4
                    ForeColor =0
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2820
                    LayoutCachedTop =4200
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =5280
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =2100
                    Top =600
                    Height =315
                    TabIndex =6
                    Name ="tbxProtocolName"
                    DefaultValue ="=[Forms]![frm_Switchboard]![Protocol_Name]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =600
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =960
                            Top =600
                            Width =1080
                            Height =315
                            Name ="lblProtocolName"
                            Caption ="Protocol:"
                            LayoutCachedLeft =960
                            LayoutCachedTop =600
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =1485
                    Top =1020
                    Width =2475
                    Height =510
                    FontSize =18
                    FontWeight =700
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxParkCode"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Unit Code\")) ORDER BY tlu_Enumerations.Enum_Code;"
                    ColumnWidths ="2160"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""

                    LayoutCachedLeft =1485
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1530
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =1020
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPark"
                            Caption ="Park"
                            LayoutCachedLeft =540
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =1535
                        End
                    End
                End
                Begin Label
                    OverlapFlags =223
                    Left =360
                    Top =3360
                    Width =3600
                    Height =660
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintPseudoEvent"
                    Caption ="Bush-hogged or other non-data collecting visit that may impact analysis"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Bush-hogged or other non-data collecting visit that may impact analysis"
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =3360
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =4020
                    ThemeFontIndex =1
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =420
                    Top =2940
                    Width =270
                    Height =299
                    TabIndex =7
                    Name ="tglPseudoEvent"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Bush-hogged or other non-data collecting visit record?"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =2940
                    LayoutCachedWidth =690
                    LayoutCachedHeight =3239
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    Gradient =12
                    BackColor =8289145
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =8289145
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =780
                            Top =2940
                            Width =2160
                            Height =315
                            BorderColor =8355711
                            ForeColor =16711680
                            Name ="lblPseudoEvent"
                            Caption ="Pseudo Event?"
                            FontName ="Franklin Gothic Book"
                            ControlTipText ="Bush-hogged or other non-data collecting visit that may impact analysis"
                            GridlineColor =10921638
                            LayoutCachedLeft =780
                            LayoutCachedTop =2940
                            LayoutCachedWidth =2940
                            LayoutCachedHeight =3255
                            ThemeFontIndex =1
                            BackThemeColorIndex =1
                            BorderThemeColorIndex =0
                            BorderTint =50.0
                            ForeTint =50.0
                            GridlineThemeColorIndex =1
                            GridlineShade =65.0
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3900
                    Top =600
                    Width =420
                    Height =300
                    FontSize =9
                    TabIndex =8
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

                    LayoutCachedLeft =3900
                    LayoutCachedTop =600
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =900
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
                    Visible = NotDefault
                    OverlapFlags =247
                    Width =3360
                    Height =615
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="dirs"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =615
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
                    Left =3240
                    Top =2940
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxPseudoEvent"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2940
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =3240
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =4020
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxEID"
                    ControlSource ="Event_ID"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =4020
                    LayoutCachedWidth =900
                    LayoutCachedHeight =4320
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =3960
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxRecordPseudoEvent"
                    ControlSource ="PseudoEvent"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =3960
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =4260
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =4440
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxRecordLocationID"
                    ControlSource ="Location_ID"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =4440
                    LayoutCachedWidth =900
                    LayoutCachedHeight =4740
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =4860
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxRecordEventDate"
                    ControlSource ="Event_Date"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =4860
                    LayoutCachedWidth =900
                    LayoutCachedHeight =5160
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ThemeFontIndex =1
                    ForeTint =50.0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =1140
                    Width =420
                    Height =315
                    TabIndex =14
                    Name ="tbxPark"
                    ControlSource ="Location_ID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1140
                    LayoutCachedWidth =480
                    LayoutCachedHeight =1455
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =1740
                    Width =420
                    Height =315
                    TabIndex =15
                    Name ="tbxPlot"
                    ControlSource ="Location_ID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1740
                    LayoutCachedWidth =480
                    LayoutCachedHeight =2055
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =2175
                    Width =420
                    Height =315
                    TabIndex =16
                    Name ="tbxDate"
                    ControlSource ="Event_Date"

                    LayoutCachedLeft =60
                    LayoutCachedTop =2175
                    LayoutCachedWidth =480
                    LayoutCachedHeight =2490
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
' MODULE:       EventAdd
' Level:        Application module
' Version:      1.03
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/5/2018 - 1.01 - added documentation, error handling
'               BLC   - 10/23/2018 - 1.02 - added Form_Open event, PseudoEvent handling
'               BLC   - 3/18/2019 - 1.03 - accommodate calling form park code
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
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
'   BLC - 3/18/2019  - accommodate passing park code from calling form
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
                
    Title = "Create New Event"
    'lblTitle.Caption = "" 'clear header title
    Directions = "dirs"
    
    'defaults
    rctPseudoEvent.BackColor = lngLtTan
    
    'disable until data allows
    cbxLocationID.Enabled = False
    tbxEventDate.Enabled = False
    tglPseudoEvent.Enabled = False
    btnCreate.Enabled = False
    
    'hints
    lblPseudoEvent.Caption = "Pseudo Event?"
    lblPseudoEvent.ForeColor = lngBlue
    lblPseudoEvent.ControlTipText = "Bush-hogged or other non-data collecting visit that may impact analysis"
    lblPseudoEvent.visible = True
    lblHintPseudoEvent.Caption = "Bush-hogged or other non-data collecting visit that may impact analysis"
    lblHintPseudoEvent.ForeColor = lngBlue
    lblHintPseudoEvent.ControlTipText = "Bush-hogged or other non-data collecting visit that may impact analysis"
    lblHintPseudoEvent.visible = True
    
    'set hover
    tglPseudoEvent.HoverColor = lngGreen
       
    'initialize values
    ClearForm Me

    'set the open record
    If Len(Me.tbxEID.Value) = 0 Then
    Debug.Print Me.Name
        DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
        Debug.Print "eid=" & Me.tbxEID
    End If
    
    'set park code
    If IsEmpty(Me.OpenArgs) = False And Me.OpenArgs <> "Choose Park" Then
        Me.cbxParkCode = Me.OpenArgs
        Me.cbxLocationID.Enabled = True
        SetPlots Nz(Me.OpenArgs, "")
    Else
        Me.cbxParkCode = Me.OpenArgs
        Me.cbxLocationID.Enabled = False
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[EventAdd form])"
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
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'go to a new record
'    DoCmd.GoToRecord , , acNewRec

    DoCmd.GoToRecord Record:=acNewRec

    'Generate string GUID for Event_ID
    If Me.NewRecord = True Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me!Event_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[EventAdd])"
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
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    NewRecordMark Me.Form

    'am i dirty?
    If Me.Dirty Then
        MsgBox "DIRTY!", vbOKCancel, "Dirt!"
    Else
        MsgBox "NO DIRT", vbOKCancel, "No Dirt :O("
    End If

Debug.Print "Dirty = " & Me.Dirty
Debug.Print "NewRec = " & Me.NewRecord

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[EventAdd])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    'Generate string GUID for Event_ID
    If Me.NewRecord Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me!Event_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxParkCode_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cboPark_Code > cbxParkCode
'               BLC   - 10/23/2018 - revised to avoid error #2448 "can't assign value to this object"
' ---------------------------------
Private Sub cbxParkCode_AfterUpdate()
On Error GoTo Err_Handler

    SetPlots cbxParkCode
'    Me.cbxLocationID.RowSource = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, " _
'            & "tbl_Locations.Panel, tbl_Locations.Frame, tbl_Locations.Unit_Code " _
'            & "FROM tbl_Locations " _
'            & "WHERE (((tbl_Locations.Panel) = [Forms]![frm_Switchboard]![Panel]) " _
'            & "AND ((tbl_Locations.Unit_Code) = '" & Me.cbxParkCode & "')) " _
'            & "ORDER BY tbl_Locations.Plot_Name;"
'
'    'enable plot
'    cbxLocationID.Enabled = True
'
'    'set focus on next field
'    cbxLocationID.SetFocus
'
'    'Me.cbxLocationID = Me.cbxLocationID.ItemData(0) #Error 2448 - can't assign value to this object

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkCode_AfterUpdate[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxLocationID_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:    BLC   - 10/23/2018 - initial version
' ---------------------------------
Private Sub cbxLocationID_AfterUpdate()
On Error GoTo Err_Handler

    'set record value
    tbxRecordLocationID = cbxLocationID
    
    'set the location
    tbxPlot = tbxRecordLocationID

    'check
    ReadyForSave
    
    'set focus on next field
    tbxEventDate.SetFocus
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxLocationID_AfterUpdate[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxEventDate_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Private Sub tbxEventDate_AfterUpdate()
On Error GoTo Err_Handler
    
    'set record value
    tbxRecordEventDate = tbxEventDate
    
    'set the event date
    tbxDate = tbxRecordEventDate
    
    'check
    ReadyForSave
    
    'set focus on button (vs. PseudoEvent)
    btnCreate.SetFocus
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEventDate_AfterUpdate[EventAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglPseudoEvent_AfterUpdate
' Description:  Toggle button after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Private Sub tglPseudoEvent_AfterUpdate()
On Error GoTo Err_Handler

    'display as checkbox
    ToggleCaption tglPseudoEvent, True
    
    'set value for PseudoEvent
    Debug.Print "pse=" & tglPseudoEvent.Value
    tbxPseudoEvent.Value = CByte(Abs(tglPseudoEvent.Value))
    Debug.Print "tbxpse=" & tbxPseudoEvent.Value
    
    'set database value
    tbxRecordPseudoEvent.Value = CByte(Abs(tglPseudoEvent.Value))
    
    'check
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglPseudoEvent_AfterUpdate[EventAdd form])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cmdEvent_Create > btnCreate
'               BLC  - 10/23/2018 - added PseudoEvent handling
' ---------------------------------
Private Sub btnCreate_Click()
On Error GoTo Err_Handler

    'Save the new event if all of the needed information is provided, and open the Event form

    Dim strDocName As String
    Dim strLinkCriteria As String
    
    If IsNull(Me!cbxLocationID) Then
        MsgBox "You must select a location before you can enter record details!", _
            vbExclamation, "Enter Location First"
        Me!cbxLocationID.SetFocus
    Else
        If IsNull(Me!tbxEventDate) Then
            MsgBox "You must enter a date before you can enter record details!", _
                vbExclamation, "Enter Start Date"
            Me!tbxEventDate.SetFocus
        Else
            
    'Generate string GUID for Event_ID
    'If Me.NewRecord = True Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
'            Me!Event_ID = fxnGUIDGen
'            Me.tbxEID = Me!Event_ID
            Me.tbxEID = fxnGUIDGen
            Me.tbxEventID = Me.tbxEID
        End If
    'End If
            
Debug.Print "Dirty = " & Me.Dirty
Debug.Print "NewRec = " & Me.NewRecord
            
    If Me.Dirty = True Then
        DoCmd.RunCommand acCmdSaveRecord
    Else
        MsgBox "nothing to save"
    End If
'            DoCmd.RunCommand acCmdSaveRecord
        DoCmd.RunCommand acCmdSaveRecord
        
            'retrieve the EventID
Debug.Print "eid = " & Me.tbxEID 'tbxEventID
            
            strDocName = "frm_Events"
            strLinkCriteria = "[Event_ID]=" & "'" & Me![tbxEventID] & "'"
Debug.Print strLinkCriteria
 '           DoCmd.OpenForm strDocName, , , strLinkCriteria, , , " (Creating)," & Me.tbxEID
            DoCmd.OpenForm strDocName, , , strLinkCriteria, , , "(Browsing)"
            
            DoCmd.Close acForm, "EventAdd", acSavePrompt
            'DoCmd.Close acForm, "frm_Event_Add"
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCreate_Click[EventAdd])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cmdEvent_Cancel > btnCancel
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
            "Error encountered (#" & Err.Number & " - btnCancel_Click[EventAdd])"
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
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    If cbxLocationID.Value > 0 Then tbxEventDate.Enabled = True
    If IsDate(tbxEventDate.Value) Then tglPseudoEvent.Enabled = True

    If Len(Nz(cbxParkCode.Value, "")) > 0 _
        And IsGUID(cbxLocationID.Value) = True _
        And IsDate(tbxEventDate.Value) = True Then '_
        
        isOK = True
        
    End If
    
    'enable save button only for new sites (tbxID = 0)
'   If tbxID = 0 Then btnSave.Enabled = isOK
    
'    btnSubstrateCover.Enabled = IIf(tbxID.Value > 0, True, False)
'    btnSetObserverRecorder.Enabled = IIf(tbxID.Value > 0, True, False)
    
    'enable create if data is ok
    btnCreate.Enabled = isOK
    
    'refresh form
    Me.Requery
   
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[EventAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          IsGUID
' Description:  Check if value is a valid GUID
' Assumptions:
'               GUID is 32 hex digits grouped into chunks of 8-4-4-4-12
'               Regex is
'                   "^(\{){0,1}[0-9a-fA-F]{8}\-" & _
'                   "[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-" & _
'                   "[0-9a-fA-F]{12}(\}){0,1}$"
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Torbis, January 16, 2007
'   http://www.vbforums.com/showthread.php?447414-Solved-Check-if-string-is-Guid
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Public Function IsGUID(strInspect As String) As Boolean
On Error GoTo Err_Handler

    Dim strPattern As String
    strPattern = "^(\{){0,1}[0-9a-fA-F]{8}\-" & _
                 "[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-" & _
                 "[0-9a-fA-F]{12}(\}){0,1}$"

    IsGUID = IsRegExpMatch(strInspect, strPattern)
   
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsGUID[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

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
' FUNCTION:     SetPlots
' Description:  filter plots by park code
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
Public Function SetPlots(ParkCode As String)
On Error GoTo Err_Handler
    
    Me.cbxLocationID.RowSource = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, " _
            & "tbl_Locations.Panel, tbl_Locations.Frame, tbl_Locations.Unit_Code " _
            & "FROM tbl_Locations " _
            & "WHERE (((tbl_Locations.Panel) = [Forms]![frm_Switchboard]![Panel]) " _
            & "AND ((tbl_Locations.Unit_Code) = '" & ParkCode & "')) " _
            & "ORDER BY tbl_Locations.Plot_Name;"

    'enable plot
    cbxLocationID.Enabled = True
    
    'set focus on next field
    cbxLocationID.SetFocus
    
    'Me.cbxLocationID = Me.cbxLocationID.ItemData(0) #Error 2448 - can't assign value to this object
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetPlots[EventAdd])"
    End Select
    Resume Exit_Handler
End Function
