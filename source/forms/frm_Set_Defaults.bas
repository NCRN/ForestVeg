Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6600
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =6240
    Top =1230
    Right =12840
    Bottom =6555
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa3c57b9aedcee240
    End
    RecordSource ="tsys_App_Defaults"
    Caption =" Set application default values"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =5340
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =60
                    Top =2880
                    Width =6480
                    Height =1560
                    TabIndex =16
                    BackColor =16709608
                    BorderColor =10921638
                    Name ="rctPaths"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =2880
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =4440
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =180
                            Top =2820
                            Width =1680
                            Height =300
                            FontWeight =600
                            TopMargin =58
                            Name ="lblPaths"
                            Caption ="Directory Paths"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =180
                            LayoutCachedTop =2820
                            LayoutCachedWidth =1860
                            LayoutCachedHeight =3120
                            BackThemeColorIndex =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =1182
                    Top =840
                    Width =1395
                    Height =252
                    FontSize =9
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboPanel"
                    ControlSource ="Panel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.[Enum_Group])=\"Sampling Panel\")) ORDER BY"
                        " tlu_Enumerations.Enum_Code; "
                    ColumnWidths ="720;5040"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =1182
                    LayoutCachedTop =840
                    LayoutCachedWidth =2577
                    LayoutCachedHeight =1092
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =435
                            Top =840
                            Width =660
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblPanel"
                            Caption ="Panel"
                            FontName ="Calibri"
                            LayoutCachedLeft =435
                            LayoutCachedTop =840
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1095
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1167
                    Top =120
                    Width =3180
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cboUser"
                    ControlSource ="User_name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Last_Name & \"_\" & First_Name FROM tlu_Contacts ORDER BY Last_Name, Firs"
                        "t_Name; "
                    FontName ="Calibri"
                    OnNotInList ="[Event Procedure]"

                    LayoutCachedLeft =1167
                    LayoutCachedTop =120
                    LayoutCachedWidth =4347
                    LayoutCachedHeight =372
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =432
                            Top =120
                            Width =663
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblUser"
                            Caption ="User"
                            FontName ="Calibri"
                            LayoutCachedLeft =432
                            LayoutCachedTop =120
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =372
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1182
                    Top =2340
                    Width =3168
                    Height =252
                    FontSize =9
                    TabIndex =6
                    Name ="cboProject"
                    ControlSource ="Project"
                    FontName ="Calibri"

                    LayoutCachedLeft =1182
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =2592
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =423
                            Top =2340
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblProject"
                            Caption ="Project"
                            FontName ="Calibri"
                            LayoutCachedLeft =423
                            LayoutCachedTop =2340
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =2592
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5760
                    Top =4560
                    Width =720
                    Height =354
                    FontSize =10
                    FontWeight =600
                    ForeColor =0
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5760
                    LayoutCachedTop =4560
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =4914
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3330
                    Top =480
                    Width =1035
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    Name ="cmdNewUser"
                    Caption ="New user"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Add a new user"

                    LayoutCachedLeft =3330
                    LayoutCachedTop =480
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =1182
                    Top =1620
                    Width =1395
                    FontSize =9
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboDatum"
                    ControlSource ="Datum"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"Datu"
                        "m\" ORDER BY Sort_Order; "
                    ColumnWidths ="720;2880"
                    FontName ="Calibri"

                    LayoutCachedLeft =1182
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2577
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =423
                            Top =1620
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblDatum"
                            Caption ="Datum"
                            FontName ="Calibri"
                            LayoutCachedLeft =423
                            LayoutCachedTop =1620
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1872
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2664
                    Left =1170
                    Top =1980
                    Width =1395
                    FontSize =9
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboUTM_Zone"
                    ControlSource ="UTM_Zone"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"UTM "
                        "Zone\" ORDER BY Sort_Order; "
                    ColumnWidths ="504;2160"
                    FontName ="Calibri"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =1980
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =390
                            Top =1980
                            Width =705
                            Height =270
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblDeclination"
                            Caption ="Zone"
                            FontName ="Calibri"
                            LayoutCachedLeft =390
                            LayoutCachedTop =1980
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =2250
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =1170
                    Top =480
                    Width =1395
                    FontSize =9
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboProtocol_Name"
                    ControlSource ="Protocol_Name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"prot"
                        "ocol\" ORDER BY Sort_Order; "
                    ColumnWidths ="2160;5040"
                    StatusBarText ="M. The name or code of the protocol governing the event (Protcl_Nam)"
                    FontName ="Calibri"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =480
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =480
                            Width =795
                            Height =240
                            FontSize =9
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label54"
                            Caption ="Protocol"
                            FontName ="Calibri"
                            LayoutCachedLeft =300
                            LayoutCachedTop =480
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =720
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =1185
                    Width =1395
                    FontSize =9
                    TabIndex =8
                    Name ="Tablet_Role"
                    ControlSource ="Entry_Role"
                    RowSourceType ="Value List"
                    RowSource ="PRIMARY;SECONDARY;SINGLE;OFFICE"
                    StatusBarText ="Data Entry Role of this Computer (Primary, Secondary, Single)"
                    FontName ="Calibri"
                    AllowValueListEdits =1

                    LayoutCachedLeft =1200
                    LayoutCachedTop =1185
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =1425
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =15
                            Top =1185
                            Width =1080
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label7"
                            Caption ="Entry_Role:"
                            FontName ="Calibri"
                            LayoutCachedLeft =15
                            LayoutCachedTop =1185
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1425
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3360
                    Top =1200
                    Width =960
                    Height =255
                    TabIndex =9
                    Name ="Text8"
                    ControlSource ="Timeframe"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1455
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2760
                            Top =1200
                            Width =540
                            Height =240
                            Name ="lblYear_Default"
                            Caption ="Year"
                            LayoutCachedLeft =2760
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3300
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1020
                    Top =3180
                    Width =3180
                    Height =255
                    TabIndex =10
                    Name ="tbxRoot"
                    ControlSource ="Root_Path"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =1020
                    LayoutCachedTop =3180
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =3435
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =240
                            Top =3180
                            Width =615
                            Height =240
                            FontWeight =700
                            Name ="lblRoot"
                            Caption ="Root"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =3180
                            LayoutCachedWidth =855
                            LayoutCachedHeight =3420
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =3120
                    Top =3540
                    Width =1980
                    Height =255
                    TabIndex =11
                    Name ="tbxData"
                    ControlSource ="Data_Path"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3120
                    LayoutCachedTop =3540
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =3795
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =240
                            Top =3540
                            Width =615
                            Height =240
                            FontWeight =700
                            Name ="lblData"
                            Caption ="Data"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =3540
                            LayoutCachedWidth =855
                            LayoutCachedHeight =3780
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3120
                    Top =3960
                    Width =1980
                    Height =255
                    TabIndex =12
                    Name ="tbxPhoto"
                    ControlSource ="Photo_Path"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    LayoutCachedLeft =3120
                    LayoutCachedTop =3960
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =4215
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =240
                            Top =3960
                            Width =615
                            Height =240
                            FontWeight =700
                            Name ="lblPhoto"
                            Caption ="Photo"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =240
                            LayoutCachedTop =3960
                            LayoutCachedWidth =855
                            LayoutCachedHeight =4200
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =3540
                    Width =1860
                    Height =300
                    Name ="lblDataRoot"
                    Caption ="Root"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Root path of the data directory"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =3540
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =3840
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =3960
                    Width =1860
                    Height =300
                    Name ="lblPhotoRoot"
                    Caption ="Root"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Root path of the photo directory"
                    LayoutCachedLeft =1020
                    LayoutCachedTop =3960
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =4260
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =4380
                    Top =3180
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =600
                    TabIndex =13
                    ForeColor =0
                    Name ="btnBrowseRoot"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Click to select the monitoring root directory"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4380
                    LayoutCachedTop =3180
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =3420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5220
                    Top =3540
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =600
                    TabIndex =14
                    ForeColor =0
                    Name ="btnBrowseData"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Click to select protocol data directory"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5220
                    LayoutCachedTop =3540
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =3780
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5220
                    Top =3960
                    Width =1260
                    Height =240
                    FontSize =9
                    FontWeight =600
                    TabIndex =15
                    ForeColor =0
                    Name ="btnBrowsePhoto"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Click to select the protocol photo directory"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5220
                    LayoutCachedTop =3960
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =4200
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
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
' MODULE:       frm_Set_Defaults
' Level:        Application module
' Version:      1.01
'
' Description:  application default related functions & procedures
' Data source:  tsys_App_Defaults
' Data access:  edit only, no deletions
'
' Source/date:  John R. Boetsch, May 16, 2006
' Adapted:  Bonnie Campbell, May, 2020
' Revisions:    JRB - 5/16/2006 - 1.00 - initial version
'               BLC - 5/5/2020  - 1.01 - add path defaults
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

'paths
Private m_RootPath As String
Private m_DataPath As String
Private m_PhotoPath As String

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
        'Me.lblTitle.Caption = m_Title
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
        'Me.lblDirections.Caption = m_Directions
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

'path info
Public Property Let RootPath(value As String)
    If Not IsNull(value) Then
        m_RootPath = value
    End If
    SetTempVar "RootPath", m_RootPath
End Property

Public Property Get RootPath() As String
    RootPath = m_RootPath
End Property

Public Property Let DataPath(value As String)
    If Not IsNull(value) Then
        m_DataPath = value
    End If
    SetTempVar "DataPath", Replace(m_DataPath, Me.RootPath, "")
    SetTempVar "FullDataPath", Nz(Me.RootPath, "") & m_DataPath
End Property

Public Property Get DataPath() As String
    DataPath = m_DataPath
End Property

Public Property Let PhotoPath(value As String)
    If Not IsNull(value) Then
        m_PhotoPath = value
    End If
    SetTempVar "PhotoPath", m_PhotoPath
    SetTempVar "FullPhotoPath", Nz(Me.RootPath, "") & m_PhotoPath
End Property

Public Property Get PhotoPath() As String
    PhotoPath = m_PhotoPath
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
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
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
'    tbxDevMode = DEV_MODE
                
    title = "Edit Date"
    'lblTitle.Caption = "" 'clear header title
    Directions = "Choose the desired date. " _
              & "Click Save to save the new date. " _
              & "Changing the date WILL change underlying data."
    
    'defaults
'    lblDirections.ForeColor = lngBlue
'    btnSave.HoverColor = lngGreen
'    btnCancel.HoverColor = lngRed
       
    'set values
'    EditTable = XML_Read("EditTable", Me.OpenArgs)
'    EditField = XML_Read("EditField", Me.OpenArgs)
'    EditIDField = XML_Read("EditIDField", Me.OpenArgs)
'    EditByID = XML_Read("UpdateByID", Me.OpenArgs)
'    EditID = XML_Read("EditID", Me.OpenArgs)
'    OriginalDate = XML_Read("ControlValue", Me.OpenArgs)
'Debug.Print Me.OpenArgs
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[" & Me.Name & " form])"
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
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[" & Me.Name & "])"
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
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[" & Me.Name & "])"
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
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[" & Me.Name & "])"
    End Select
    Resume Exit_Handler
End Sub



Private Sub cboUser_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    MsgBox "User not found.  To add this user, click the New user button.", vbOKOnly, "User Not In List"
    Me.ActiveControl.Undo
    Response = acDataErrContinue
    Me!cmdNewUser.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdNewUser_Click()
    On Error GoTo Err_Handler
    
    ' Open the contacts form
    DoCmd.OpenForm "frm_Contacts", , , , , , "new"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub


Private Sub cmdOK_Click()
    On Error GoTo Err_Handler

    Dim varOpenArgs As Variant
    
    varOpenArgs = Me.OpenArgs
    
    ' Make sure the information is valid before updating the record
    If varOpenArgs <> 0 Then
        '  Verify that the critical data elements have been completed before saving
        If IsNull(Me!User_name) Then
            MsgBox "Please indicate the user name", vbOKOnly, "Validation error"
            Me!cboUser.SetFocus
            GoTo Exit_Procedure
       ' ElseIf IsNull(Me!Park) Then
       '    MsgBox "Please indicate the park", vbOKOnly, "Validation error"
       '    Me!cboPark.SetFocus
       '    GoTo Exit_Procedure
        End If
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm "frm_Switchboard"
    Select Case varOpenArgs
        Case 1
            DoCmd.OpenForm "frm_Data_Gateway", , , , , , varOpenArgs
        Case 2
            DoCmd.OpenForm "frm_Browser", , , , , , varOpenArgs
        Case 3
            DoCmd.OpenForm "frm_QA_Tool", , , , , , varOpenArgs
        Case 4
            ' opened by switchboard only ... do nothing
        Case Else
            MsgBox "Error: OpenArgs property out of range", vbCritical
    End Select

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' ---------------------------------
' SUB:          btnBrowseRoot_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
' ---------------------------------
Private Sub btnBrowseRoot_Click()
On Error GoTo Err_Handler

    ' Exit if the user didn't specify a file
'    If IsNull(varFilePath) Then GoTo Exit_Handler
    Me.RootPath = SelectFolder("Choose the current monitoring directory root.")
    
    'populate paths
    Me.tbxRoot = Me.RootPath
    Me.lblDataRoot.Caption = Replace(Me.RootPath, "&", "&&") & "\"
    Me.lblPhotoRoot.Caption = Replace(Me.RootPath, "&", "&&") & "\"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowseRoot_Click[" & Me.Name & "])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnBrowseData_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
' ---------------------------------
Private Sub btnBrowseData_Click()
On Error GoTo Err_Handler

    ' Exit if the user didn't specify a file
'    If IsNull(varFilePath) Then GoTo Exit_Handler
    Me.DataPath = Replace(SelectFolder("Choose the current monitoring data directory root.", TempVars("RootPath")), TempVars("RootPath"), "")

    Me.tbxData = TempVars("DataPath")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowseData_Click[" & Me.Name & "])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnBrowsePhoto_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 5/5/2020 - initial version
' ---------------------------------
Private Sub btnBrowsePhoto_Click()
On Error GoTo Err_Handler

    ' Exit if the user didn't specify a file
'    If IsNull(varFilePath) Then GoTo Exit_Handler
    Me.PhotoPath = Nz(Replace(SelectFolder("Choose the current monitoring photo directory root.", Me.RootPath), Me.RootPath, ""), "")

    Me.tbxPhoto = Nz(Me.PhotoPath, "")

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowsePhoto_Click[" & Me.Name & "])"
    End Select
    Resume Exit_Handler
End Sub
