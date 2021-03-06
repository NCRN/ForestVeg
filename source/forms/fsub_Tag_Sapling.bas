﻿Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13650
    DatasheetFontHeight =9
    ItemSuffix =37
    Left =150
    Top =135
    Right =11685
    Bottom =2535
    DatasheetGridlinesColor =15062992
    OrderBy ="[tbl_Tags].[Tag]"
    RecSrcDt = Begin
        0xbb20843c6eaee340
    End
    RecordSource ="tbl_Tags"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin FormHeader
            Height =0
            BackColor =16768194
            Name ="FormHeader"
        End
        Begin Section
            Height =853
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5340
                    Top =480
                    Width =900
                    Height =315
                    TabIndex =11
                    Name ="tbxRFSHighlight"
                    ConditionalFormat = Begin
                        0x010000007e000000010000000100000000000000000000000e00000001000000 ,
                        0x00000000fff20000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b005200460053005d003d00540072007500650000000000
                    End

                    LayoutCachedLeft =5340
                    LayoutCachedTop =480
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =795
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000fff200000d0000005b00 ,
                        0x630068006b005200460053005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =479
                    Width =839
                    Height =360
                    FontSize =12
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxAzimuth"
                    ControlSource ="Azimuth"
                    StatusBarText ="Azimuth from plot center to specimen (true north)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =479
                    LayoutCachedWidth =9179
                    LayoutCachedHeight =839
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7260
                            Top =479
                            Width =1019
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblAzimuth"
                            Caption ="Azimuth:"
                            LayoutCachedLeft =7260
                            LayoutCachedTop =479
                            LayoutCachedWidth =8279
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10319
                    Top =479
                    Width =900
                    Height =360
                    FontSize =12
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxDistance"
                    ControlSource ="Distance"
                    StatusBarText ="Distance (m) from plot center to near EDGE of tree"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10319
                    LayoutCachedTop =479
                    LayoutCachedWidth =11219
                    LayoutCachedHeight =839
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9240
                            Top =479
                            Width =1019
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblDistance"
                            Caption ="Distance:"
                            LayoutCachedLeft =9240
                            LayoutCachedTop =479
                            LayoutCachedWidth =10259
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12359
                    Top =479
                    Width =839
                    Height =360
                    FontSize =12
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxMicroplotNumber"
                    ControlSource ="Microplot_Number"
                    StatusBarText ="The Microplot location of specimen"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b004d006900630072006f0070006c006f00 ,
                        0x74005f004e0075006d006200650072005d00290000000000
                    End

                    LayoutCachedLeft =12359
                    LayoutCachedTop =479
                    LayoutCachedWidth =13198
                    LayoutCachedHeight =839
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b004d006900630072006f0070006c006f007400 ,
                        0x5f004e0075006d006200650072005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11280
                            Top =479
                            Width =1019
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblMicroplotNumber"
                            Caption ="Microplot:"
                            LayoutCachedLeft =11280
                            LayoutCachedTop =479
                            LayoutCachedWidth =12299
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =60
                    Width =1620
                    Height =360
                    FontSize =12
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxStartDate"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking began on this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7079
                            Top =60
                            Width =1199
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblStartDate"
                            Caption ="Start_Date:"
                            LayoutCachedLeft =7079
                            LayoutCachedTop =60
                            LayoutCachedWidth =8278
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11639
                    Top =60
                    Width =1559
                    Height =360
                    FontSize =12
                    TabIndex =6
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxStopDate"
                    ControlSource ="Stop_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking ended for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =11639
                    LayoutCachedTop =60
                    LayoutCachedWidth =13198
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10319
                            Top =60
                            Width =1260
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblStopDate"
                            Caption ="Stop_Date:"
                            LayoutCachedLeft =10319
                            LayoutCachedTop =60
                            LayoutCachedWidth =11579
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin ComboBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =5040
                    Left =2880
                    Top =60
                    Width =3839
                    Height =360
                    FontSize =12
                    FontWeight =700
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cbxTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN, tlu_Plants.Favorite, IIf([Latin_Name]=\"Kalmia latifolia\""
                        ",[Latin_Name] & \"***\",IIf([Latin_Name]=\"Lindera benzoin\",[Latin_Name] & \"**"
                        "*\",IIf([Latin_Name]=\"Ilex verticillata\",[Latin_Name] & \"***\",[Latin_Name]))"
                        ") AS Name, IIf([Tree]=True,\"Tree\",\"Shrub\") AS Habit FROM tlu_Plants WHERE (("
                        "(tlu_Plants.Tree)=True) AND ((tlu_Plants.Accepted_Found)=False)) OR (((tlu_Plant"
                        "s.Accepted_Found)=False) AND ((tlu_Plants.Shrub)=True)) ORDER BY tlu_Plants.Favo"
                        "rite, IIf([Latin_Name]=\"Kalmia latifolia\",[Latin_Name] & \"***\",IIf([Latin_Na"
                        "me]=\"Lindera benzoin\",[Latin_Name] & \"***\",IIf([Latin_Name]=\"Ilex verticill"
                        "ata\",[Latin_Name] & \"***\",[Latin_Name])));"
                    ColumnWidths ="0;0;3600;1440"
                    StatusBarText ="TSN of Specimen"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =6719
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2039
                            Top =60
                            Width =780
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblSpecies"
                            Caption ="Taxon:"
                            LayoutCachedLeft =2039
                            LayoutCachedTop =60
                            LayoutCachedWidth =2819
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1545
                    Height =479
                    FontSize =16
                    FontWeight =700
                    TabIndex =7
                    Name ="tbxTag"
                    ControlSource ="Tag"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =539
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =119
                    Top =540
                    Width =1455
                    Height =270
                    FontSize =9
                    TabIndex =8
                    ForeColor =0
                    Name ="btnReplaceTag"
                    Caption ="Replace Tag"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =119
                    LayoutCachedTop =540
                    LayoutCachedWidth =1574
                    LayoutCachedHeight =810
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =479
                    Width =2340
                    Height =374
                    FontSize =12
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x010000007e010000010000000100000000000000000000008e00000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4c00650066007400280046006f0072006d00730021005b006600730075006200 ,
                        0x5f005300610070006c0069006e0067005f0044006100740061005d0021005b00 ,
                        0x6300620078005300610070006c0069006e006700530074006100740075007300 ,
                        0x5d002c00340029003d002200270044006500610064002200200041006e006400 ,
                        0x200028005b006300620078005400610067005300740061007400750073005d00 ,
                        0x3c003e00220052006500740069007200650064002000280049006e0020004f00 ,
                        0x660066006900630065002900220020004f00720020005b006300620078005400 ,
                        0x610067005300740061007400750073005d003c003e00270049006e0061006300 ,
                        0x7400690076006500200028004c006f007300740029002700290000000000
                    End
                    Name ="cbxTagStatus"
                    ControlSource ="Tag_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Tag Status\")) ORDER BY tlu_Enumerations.Sort_Order;"
                    StatusBarText ="Last sampled as tree or sapling?"
                    BeforeUpdate ="[Event Procedure]"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2880
                    LayoutCachedTop =479
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =853
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ffff99008d0000004c00 ,
                        0x650066007400280046006f0072006d00730021005b0066007300750062005f00 ,
                        0x5300610070006c0069006e0067005f0044006100740061005d0021005b006300 ,
                        0x620078005300610070006c0069006e0067005300740061007400750073005d00 ,
                        0x2c00340029003d002200270044006500610064002200200041006e0064002000 ,
                        0x28005b006300620078005400610067005300740061007400750073005d003c00 ,
                        0x3e00220052006500740069007200650064002000280049006e0020004f006600 ,
                        0x66006900630065002900220020004f00720020005b0063006200780054006100 ,
                        0x67005300740061007400750073005d003c003e00270049006e00610063007400 ,
                        0x690076006500200028004c006f00730074002900270029000000000000000000 ,
                        0x00000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1679
                            Top =479
                            Width =1140
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTagStatus"
                            Caption ="Tag Status:"
                            LayoutCachedLeft =1679
                            LayoutCachedTop =479
                            LayoutCachedWidth =2819
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6300
                    Top =480
                    Width =660
                    Height =300
                    TabIndex =9
                    Name ="tbxRFS"
                    ControlSource ="RFS"
                    Format ="True/False"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =480
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =780
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5940
                    Top =540
                    Width =240
                    TabIndex =10
                    BorderColor =10921638
                    Name ="chkRFS"
                    ControlSource ="RFS"
                    GridlineColor =10921638

                    LayoutCachedLeft =5940
                    LayoutCachedTop =540
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5460
                            Top =480
                            Width =405
                            Height =315
                            Name ="lblChkRFS"
                            Caption ="RFS"
                            ControlTipText ="Removed from Study (RFS)"
                            LayoutCachedLeft =5460
                            LayoutCachedTop =480
                            LayoutCachedWidth =5865
                            LayoutCachedHeight =795
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
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
' MODULE:       fsub_Tag_Sapling
' Level:        Application module
' Version:      1.05
'
' Description:  Sapling tag related functions & procedures for version control
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 3/26/2018 - 1.01 - added documentation, error handling
'               BLC   - 4/9/2018  - 1.02 - renamed cbo's > cbx, txt > tbx
'               BLC   - 11/5/2018 - 1.03 - fix cbxTag_Status reference to cbxTagStatus,
'                                          set cbxTagStatus.Locked = No vs. Yes (tag properties)
'               BLC - 5/20/2019   - 1.04 - updated fsub_Tree_Data.tbxHabit based on species
'               BLC - 5/23/2019   - 1.05 - add detail color based on if RFS set
' =================================

' ---------------------------------
'  Properties
' ---------------------------------
Public SaplingHabit As String

' ---------------------------------
'  Events
' ---------------------------------
' ---------------------------------
' SUB:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 20, 2019
' Adapted:      -
' Revisions:
'   BLC - 5/20/2019 - initial version
'   BLC - 5/23/2019 - update detail background based on RFS value
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
    
    If IsLoaded("fsub_Tag_Sapling") Then Me.Parent.Form!cbxHabit = Nz(Me.SaplingHabit, "")

    Me.Detail.BackColor = IIf(tbxRFS, lngLtRose, lngWhite)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  BeforeUpdate Events
' ----------------
' ---------------------------------
' SUB:          cbxTagStatus_BeforeUpdate
' Description:  Tag status actions for before record update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
'   BLC - 4/9/2018 - rename cboTag_Status > cbxTagStatus
'   BLC - 11/5/2018 - fix cbxTag_Status reference to cbxTagStatus, set cbxTagStatus.Locked = No vs. Yes (tag properties)
' ---------------------------------
Public Sub cbxTagStatus_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!cbxTagStatus
    ChangeDescription = "Please confirm the revised TAG STATUS below"
    ChangeFieldType = "Combo_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenConfirmValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, _
        "tbl_Tags", "Tag_Status", "Tag_ID", _
        Me!Tag_ID, Nz(Me!cbxTagStatus.OldValue, "Null"), _
        Me!cbxTagStatus, "", "", ""
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTagStatus_BeforeUpdate[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTSN_BeforeUpdate
' Description:  TSN actions before record update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
'   BLC - 4/9/2018 - renamed cboTSN > cbxTSN
' ---------------------------------
Public Sub cbxTSN_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    'Dim frm As Form
    'Dim ctl As Control
    
    'Set frm = Me
    'Set ctl = Me!cboTSN
    
    'OpenChangeHistory frm, ctl, "tbl_Tags", "TSN", "Tag_ID", Me!Tag_ID, Me!cboTSN.OldValue, Me!cboTSN, "tlu_Plants", "Latin_Name", "TSN", dbLong
    
    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!cbxTSN
    ChangeDescription = "Please confirm the revised SPECIES ID below"
    ChangeFieldType = "Combo_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenConfirmValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, _
        "tbl_Tags", "TSN", "Tag_ID", Me!Tag_ID, _
        Nz(Me!cbxTSN.OldValue, "Null"), _
        Me!cbxTSN, "tlu_Plants", "Latin_Name", "TSN"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTSN_BeforeUpdate[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click Events
' ----------------
' ---------------------------------
' SUB:          btnReplaceTag_Click
' Description:  Replace tag button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
'   BLC - 4/9/2018 - renamed cmdReplace_Tag > btnReplaceTag
' ---------------------------------
Public Sub btnReplaceTag_Click()
On Error GoTo Err_Handler

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!tbxTag
    ChangeDescription = "Please enter the new TAG NUMBER below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, _
        "tbl_Tags", "Tag", "Tag_ID", Me!Tag_ID, Me!Tag.Value
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReplaceTag_Click[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxAzimuth_Click
' Description:  Azimuth textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
'   BLC - 4/9/2018 - rename txtAzimuth > tbxAzimuth
' ---------------------------------
Public Sub tbxAzimuth_Click()
On Error GoTo Err_Handler

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!tbxAzimuth
    ChangeDescription = "Please enter the revised AZIMUTH below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, _
        "tbl_Tags", "Azimuth", "Tag_ID", Me!Tag_ID, _
        Nz(Me!Azimuth.Value, "Null")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxAzimuth_Click[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDistance_Click()
' Description:  Distance textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
'   BLC - 4/9/2018 - renamed txtDistance > tbxDistance
' ---------------------------------
Public Sub tbxDistance_Click()
On Error GoTo Err_Handler

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!tbxDistance
    ChangeDescription = "Please enter the revised DISTANCE below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, _
        "tbl_Tags", "Distance", "Tag_ID", Me!Tag_ID, _
        Nz(Me!Distance.Value, "Null")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDistance_Click[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxMicroplotNumber_Click
' Description:  Microplot number textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
'   BLC - 4/9/2018 - renamed txtMicroplot_Number > tbxMicroplotNumber
' ---------------------------------
Public Sub tbxMicroplotNumber_Click()
On Error GoTo Err_Handler

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!tbxMicroplotNumber
    ChangeDescription = "Please enter the revised MICROPLOT NUMBER below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, _
        "tbl_Tags", "Microplot_Number", "Tag_ID", Me!Tag_ID, _
        Nz(Me!Microplot_Number.Value, "Null")
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxMicroplotNumber_Click[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  AfterUpdate Events
' ----------------
' ---------------------------------
' SUB:          cbxTSN_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 20, 2019
' Adapted:      -
' Revisions:
'   BLC - 5/20/2019 - initial version
' ---------------------------------
Private Sub cbxTSN_AfterUpdate()
On Error GoTo Err_Handler
   
    'update the habit field
    Me.SaplingHabit = cbxTSN.Column(3)
    'MsgBox cbxTSN & " habit is " & cbxTSN.Column(3) & " " & Me.SaplingHabit
    Me.Parent.Form!cbxHabit = Me.SaplingHabit

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTSN_AfterUpdate[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Methods
' ---------------------------------

' ---------------------------------
' SUB:          SaveRecord
' Description:  Save record actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, March 26, 2018
' Adapted:      -
' Revisions:
'   BLC - 3/26/2018 - initial version (w/ documentation)
' ---------------------------------
Public Sub SaveRecord()
On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SaveRecord[fsub_Tag_Sapling])"
    End Select
    Resume Exit_Handler
End Sub
