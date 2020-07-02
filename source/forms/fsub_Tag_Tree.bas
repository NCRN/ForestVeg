Version =21
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
    Width =13500
    DatasheetFontHeight =9
    ItemSuffix =33
    Left =210
    Top =3390
    Right =10455
    Bottom =5700
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
            Height =1260
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5280
                    Top =540
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

                    LayoutCachedLeft =5280
                    LayoutCachedTop =540
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =855
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000fff200000d0000005b00 ,
                        0x630068006b005200460053005d003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
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
                    ListWidth =3600
                    Left =2880
                    Top =60
                    Width =3840
                    Height =360
                    FontSize =12
                    FontWeight =700
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cbxTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN, tlu_Plants.Favorite, tlu_Plants.Latin_Name, tlu_Plants.Tr"
                        "ee FROM tlu_Plants WHERE (((tlu_Plants.Tree)=True) AND ((tlu_Plants.Accepted_Fou"
                        "nd)=False)) ORDER BY tlu_Plants.Favorite, tlu_Plants.Latin_Name;"
                    ColumnWidths ="0;0;3600"
                    StatusBarText ="TSN of Specimen"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2040
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
                            LayoutCachedLeft =2040
                            LayoutCachedTop =60
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =480
                    Width =840
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
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0041007a0069006d007500740068005d00 ,
                        0x290000000000
                    End

                    LayoutCachedLeft =8340
                    LayoutCachedTop =480
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =840
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500110000004900 ,
                        0x73004e0075006c006c0028005b0041007a0069006d007500740068005d002900 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7260
                            Top =480
                            Width =1020
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblAzimuh"
                            Caption ="Azimuth:"
                            LayoutCachedLeft =7260
                            LayoutCachedTop =480
                            LayoutCachedWidth =8280
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10320
                    Top =480
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
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00440069007300740061006e0063006500 ,
                        0x5d00290000000000
                    End

                    LayoutCachedLeft =10320
                    LayoutCachedTop =480
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =840
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500120000004900 ,
                        0x73004e0075006c006c0028005b00440069007300740061006e00630065005d00 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9240
                            Top =480
                            Width =1020
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblDistance"
                            Caption ="Distance:"
                            LayoutCachedLeft =9240
                            LayoutCachedTop =480
                            LayoutCachedWidth =10260
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12360
                    Top =480
                    Width =840
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

                    LayoutCachedLeft =12360
                    LayoutCachedTop =480
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11280
                            Top =480
                            Width =1020
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblMicroplot_Number"
                            Caption ="Microplot:"
                            LayoutCachedLeft =11280
                            LayoutCachedTop =480
                            LayoutCachedWidth =12300
                            LayoutCachedHeight =840
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
                            Left =7080
                            Top =60
                            Width =1200
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblStartDate"
                            Caption ="Start_Date:"
                            LayoutCachedLeft =7080
                            LayoutCachedTop =60
                            LayoutCachedWidth =8280
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11640
                    Top =60
                    Width =1560
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

                    LayoutCachedLeft =11640
                    LayoutCachedTop =60
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10320
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
                            LayoutCachedLeft =10320
                            LayoutCachedTop =60
                            LayoutCachedWidth =11580
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
                    Height =480
                    FontSize =16
                    FontWeight =700
                    TabIndex =7
                    Name ="tbxTag"
                    ControlSource ="Tag"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =120
                    Top =540
                    Width =1455
                    Height =270
                    FontSize =9
                    TabIndex =8
                    ForeColor =0
                    Name ="btnReplaceTag"
                    Caption ="Replace Tag"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Replace a lost tag with a new one"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =1575
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
                    Top =480
                    Width =2340
                    Height =374
                    FontSize =12
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
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
                    LayoutCachedTop =480
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =854
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1680
                            Top =480
                            Width =1140
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTagStatus"
                            Caption ="Tag Status:"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =480
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6300
                    Top =540
                    Width =660
                    Height =300
                    TabIndex =9
                    Name ="tbxRFS"
                    ControlSource ="RFS"
                    Format ="True/False"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =540
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =840
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =5400
                    Top =900
                    Width =2115
                    Height =285
                    FontSize =10
                    FontWeight =500
                    ForeColor =1643706
                    Name ="lblRFS"
                    Caption ="-- Removed from Study --"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =900
                    LayoutCachedWidth =7515
                    LayoutCachedHeight =1185
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5820
                    Top =600
                    Width =240
                    TabIndex =10
                    BorderColor =10921638
                    Name ="chkRFS"
                    ControlSource ="RFS"
                    GridlineColor =10921638

                    LayoutCachedLeft =5820
                    LayoutCachedTop =600
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5340
                            Top =540
                            Width =405
                            Height =315
                            Name ="lblChkRFS"
                            Caption ="RFS"
                            ControlTipText ="Removed from Study (RFS)"
                            LayoutCachedLeft =5340
                            LayoutCachedTop =540
                            LayoutCachedWidth =5745
                            LayoutCachedHeight =855
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
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
' MODULE:       fsub_Tag_Tree
' Level:        Form module
' Version:      1.05
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/5/2018 - 1.01 - added documentation, error handling
'               BLC   - 4/16/2018 - 1.02 - renamed cboTag_Status > cbxTagStatus,
'                                          txtMicroplot_Number > tbxMicroplotNumber,
'                                          txtDistance > tbxDistance
'               BLC - 5/20/2019   - 1.03 - updated fsub_Tree_Data.tbxHabit based on species
'               BLC - 5/23/2019   - 1.04 - add detail color based on if RFS set
'               BLC - 9/15/2019   - 1.05 - updated to display RFS label when tag is removed from study
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

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
'   BLC - 9/15/2019 - display lblRFS if tbxRFS = True
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
    
    Me.Detail.backcolor = IIf(tbxRFS, lngLtRose, lngWhite)
    
    'default
    Me.tbxRFS.visible = False
    Me.lblRFS.visible = False
    
    'display RFS only if it *is* removed from study (tbxRFS = true)
    If Me.tbxRFS = True Then Me.lblRFS.visible = True
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  BeforeUpdate Events
' ----------------
' ---------------------------------
' SUB:          cbxTagStatus_BeforeUpdate
' Description:  combobox before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
'               BLC   - 4/16/2018 - renamed cboTag_Status > cbxTagStatus
' ---------------------------------
Private Sub cbxTagStatus_BeforeUpdate(Cancel As Integer)
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
            "Error encountered (#" & Err.Number & " - cbxTagStatus_BeforeUpdate[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTSN_BeforeUpdate
' Description:  combobox before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub cbxTSN_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

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
            "Error encountered (#" & Err.Number & " - cbxTSN_BeforeUpdate[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click Events
' ----------------
' ---------------------------------
' SUB:          btnReplaceTag_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnReplaceTag_Click()
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
            "Error encountered (#" & Err.Number & " - btnReplaceTag_Click[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxAzimuth_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub tbxAzimuth_Click()
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
            "Error encountered (#" & Err.Number & " - tbxAzimuth_Click[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDistance_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
'               BLC   - 4/16/2018 - renamed txtDistance > tbxDistance
' ---------------------------------
Private Sub tbxDistance_Click()
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
            "Error encountered (#" & Err.Number & " - tbxDistance_Click[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxMicroplotNumber_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
'               BLC   - 4/16/2018 - renamed txtMicroplot_Number > tbxMicroplotNumber
' ---------------------------------
Private Sub tbxMicroplotNumber_Click()
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
            "Error encountered (#" & Err.Number & " - tbxMicroplotNumber_Click[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTSN_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' Requires:     Keypad_Utils module
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub cbxTSN_Enter()
On Error GoTo Err_Handler

    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    Dim strSpeciesType As String
        
    'set keypad to open, control to be updated
    strKeypadFormName = "frm_Pad_Species"
    strControlToUpdate = "cbxTSN"
    strSpeciesType = "Tree"
    
    'open keypad
    Set frmFormToUpdate = Me
    Call OpenSpeciespad(strKeypadFormName, frmFormToUpdate, strControlToUpdate, strSpeciesType)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTSN_Enter[fsub_Tag_Tree])"
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
   
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTSN_AfterUpdate[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Methods
' ---------------------------------

' ---------------------------------
' SUB:          SaveRecord
' Description:  save record actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - SaveRecord[fsub_Tag_Tree])"
    End Select
    Resume Exit_Handler
End Sub
