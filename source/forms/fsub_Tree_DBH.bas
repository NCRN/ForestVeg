Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2639
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =2520
    Top =8475
    Right =7065
    Bottom =11895
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcd25f3b3b063e440
    End
    RecordSource ="SELECT tbl_Tree_DBH.Tree_DBH_ID, tbl_Tree_DBH.Tree_Data_ID, tbl_Tree_DBH.DBH, tb"
        "l_Tree_DBH.Live FROM tbl_Tree_DBH;"
    Caption ="Stems"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =420
            BackColor =15527148
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =60
                    Width =900
                    Height =300
                    ColumnOrder =1
                    FontSize =12
                    FontWeight =700
                    Name ="tbxEquivDBH"
                    ControlSource ="=(((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2"
                    Format ="Fixed"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000070000000020000000000000005000000000000000300000001010000 ,
                        0xff000000ffffff00000000000600000004000000070000000101000022b14c00 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x31003000000000003100300000000000
                    End

                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x010002000000000000000500000001010000ff000000ffffff00020000003100 ,
                        0x3000000000000000000000000000000000000000000000000000000600000001 ,
                        0x01000022b14c00ffffff00020000003100300000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =60
                            Width =420
                            Height =300
                            FontSize =12
                            Name ="lblLD"
                            Caption ="L/D:"
                            FontName ="Calibri"
                            LayoutCachedTop =60
                            LayoutCachedWidth =420
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2100
                    Top =60
                    Width =336
                    Height =306
                    TabIndex =1
                    Name ="btnRefreshCalc"
                    Caption ="Command10"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddad000000000dadaadad00adada0adaddadad00adadadada ,
                        0xadadad00adadadaddadadad00adadadaadadad00adadadaddadad00adadadada ,
                        0xadad00adada0adaddad000000000dadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Refresh"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =2436
                    LayoutCachedHeight =366
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =720
                    Height =300
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BackColor =8421504
                    Name ="tbxSumDBH"
                    ControlSource ="=(((Sum(3.1415*((IIf([Live]=False,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2"
                    Format ="Fixed"
                    FontName ="Calibri"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                End
            End
        End
        Begin Section
            Height =420
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1620
                    Top =60
                    Width =720
                    Height =300
                    TabIndex =3
                    BackColor =15527148
                    Name ="tbxHighlightLive"
                    ConditionalFormat = Begin
                        0x01000000d8010000020000000100000000000000000000005c00000001000000 ,
                        0x00000000ffff990001000000000000005d000000bb0000000100000000000000 ,
                        0xffff990000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028004c0065006600740028005b0050006100720065006e007400 ,
                        0x5d002e0046006f0072006d002e0043006f006e00740072006f006c0073002800 ,
                        0x2200630062007800540072006500650053007400610074007500730022002900 ,
                        0x2e00560061006c00750065002c00340029003d00220044006500610064002200 ,
                        0x2c0049004900660028005b00630068006b004c006900760065005d003d005400 ,
                        0x7200750065002c0031002c00300029002c003000290000000000490049006600 ,
                        0x28004c0065006600740028005b0050006100720065006e0074005d002e004600 ,
                        0x6f0072006d002e0043006f006e00740072006f006c0073002800220063006200 ,
                        0x78005400720065006500530074006100740075007300220029002e0056006100 ,
                        0x6c00750065002c00350029003d00220041006c0069007600650022002c004900 ,
                        0x4900660028005b00630068006b004c006900760065005d003d00460061006c00 ,
                        0x730065002c0031002c00300029002c003000290000000000
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ffff99005b0000004900 ,
                        0x4900660028004c0065006600740028005b0050006100720065006e0074005d00 ,
                        0x2e0046006f0072006d002e0043006f006e00740072006f006c00730028002200 ,
                        0x6300620078005400720065006500530074006100740075007300220029002e00 ,
                        0x560061006c00750065002c00340029003d002200440065006100640022002c00 ,
                        0x49004900660028005b00630068006b004c006900760065005d003d0054007200 ,
                        0x750065002c0031002c00300029002c0030002900000000000000000000000000 ,
                        0x00000000000000000001000000000000000100000000000000ffff99005d0000 ,
                        0x0049004900660028004c0065006600740028005b0050006100720065006e0074 ,
                        0x005d002e0046006f0072006d002e0043006f006e00740072006f006c00730028 ,
                        0x0022006300620078005400720065006500530074006100740075007300220029 ,
                        0x002e00560061006c00750065002c00350029003d00220041006c006900760065 ,
                        0x0022002c0049004900660028005b00630068006b004c006900760065005d003d ,
                        0x00460061006c00730065002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =960
                    Top =60
                    Width =600
                    Height =300
                    ColumnWidth =900
                    FontSize =12
                    Name ="tbxDBH"
                    ControlSource ="DBH"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    OnLostFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000048010000020000000100000000000000000000006500000001000000 ,
                        0x00000000ffcccc00010000000000000066000000730000000100000000000000 ,
                        0xd6dfec0000000000000000000000000000000000000000000000000000000000 ,
                        0x49004900660028005b004c006900760065005d003d0054007200750065002c00 ,
                        0x49004900660028005b0074006200780045007100750069007600440042004800 ,
                        0x5d002e005b0046006f007200650043006f006c006f0072005d003d005b007600 ,
                        0x62005200650064005d002c0031002c00300029002c0049004900660028005b00 ,
                        0x740062007800530075006d004400420048005d002e005b0046006f0072006500 ,
                        0x43006f006c006f0072005d003d005b00760062005200650064005d002c003100 ,
                        0x2c00300029002900000000005b004c006900760065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =960
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ffcccc00640000004900 ,
                        0x4900660028005b004c006900760065005d003d0054007200750065002c004900 ,
                        0x4900660028005b00740062007800450071007500690076004400420048005d00 ,
                        0x2e005b0046006f007200650043006f006c006f0072005d003d005b0076006200 ,
                        0x5200650064005d002c0031002c00300029002c0049004900660028005b007400 ,
                        0x62007800530075006d004400420048005d002e005b0046006f00720065004300 ,
                        0x6f006c006f0072005d003d005b00760062005200650064005d002c0031002c00 ,
                        0x3000290029000000000000000000000000000000000000000000000100000000 ,
                        0x0000000100000000000000d6dfec000c0000005b004c006900760065005d003d ,
                        0x00460061006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =420
                            Top =60
                            Width =480
                            Height =300
                            FontSize =12
                            Name ="lblDBH"
                            Caption ="DBH"
                            FontName ="Calibri"
                            LayoutCachedLeft =420
                            LayoutCachedTop =60
                            LayoutCachedWidth =900
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =45
                    Width =351
                    Height =291
                    TabIndex =1
                    Name ="btnDeleteTreeDBH"
                    Caption ="Command6"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =60
                    LayoutCachedTop =45
                    LayoutCachedWidth =411
                    LayoutCachedHeight =336
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =4819194
                    HoverThemeColorIndex =5
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
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2100
                    Top =120
                    Width =245
                    TabIndex =2
                    Name ="chkLive"
                    ControlSource ="Live"
                    StatusBarText ="Indicates that the stem is alive"
                    DefaultValue ="True"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =120
                    LayoutCachedWidth =2345
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1620
                            Top =60
                            Width =420
                            Height =299
                            FontSize =12
                            Name ="lblLive"
                            Caption ="Live"
                            FontName ="Calibri"
                            LayoutCachedLeft =1620
                            LayoutCachedTop =60
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =359
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
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
' FORM:         fsub_Tree_DBH
' Level:        Application report
' Version:      1.04
'
' Description:  Form related functions & procedures for application
' Requires:     Keypad Utils module
'
' Source/date:  Bonnie Campbell, April 3, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/3/2018 - 1.01 - added documentation, error handling
'               BLC   - 4/19/2018 - 1.02 - validate DBH
'               BLC - 4/21/2018   - 1.03 - added record count check, ctbxDBH lost focus event
'                                          code cleanup
'               BLC - 4/30/2018   - 1.04 - remove DBH validation (shift to fsub Exit event)
' =================================

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tree_DBH", "Tree_DBH_ID") = dbText Then
            Me!Tree_DBH_ID = fxnGUIDGen
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_AfterUpdate
' Description:  form after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler

    Me.Refresh

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterUpdate[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDBH_Click
' Description:  DBH textbox click actions
' Requires:     Keypad Utils module
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub tbxDBH_Click()
On Error GoTo Err_Handler

    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'set keypad form to launch & control on this form to be updated by it
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "tbxDBH"
    
    'launch keypad
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_Click[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnKeypadDBH_Click
' Description:  DBH keypad button click actions
' Requires:     Keypad Utils module
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub btnKeypadDBH_Click()
On Error GoTo Err_Handler

    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'set keypad form to launch & control on this form to be updated by it
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "tbxDBH"
    
    'launch keypad
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnKeypadDBH_Click[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDeleteTreeDBH_Click
' Description:  delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'   BLC - 4/19/2018 - validate DBH
'   BLC - 4/30/2018 - remove DBH validation
' ---------------------------------
Private Sub btnDeleteTreeDBH_Click()
On Error GoTo Err_Handler

    'If MsgBox("You are about to DELETE all data for this tree for this sampling event only." & vbNewLine & vbNewLine & "Is this OK?", vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Procedure
    With CodeContextObject
        On Error Resume Next
        DoCmd.GoToControl Screen.PreviousControl.Name
        Err.Clear
        If (Not .Form.NewRecord) Then
            DoCmd.RunCommand acCmdDeleteRecord
        End If
        If (.Form.NewRecord And Not .Form.Dirty) Then
            Beep
        End If
        If (.Form.NewRecord And .Form.Dirty) Then
            DoCmd.RunCommand acCmdUndo
        End If
    End With

    'check DBH
    'ValidDBH "Tree"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDeleteTreeDBH_Click[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDBH_Change
' Description:  DBH textbox change actions
' Requires:     -
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 19, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/19/2018 - initial version
'   BLC - 4/21/2018 - added record count check
'   BLC - 4/30/2018 - remove DBH validation
' ---------------------------------
Private Sub tbxDBH_Change()
On Error GoTo Err_Handler
    
    'If Me.Recordset.RecordCount > 0 Then _
    '    ValidDBH "Tree"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_Change[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDBH_LostFocus
' Description:  DBH textbox LostFocus actions
' Requires:     -
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/21/2018 - initial version
'   BLC - 4/30/2018 - remove DBH validation
' ---------------------------------
Private Sub tbxDBH_LostFocus()
On Error GoTo Err_Handler
    
'    If Me.Recordset.RecordCount > 0 Then _
'        ValidDBH "Tree"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_LostFocus[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxDBH_AfterUpdate
' Description:  DBH textbox AfterUpdate actions
' Requires:     -
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/21/2018 - initial version
'   BLC - 4/30/2018 - remove DBH validation
' ---------------------------------
Private Sub tbxDBH_AfterUpdate()
On Error GoTo Err_Handler
    
'    If Me.Recordset.RecordCount > 0 Then _
'        ValidDBH "Tree"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_AfterUpdate[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnRefreshCalc_Click
' Description:  refresh calculation button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub btnRefreshCalc_Click()
On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdRefresh
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRefreshCalc_Click[fsub_Tree_DBH])"
    End Select
    Resume Exit_Handler
End Sub
