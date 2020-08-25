Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =2460
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =2535
    Top =5865
    Right =5325
    Bottom =8055
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x37d06983d518e540
    End
    RecordSource ="SELECT tbl_Sapling_DBH.Sapling_DBH_ID, tbl_Sapling_DBH.Sapling_Data_ID, tbl_Sapl"
        "ing_DBH.DBH, tbl_Sapling_DBH.Live FROM tbl_Sapling_DBH;"
    Caption ="Stems"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
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
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =60
                    Width =900
                    Height =299
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    Name ="tbxEquivDBH"
                    ControlSource ="=(((Sum(3.1415*([DBH]/2)^2))*(1/3.1415))^0.5)*2"
                    Format ="Fixed"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000078000000030000000000000006000000000000000300000001010000 ,
                        0xff000000ffffff00000000000000000004000000060000000101000022b14c00 ,
                        0xffffff000000000005000000090000000b00000001010000ff000000ffffff00 ,
                        0x310030000000000031000000310030000000310000000000
                    End

                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =359
                    ConditionalFormat14 = Begin
                        0x010003000000000000000600000001010000ff000000ffffff00020000003100 ,
                        0x3000000000000000000000000000000000000000000000000000000000000001 ,
                        0x01000022b14c00ffffff00010000003100020000003100300000000000000000 ,
                        0x00000000000000000000000000000500000001010000ff000000ffffff000100 ,
                        0x00003100000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =60
                            Width =420
                            Height =299
                            FontSize =12
                            Name ="lblLD"
                            Caption ="L/D:"
                            FontName ="Calibri"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =480
                            LayoutCachedHeight =359
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2099
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

                    LayoutCachedLeft =2099
                    LayoutCachedTop =60
                    LayoutCachedWidth =2435
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
                    ColumnOrder =1
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
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =420
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1680
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
                        0x22006300620078005300610070006c0069006e00670053007400610074007500 ,
                        0x7300220029002e00560061006c00750065002c00340029003d00220044006500 ,
                        0x6100640022002c0049004900660028005b004c006900760065005d003d005400 ,
                        0x7200750065002c0031002c00300029002c003000290000000000490049006600 ,
                        0x28004c0065006600740028005b0050006100720065006e0074005d002e004600 ,
                        0x6f0072006d002e0043006f006e00740072006f006c0073002800220063006200 ,
                        0x78005300610070006c0069006e00670053007400610074007500730022002900 ,
                        0x2e00560061006c00750065002c00350029003d00220041006c00690076006500 ,
                        0x22002c0049004900660028005b004c006900760065005d003d00460061006c00 ,
                        0x730065002c0031002c00300029002c003000290000000000
                    End

                    LayoutCachedLeft =1680
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ffff99005b0000004900 ,
                        0x4900660028004c0065006600740028005b0050006100720065006e0074005d00 ,
                        0x2e0046006f0072006d002e0043006f006e00740072006f006c00730028002200 ,
                        0x6300620078005300610070006c0069006e006700530074006100740075007300 ,
                        0x220029002e00560061006c00750065002c00340029003d002200440065006100 ,
                        0x640022002c0049004900660028005b004c006900760065005d003d0054007200 ,
                        0x750065002c0031002c00300029002c0030002900000000000000000000000000 ,
                        0x00000000000000000001000000000000000100000000000000ffff99005d0000 ,
                        0x0049004900660028004c0065006600740028005b0050006100720065006e0074 ,
                        0x005d002e0046006f0072006d002e0043006f006e00740072006f006c00730028 ,
                        0x0022006300620078005300610070006c0069006e006700530074006100740075 ,
                        0x007300220029002e00560061006c00750065002c00350029003d00220041006c ,
                        0x0069007600650022002c0049004900660028005b004c006900760065005d003d ,
                        0x00460061006c00730065002c0031002c00300029002c00300029000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =600
                    Height =299
                    ColumnWidth =900
                    FontSize =12
                    Name ="tbxDBH"
                    ControlSource ="DBH"
                    FontName ="Calibri"
                    OnLostFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000050010000030000000100000000000000000000006500000001000000 ,
                        0x00000000ffcccc000000000006000000660000006900000001010000ff000000 ,
                        0xffffff0001000000000000006a000000770000000100000000000000d6dfec00 ,
                        0x49004900660028005b004c006900760065005d003d0054007200750065002c00 ,
                        0x49004900660028005b0074006200780045007100750069007600440042004800 ,
                        0x5d002e005b0046006f007200650043006f006c006f0072005d003d005b007600 ,
                        0x62005200650064005d002c0031002c00300029002c0049004900660028005b00 ,
                        0x740062007800530075006d004400420048005d002e005b0046006f0072006500 ,
                        0x43006f006c006f0072005d003d005b00760062005200650064005d002c003100 ,
                        0x2c003000290029000000000031003000000000005b004c006900760065005d00 ,
                        0x3d00460061006c007300650000000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =359
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffcccc00640000004900 ,
                        0x4900660028005b004c006900760065005d003d0054007200750065002c004900 ,
                        0x4900660028005b00740062007800450071007500690076004400420048005d00 ,
                        0x2e005b0046006f007200650043006f006c006f0072005d003d005b0076006200 ,
                        0x5200650064005d002c0031002c00300029002c0049004900660028005b007400 ,
                        0x62007800530075006d004400420048005d002e005b0046006f00720065004300 ,
                        0x6f006c006f0072005d003d005b00760062005200650064005d002c0031002c00 ,
                        0x3000290029000000000000000000000000000000000000000000000000000006 ,
                        0x00000001010000ff000000ffffff000200000031003000000000000000000000 ,
                        0x00000000000000000000000001000000000000000100000000000000d6dfec00 ,
                        0x0c0000005b004c006900760065005d003d00460061006c007300650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =479
                            Top =60
                            Width =480
                            Height =299
                            FontSize =12
                            Name ="lblDBH"
                            Caption ="DBH"
                            FontName ="Calibri"
                            LayoutCachedLeft =479
                            LayoutCachedTop =60
                            LayoutCachedWidth =959
                            LayoutCachedHeight =359
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
                    Name ="btnDeleteDBH"
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

                    LayoutCachedLeft =60
                    LayoutCachedTop =45
                    LayoutCachedWidth =411
                    LayoutCachedHeight =336
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2160
                    Top =120
                    Width =245
                    TabIndex =2
                    Name ="Live"
                    ControlSource ="Live"
                    StatusBarText ="Indicates that the stem is alive"
                    DefaultValue ="True"

                    LayoutCachedLeft =2160
                    LayoutCachedTop =120
                    LayoutCachedWidth =2405
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1680
                            Top =60
                            Width =420
                            Height =299
                            FontSize =12
                            Name ="lblLive"
                            Caption ="Live"
                            FontName ="Calibri"
                            LayoutCachedLeft =1680
                            LayoutCachedTop =60
                            LayoutCachedWidth =2100
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
' FORM:         fsub_Sapling_DBH
' Level:        Application report
' Version:      1.04
'
' Description:  Form related functions & procedures for application
' Requires:     Keypad Utils module
'
' Source/date:  Bonnie Campbell, April 19, 2018
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 4/19/2018 - 1.01 - added documentation, error handling
'                                          field renaming cmd>btn, Label>lbl, txt>tbx
'                                          cmd_DBH_Keypad_Click() removed
'               BLC   - 4/21/2018 - 1.02 - added tbxDBH lost focus event
'               BLC   - 7/31/2020 - 1.03 - issue: cannot enter > 1 stems
'                                          fix: commented out tbxDBH_LostFocus code
'               BLC   - 8/7/2020  - 1.04 - adjusted ValidDBH to include event date parameter
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
' Adapted:      Bonnie Campbell, April 19, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/19/2018 - added error handling, documentation
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Sapling_DBH", "Sapling_DBH_ID") = dbText Then
            Me!Sapling_DBH_ID = fxnGUIDGen
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Sapling_DBH])"
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
' Adapted:      Bonnie Campbell, April 19, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/19/2018 - added error handling, documentation
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
            "Error encountered (#" & Err.Number & " - Form_AfterUpdate[fsub_Sapling_DBH])"
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
' Adapted:      Bonnie Campbell, April 19, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/19/2018 - added error handling, documentation
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
            "Error encountered (#" & Err.Number & " - tbxDBH_Click[fsub_Sapling_DBH])"
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
' Adapted:      Bonnie Campbell, April 19, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/19/2018 - added error handling, documentation
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
            "Error encountered (#" & Err.Number & " - btnKeypadDBH_Click[fsub_Sapling_DBH])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDeleteDBH_Click
' Description:  delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 19, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/19/2018 - added error handling, documentation
'                     validate DBH
'   BLC - 8/7/2020  - adjusted ValidDBH to include event date parameter
' ---------------------------------
Private Sub btnDeleteDBH_Click()
On Error GoTo Err_Handler

    'If MsgBox("You are about to DELETE all data for this sapling for this sampling event only." & vbNewLine & vbNewLine & "Is this OK?", vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Procedure
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
    ValidDBH "Sapling", Me.Parent.tbxEventDate

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDeleteDBH_Click[fsub_Sapling_DBH])"
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
'   BLC - 8/7/2020  - adjusted ValidDBH to include event date parameter
' ---------------------------------
Private Sub tbxDBH_Change()
On Error GoTo Err_Handler
    
    If Me.Recordset.RecordCount > 0 Then _
        ValidDBH "Sapling", Me.Parent.tbxEventDate
        
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_Change[fsub_Sapling_DBH])"
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
'   BLC - 7/31/2020 - issue: prevents users from entering > 1 stems
'                     fix: commented out code
' ---------------------------------
Private Sub tbxDBH_LostFocus()
On Error GoTo Err_Handler
    
'    If Me.Recordset.RecordCount > 0 Then _
'        ValidDBH "Sapling"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_LostFocus[fsub_Sapling_DBH])"
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
' Adapted:      Bonnie Campbell, April 19, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/19/2018 - added error handling, documentation
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
            "Error encountered (#" & Err.Number & " - btnRefreshCalc_Click[fsub_Sapling_DBH])"
    End Select
    Resume Exit_Handler
End Sub
