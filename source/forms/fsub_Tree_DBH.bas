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
    ItemSuffix =16
    Left =750
    Top =1620
    Right =5805
    Bottom =5220
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
                    Name ="Text12"
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
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000bc000000020000000100000000000000000000001f00000001000000 ,
                        0x00000000ffcccc000100000000000000200000002d0000000100000000000000 ,
                        0xd6dfec0000000000000000000000000000000000000000000000000000000000 ,
                        0x41006200730028005b007400620078004400420048005d002d005b0074006200 ,
                        0x78005000720069006f0072004400420048005d0029003e003d00340000000000 ,
                        0x5b004c006900760065005d003d00460061006c007300650000000000
                    End

                    LayoutCachedLeft =960
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ffcccc001e0000004100 ,
                        0x6200730028005b007400620078004400420048005d002d005b00740062007800 ,
                        0x5000720069006f0072004400420048005d0029003e003d003400000000000000 ,
                        0x00000000000000000000000000000001000000000000000100000000000000d6 ,
                        0xdfec000c0000005b004c006900760065005d003d00460061006c007300650000 ,
                        0x0000000000000000000000000000000000000000
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
                    OverlapFlags =93
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
                            OverlapFlags =93
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
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1740
                    Top =60
                    Width =780
                    TabIndex =3
                    Name ="tbxPriorDBH"
                    ControlSource ="=IIf(IsNull([Tree_Data_ID]),0,Nz(GetPriorDBH([Tree_Data_ID],\"Tree\"),999))"

                    LayoutCachedLeft =1740
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =300
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
' Version:      1.01
'
' Description:  Form related functions & procedures for application
' Requires:     Keypad Utils module
'
' Source/date:  Bonnie Campbell, April 3, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/3/2018 - 1.01 - added documentation, error handling
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
