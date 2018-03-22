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
    Width =2459
    DatasheetFontHeight =10
    ItemSuffix =11
    Left =2820
    Top =6525
    Right =5355
    Bottom =8460
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb554b7c8b163e440
    End
    RecordSource ="SELECT tbl_Sapling_DBH.Sapling_DBH_ID, tbl_Sapling_DBH.Sapling_Data_ID, tbl_Sapl"
        "ing_DBH.DBH, tbl_Sapling_DBH.Live FROM tbl_Sapling_DBH;"
    Caption ="Stems"
    BeforeUpdate ="[Event Procedure]"
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
                    OverlapFlags =223
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =60
                    Width =900
                    Height =299
                    FontSize =12
                    FontWeight =700
                    Name ="txtEquivDBH"
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
                            Name ="Label8"
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
                    Name ="cmdRefresh_Calculation"
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
                    Left =1020
                    Top =60
                    Width =600
                    Height =299
                    ColumnWidth =900
                    FontSize =12
                    Name ="txtDBH"
                    ControlSource ="DBH"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000084000000020000000000000006000000000000000300000001010000 ,
                        0xff000000ffffff00010000000000000004000000110000000100000000000000 ,
                        0xd6dfec0000000000000000000000000000000000000000000000000000000000 ,
                        0x31003000000000005b004c006900760065005d003d00460061006c0073006500 ,
                        0x00000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =359
                    ConditionalFormat14 = Begin
                        0x010002000000000000000600000001010000ff000000ffffff00020000003100 ,
                        0x3000000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000d6dfec000c0000005b004c006900760065005d003d00460061 ,
                        0x006c0073006500000000000000000000000000000000000000000000
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
                            Name ="DBH_Label"
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
                    Name ="cmd_Tree_DBH_delete"
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
                    OverlapFlags =85
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
                            OverlapFlags =85
                            Left =1680
                            Top =60
                            Width =420
                            Height =299
                            FontSize =12
                            Name ="Label11"
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

Private Sub cmd_DBH_Keypad_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
    'This routine requires the presence of the Keypad_Utils module.
    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'The two lines below should be changed to reflect the name of the keypad to open
    '    and the name of the control to be updated.
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "txtDBH"
    'The lines below should not usually be edited.
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
    Exit Sub
Err_cmdOpenKeyPad_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub cmd_Tree_DBH_delete_Click()
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

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Error$
    Resume Exit_Procedure
End Sub

Private Sub cmdRefresh_Calculation_Click()
On Error GoTo Err_cmdRefresh_Calculation_Click

    DoCmd.RunCommand acCmdRefresh

Exit_cmdRefresh_Calculation_Click:
    Exit Sub
Err_cmdRefresh_Calculation_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Calculation_Click
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Sapling_DBH", "Sapling_DBH_ID") = dbText Then
            Me!Sapling_DBH_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub txtDBH_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
    'This routine requires the presence of the Keypad_Utils module.
    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'The two lines below should be changed to reflect the name of the keypad to open
    '    and the name of the control to be updated.
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "txtDBH"
    'The lines below should not usually be edited.
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
    Exit Sub
Err_cmdOpenKeyPad_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenKeyPad_Click
End Sub
