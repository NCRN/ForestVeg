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
    Width =3240
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =1140
    Top =1980
    Right =6165
    Bottom =5145
    DatasheetGridlinesColor =12632256
    AfterDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x9a1141aeeb00e340
    End
    RecordSource ="tbl_Tree_Foliage_Conditions"
    Caption ="Stems"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
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
            Height =240
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =720
                    Width =975
                    Height =240
                    FontSize =10
                    Name ="Label7"
                    Caption ="Condition"
                    FontName ="Calibri"
                    LayoutCachedLeft =720
                    LayoutCachedWidth =1695
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =85
                    Left =2100
                    Width =960
                    Height =240
                    FontSize =10
                    Name ="Label8"
                    Caption ="% Afflicted"
                    FontName ="Calibri"
                    LayoutCachedLeft =2100
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =240
                End
            End
        End
        Begin Section
            Height =360
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1620
                    Top =60
                    Width =900
                    Height =255
                    ColumnWidth =2310
                    Name ="DBH_ID"
                    ControlSource ="Tree_Foliage_Condition_ID"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =960
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Tree_Data_ID"
                    ControlSource ="Tree_Data_ID"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =315
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2016
                    Left =420
                    Top =60
                    Width =1680
                    Height =300
                    ColumnWidth =1215
                    FontSize =11
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="DBH"
                    ControlSource ="Condition"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description, tlu_Enumer"
                        "ations.Enum_Group FROM tlu_Enumerations WHERE (((tlu_Enumerations.Enum_Group)=\""
                        "Foliage Condition\")) ORDER BY tlu_Enumerations.Sort_Order; "
                    ColumnWidths ="288;1728"
                    FontName ="Calibri"

                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =900
                    Height =300
                    ColumnWidth =1800
                    FontSize =11
                    TabIndex =3
                    Name ="Percent_Afflicted"
                    ControlSource ="Percent_Afflicted"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Top =60
                    Width =351
                    Height =291
                    TabIndex =4
                    Name ="cmd_Tree_Condition_Delete"
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
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedTop =60
                    LayoutCachedWidth =351
                    LayoutCachedHeight =351
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

Private Sub cmd_Tree_Condition_Delete_Click()
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

Private Sub Form_AfterDelConfirm(Status As Integer)
    Me.Parent.Refresh
End Sub

Private Sub Form_AfterUpdate()
    Forms![frm_Events]![fsub_Tree_Data]![chkFoliage_Conditions_Checked].Value = True
    Me.Parent.Refresh
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tree_Foliage_Condtions", "Tree_Foliage_Condition_ID") = dbText Then
            Me!Tree_Foliage_Condition_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Percent_Afflicted_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "Percent_Afflicted"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub
