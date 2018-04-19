Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9780
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =4500
    Top =3855
    Right =14535
    Bottom =10260
    DatasheetGridlinesColor =12632256
    Filter ="[Transect_Azimuth] = \"360\" "
    RecSrcDt = Begin
        0xe883dba6cff2e240
    End
    RecordSource ="tbl_CWD_Data"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
    ShowPageMargins =0
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
            Height =1380
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =840
                    Width =7800
                    Height =483
                    FontSize =10
                    Name ="txtComments"
                    ControlSource ="CWD_Notes"
                    FontName ="Calibri"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =840
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =1323
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =420
                            Top =840
                            Width =1140
                            Height =300
                            FontSize =12
                            Name ="lblComments"
                            Caption ="Comments:"
                            FontName ="Calibri"
                            LayoutCachedLeft =420
                            LayoutCachedTop =840
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =4
                    BorderWidth =2
                    OverlapFlags =85
                    BorderLineStyle =3
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =25
                    ListWidth =5760
                    Left =6900
                    Top =60
                    Width =300
                    Height =300
                    FontSize =12
                    TabIndex =5
                    BackColor =-2147483643
                    BorderColor =5026082
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cboQuickFind"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.TSN, tlu_Plants.Genu"
                        "s, tlu_Plants.Species, tlu_Plants.Family FROM tlu_Plants WHERE (((tlu_Plants.Woo"
                        "dy)=True) AND ((tlu_Plants.Favorite)=True)) ORDER BY tlu_Plants.Latin_Name;"
                    ColumnWidths ="2520;3240;0;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Quick Find"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =8580
                    Top =120
                    Width =840
                    Height =255
                    FontSize =12
                    TabIndex =6
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtTSN"
                    ControlSource ="TSN"
                    FontName ="Calibri"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =120
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8100
                            Top =120
                            Width =420
                            Height =255
                            FontSize =12
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblTSN"
                            Caption ="TSN"
                            FontName ="Calibri"
                            LayoutCachedLeft =8100
                            LayoutCachedTop =120
                            LayoutCachedWidth =8520
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =420
                    Width =900
                    Height =330
                    FontSize =12
                    TabIndex =2
                    Name ="txtDiameter"
                    ControlSource ="Diameter"
                    StatusBarText ="The diameter of the debris at the intersection of the transect."
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b007400780074004400690061006d006500 ,
                        0x7400650072005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =1620
                    LayoutCachedTop =420
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =750
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b007400780074004400690061006d0065007400 ,
                        0x650072005d0029003d0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =420
                            Width =1500
                            Height =330
                            FontSize =12
                            Name ="lblDiameter"
                            Caption ="Diameter (cm)"
                            FontName ="Calibri"
                            LayoutCachedLeft =60
                            LayoutCachedTop =420
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =750
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6960
                    Top =480
                    Height =300
                    TabIndex =4
                    Name ="chkHollow"
                    ControlSource ="Hollow"
                    StatusBarText ="Considered hollow if cavity extends 0.5m along the central longitudinal axis of "
                        "the piece and the cavity entrance is at least 1/4 the diameter of the piece."

                    LayoutCachedLeft =6960
                    LayoutCachedTop =480
                    LayoutCachedWidth =7220
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =7205
                            Top =420
                            Width =720
                            Height =360
                            FontSize =12
                            Name ="lblHollow"
                            Caption ="Hollow"
                            FontName ="Calibri"
                            LayoutCachedLeft =7205
                            LayoutCachedTop =420
                            LayoutCachedWidth =7925
                            LayoutCachedHeight =780
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =1080
                    Width =360
                    ColumnWidth =1755
                    TabIndex =7
                    Name ="txtTransect_Azimuth"
                    ControlSource ="Transect_Azimuth"
                    DefaultValue ="360"
                    FontName ="Calibri"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1080
                    LayoutCachedWidth =420
                    LayoutCachedHeight =1320
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7320
                    Top =60
                    Width =306
                    Height =306
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    Name ="cmdAdd_To_Quickfind"
                    Caption ="i"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Show Taxon Details"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7320
                    LayoutCachedTop =60
                    LayoutCachedWidth =7626
                    LayoutCachedHeight =366
                    Alignment =4
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3780
                    Top =420
                    Width =780
                    Height =330
                    FontSize =12
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0044006500630061007900 ,
                        0x5f0043006c006100730073005d0029003d00540072007500650000000000
                    End
                    Name ="cboDecay_Class"
                    ControlSource ="Decay_Class"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Decay Class\")) ORDER BY tlu_"
                        "Enumerations.Sort_Order; "
                    ValidationText =">=1"
                    FontName ="Calibri"

                    LayoutCachedLeft =3780
                    LayoutCachedTop =420
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =750
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                        0x73004e0075006c006c0028005b00630062006f00440065006300610079005f00 ,
                        0x43006c006100730073005d0029003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =9000
                    Top =60
                    Width =456
                    Height =366
                    TabIndex =9
                    ForeColor =0
                    Name ="cmdDeleteRec"
                    Caption ="Command17"
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9000
                    LayoutCachedTop =60
                    LayoutCachedWidth =9456
                    LayoutCachedHeight =426
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
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =30
                    ListWidth =3600
                    Left =1620
                    Top =60
                    Width =5220
                    Height =300
                    FontSize =12
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cboTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN, tlu_Plants.Latin_Name FROM tlu_Plants WHERE (((tlu_Plants"
                        ".Tree)=True)) OR (((tlu_Plants.Vine)=True)) OR (((tlu_Plants.Shrub)=True)) ORDER"
                        " BY tlu_Plants.Latin_Name;"
                    ColumnWidths ="0;3600"
                    FontName ="Calibri"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =960
                            Top =60
                            Width =600
                            Height =290
                            FontSize =12
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblcbo_TSN"
                            Caption ="Taxon"
                            FontName ="Calibri"
                            LayoutCachedLeft =960
                            LayoutCachedTop =60
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =350
                        End
                    End
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =2940
                    Top =420
                    Width =786
                    Height =366
                    FontSize =12
                    TabIndex =10
                    ForeColor =6108695
                    Name ="cmdOpen_Form_Decay_Class"
                    Caption ="Decay Class"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Open Form"
                    BackStyle =0

                    LayoutCachedLeft =2940
                    LayoutCachedTop =420
                    LayoutCachedWidth =3726
                    LayoutCachedHeight =786
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =7200
                    Left =5280
                    Top =420
                    Width =1560
                    Height =330
                    FontSize =12
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cboTag_ID"
                    ControlSource ="Tag_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, tbl_Tags.Tag_Status, IIf(IsNull([azimuth])"
                        ",\"\",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, tbl_Tags.Microplot_"
                        "Number AS MP, tbl_Tags.Location_ID, Forms!frm_Events!Location_ID AS Loc_field, F"
                        "orms!frm_Events!txtLocation_ID AS Loc_ctrl FROM tbl_Tags WHERE (((tbl_Tags.Locat"
                        "ion_ID)=Forms!frm_Events!Location_ID)) ORDER BY tbl_Tags.Tag_Status DESC , tbl_T"
                        "ags.Tag; "
                    ColumnWidths ="0;1440;2160;2160;1440"
                    ValidationText =">=1"
                    FontName ="Calibri"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =420
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4740
                            Top =420
                            Width =495
                            Height =360
                            FontSize =12
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblTag_ID"
                            Caption ="Tag"
                            FontName ="Calibri"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =420
                            LayoutCachedWidth =5235
                            LayoutCachedHeight =780
                        End
                    End
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

Private Sub CboQuickFind_AfterUpdate()
    Me!cboTSN.Value = Me!CboQuickFind.Column(2)
    Me!CboQuickFind = ""
    Me!txtDiameter.SetFocus
End Sub

Private Sub cboTSN_GotFocus()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim strSpeciesType As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Species"
  strControlToUpdate = "cboTSN"
  'Choose TREE, SAPLING, SEEDLING, CWD, VINE or TARGETED HERB
  strSpeciesType = "CWD"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenSpeciespad(strKeypadFormName, frmFormToUpdate, strControlToUpdate, strSpeciesType)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub cmdAdd_To_Quickfind_Click()
On Error GoTo Err_cmdAdd_To_Quickfind_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Plants"
    stLinkCriteria = "[TSN]=" & Me!txtTSN
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    'Form_frm_Field_Data_Foliage_Problems.Data_ID.DefaultValue = StringFromGUID(Me!Data_ID)
    
Exit_cmdAdd_To_Quickfind_Click:
    Exit Sub

Err_cmdAdd_To_Quickfind_Click:
    MsgBox Err.Description
    Resume Exit_cmdAdd_To_Quickfind_Click
    
End Sub

Private Sub cmdCWD_Diameter_Keypad_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "txtDiameter"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub cmdDeleteRec_Click()
On Error GoTo Err_Handler
    
    If MsgBox("You are about to DELETE one CWD record." & vbNewLine & vbNewLine & "Is this OK?", vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Procedure
    
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

Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler

    Select Case Transect_Azimuth
        Case 360
            Forms!frm_Events!chkTransectChecked_360 = True
            Forms!frm_Events!lblTransectChecked_360.Requery
        Case 120
            Forms!frm_Events!chkTransectChecked_120 = True
            Forms!frm_Events!lblTransectChecked_120.Requery
        Case 240
            Forms!frm_Events!chkTransectChecked_240 = True
            Forms!frm_Events!lblTransectChecked_240.Requery
    End Select
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_CWD_Data", "CWD_Data_ID") = dbText Then
            Me!CWD_Data_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Load()
Dim strTransect As String
    strTransect = "360"
    Me.txtTransect_Azimuth.DefaultValue = strTransect
    Forms![frm_Events]![fsub_Transects].Form.Filter = "[Transect_Azimuth] = """ & strTransect & """ "
    Forms![frm_Events]![fsub_Transects].Form.FilterOn = True
End Sub
Private Sub cmdOpen_Form_Decay_Class_Click()
On Error GoTo Err_cmdOpen_Form_Decay_Class_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Decay_Classes"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Form_Decay_Class_Click:
    Exit Sub
Err_cmdOpen_Form_Decay_Class_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Form_Decay_Class_Click
End Sub

Private Sub txtDiameter_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "txtDiameter"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub
