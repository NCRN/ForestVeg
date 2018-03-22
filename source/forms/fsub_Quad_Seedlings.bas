Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7080
    DatasheetFontHeight =10
    ItemSuffix =28
    Left =5490
    Top =4650
    Right =12525
    Bottom =7785
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x43120bc17fa7e340
    End
    RecordSource ="SELECT tbl_Quadrat_Seedlings_Data.*, tlu_Plants.Latin_Name FROM tbl_Quadrat_Seed"
        "lings_Data LEFT JOIN tlu_Plants ON tbl_Quadrat_Seedlings_Data.TSN=tlu_Plants.TSN"
        "; "
    Caption ="sfrm_Quad_Seedlings"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
            Height =240
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =900
                    Width =1740
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Taxon"
                    FontName ="Calibri"
                    LayoutCachedLeft =900
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =85
                    Left =4140
                    Width =1020
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label24"
                    Caption ="Height (cm)"
                    FontName ="Calibri"
                    LayoutCachedLeft =4140
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =85
                    Left =5220
                    Width =1740
                    Height =240
                    FontSize =10
                    FontWeight =700
                    Name ="Label25"
                    Caption ="Browsable/Browsed"
                    FontName ="Calibri"
                    LayoutCachedLeft =5220
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =240
                End
            End
        End
        Begin Section
            Height =420
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =30
                    ListWidth =3600
                    Left =480
                    Top =60
                    Width =3006
                    Height =306
                    FontSize =11
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cboTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants.TSN FROM tlu_P"
                        "lants WHERE (((tlu_Plants.Woody)=True)) ORDER BY tlu_Plants.Latin_Name; "
                    ColumnWidths ="0;3600"
                    OnEnter ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Full Plant List"

                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =3486
                    LayoutCachedHeight =366
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =600
                    Height =300
                    ColumnWidth =1185
                    FontSize =11
                    Name ="txtHeight"
                    ControlSource ="Height"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740048006500690067006800 ,
                        0x74005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500180000004900 ,
                        0x73004e0075006c006c0028005b00740078007400480065006900670068007400 ,
                        0x5d0029003d005400720075006500000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =4
                    BorderWidth =2
                    OverlapFlags =119
                    BorderLineStyle =3
                    IMESentenceMode =3
                    ColumnCount =6
                    ListRows =40
                    ListWidth =6480
                    Left =3480
                    Top =60
                    Width =300
                    Height =299
                    ColumnWidth =1320
                    FontSize =11
                    TabIndex =1
                    BorderColor =5026082
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cboQuickFind"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.Latin_Name, tlu_Plants.Genus, tlu_Plants.Species, tlu_Plants.T"
                        "SN, tlu_Plants.Family, tlu_Plants.Common FROM tlu_Plants WHERE (((tlu_Plants.Woo"
                        "dy)=True) AND ((tlu_Plants.Favorite)=True)) ORDER BY tlu_Plants.Latin_Name; "
                    ColumnWidths ="2520;0;0;0;0;3960"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    FontName ="Calibri"
                    OnGotFocus ="[Event Procedure]"
                    ControlTipText ="Quick Find Species List"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =60
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =359
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3840
                    Top =60
                    Width =306
                    Height =306
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdAdd_To_Quickfind"
                    Caption ="i"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Show Taxon Details"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4146
                    LayoutCachedHeight =366
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
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =366
                    Height =306
                    TabIndex =3
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
                    LayoutCachedTop =60
                    LayoutCachedWidth =426
                    LayoutCachedHeight =366
                    ForeThemeColorIndex =0
                    UseTheme =1
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
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =6000
                    Top =60
                    Width =120
                    Height =240
                    FontSize =11
                    FontWeight =700
                    Name ="Label21"
                    Caption ="/"
                    FontName ="Calibri"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =60
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5400
                    Top =60
                    Width =540
                    Height =300
                    FontSize =11
                    TabIndex =5
                    Name ="txtBrowsable"
                    ControlSource ="Browsable"
                    StatusBarText ="This seedling was browseable (leaves below 2 meters)"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000009a000000010000000100000000000000000000001c00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00740078007400420072006f0077007300 ,
                        0x610062006c0065005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001b0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400420072006f00770073006100 ,
                        0x62006c0065005d0029003d005400720075006500000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6180
                    Top =60
                    Width =540
                    Height =300
                    FontSize =11
                    TabIndex =6
                    Name ="txtBrowsed"
                    ControlSource ="Browsed"
                    StatusBarText ="Deer browse was noticable on this seedling"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000096000000010000000100000000000000000000001a00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00740078007400420072006f0077007300 ,
                        0x650064005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =6180
                    LayoutCachedTop =60
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500190000004900 ,
                        0x73004e0075006c006c0028005b00740078007400420072006f00770073006500 ,
                        0x64005d0029003d00540072007500650000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =4320
                    Left =6780
                    Top =60
                    Width =240
                    Height =300
                    FontSize =12
                    FontWeight =500
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboBrowsePick"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Browse_Status\")) ORDER BY tl"
                        "u_Enumerations.[Sort_Order];"
                    ColumnWidths ="1080;3240"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =360
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

Private Sub cboBrowsePick_AfterUpdate()
On Error GoTo Err_Handler
    Select Case Me!cboBrowsePick.Column(0)
        Case "Yes / Yes"
            Me!txtBrowsable.Value = "Yes"
            Me!txtBrowsed.Value = "Yes"
        Case "Yes / No"
            Me!txtBrowsable.Value = "Yes"
            Me!txtBrowsed.Value = "No"
        Case "No / No"
            Me!txtBrowsable.Value = "No"
            Me!txtBrowsed.Value = "No"
    End Select
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Error$
    Resume Exit_Procedure
End Sub

Private Sub CboQuickFind_AfterUpdate()
    Me!cboTSN.Value = Me!CboQuickFind.Column(3)
    Me!CboQuickFind = ""
    Me!txtHeight.SetFocus
End Sub

Private Sub cboQuickFind_Enter()
    Me!CboQuickFind.Requery
End Sub

Private Sub cboQuickFind_GotFocus()
    Me!CboQuickFind.Requery
End Sub

Private Sub cboTSN_Enter()
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
  strSpeciesType = "SEEDLING"
  
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
    stLinkCriteria = "[TSN]=" & Me!cboTSN
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_cmdAdd_To_Quickfind_Click:
    Exit Sub

Err_cmdAdd_To_Quickfind_Click:
    MsgBox Err.Description
    Resume Exit_cmdAdd_To_Quickfind_Click
    
End Sub

Private Sub cmdDeleteRec_Click()
On Error GoTo Err_Handler

    If MsgBox("You are about to DELETE all data for this seedling for this sampling event only." & vbNewLine & vbNewLine & "Is this OK?", vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Procedure
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

Private Sub cmdHerb_Seedlings_Keypad_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "txtHeight"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Quadrat_Seedlings_Data", "Quadrat_Seedlings_ID") = dbText Then
            Me!Quadrat_Seedlings_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub txtHeight_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "txtHeight"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub txtLatin_Name_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim strSpeciesType As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Species"
  strControlToUpdate = "TSN"
  'Choose TREE, SAPLING, SEEDLING, CWD, VINE or TARGETED HERB
  strSpeciesType = "VINE"
  
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenSpeciespad(strKeypadFormName, frmFormToUpdate, strControlToUpdate, strSpeciesType)

Exit_cmdOpenKeyPad_Click:
  Exit Sub
Err_cmdOpenKeyPad_Click:
  MsgBox Err.Description
  Resume Exit_cmdOpenKeyPad_Click
End Sub
