Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =127
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13140
    DatasheetFontHeight =10
    ItemSuffix =45
    Left =7830
    Top =3135
    Right =20880
    Bottom =10590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf8c4ff537de0e240
    End
    Caption ="Species Pad"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1320
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =9544
                    Height =480
                    ColumnOrder =0
                    FontSize =18
                    FontWeight =700
                    Name ="txtLatin"

                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =9604
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11460
                    Top =60
                    Width =1260
                    Height =474
                    FontSize =14
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClear"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11460
                    LayoutCachedTop =60
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =534
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =10856415
                    PressedColor =413911
                    PressedThemeColorIndex =5
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
                    OverlapFlags =85
                    Left =9780
                    Top =60
                    Width =1620
                    Height =1199
                    TabIndex =2
                    ForeColor =0
                    Name ="cmdAssign"
                    Caption ="Assign"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100040000000000000200000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777778 ,
                        0x7777777777777777777777777777777880777777777777777777777777777778 ,
                        0x8007777777777777777777777777777880007777777777777777777777777778 ,
                        0x8000077777777777777788888888888880000077777777777777880000000000 ,
                        0x0000000777777777777788000000000000000000777777777777880000000000 ,
                        0x0000000007777777777788000000000000000000007777777777880000000000 ,
                        0x0000000000077777777788000000000000000000007077777777880000000000 ,
                        0x0000000007077777777788000000000000000000707777777777880000000000 ,
                        0x0000000707777777777788077777777770000070777777777777770000000000 ,
                        0x0000070777777777777777777777777880007077777777777777777777777778 ,
                        0x8007077777777777777777777777777880707777777777777777777777777778 ,
                        0x7007777777777777777777777777777770777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =1259
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
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
                    OverlapFlags =85
                    Left =11460
                    Top =600
                    Width =1260
                    Height =654
                    FontSize =14
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11460
                    LayoutCachedTop =600
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =1254
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =10856415
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
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8040
                    Top =60
                    Width =1564
                    Height =300
                    ColumnOrder =1
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="txtValue"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =60
                    LayoutCachedWidth =9604
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7140
                            Top =60
                            Width =844
                            Height =300
                            FontSize =12
                            FontWeight =700
                            Name ="lblValue"
                            Caption ="TSN:"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =60
                            LayoutCachedWidth =7984
                            LayoutCachedHeight =360
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =6900
                    Height =305
                    ColumnOrder =2
                    FontSize =11
                    FontWeight =700
                    TabIndex =5
                    Name ="txtPlantType"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =365
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =960
                    Width =9544
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="txtCommon"

                    LayoutCachedLeft =60
                    LayoutCachedTop =960
                    LayoutCachedWidth =9604
                    LayoutCachedHeight =1260
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6060
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =6135
                    Height =5640
                    BorderColor =10921638
                    Name ="sfrm_Pad_Species_Favorites"
                    SourceObject ="Form.frm_Pad_Species_Favorites"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =5760
                End
                Begin Subform
                    OverlapFlags =85
                    Left =6600
                    Top =120
                    Width =6135
                    Height =5639
                    TabIndex =1
                    BorderColor =10921638
                    Name ="sfrm_Pad_Species_Common"
                    SourceObject ="Form.frm_Pad_Species_Common"
                    GridlineColor =10921638

                    LayoutCachedLeft =6600
                    LayoutCachedTop =120
                    LayoutCachedWidth =12735
                    LayoutCachedHeight =5759
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

Private Sub Form_Load()
    txtPlantType.Value = Me.OpenArgs & " Species Lookup"
    txtValue = ""
    txtLatin = ""
    txtCommon = ""
    cmdCancel.SetFocus
End Sub

Private Sub Form_Open(Cancel As Integer)
    txtPlantType.Value = Me.OpenArgs & " Species Lookup"
    txtValue = ""
    txtLatin = ""
    txtCommon = ""
    cmdCancel.SetFocus
End Sub

Private Sub cmdAssign_Click()
On Error GoTo Err_Handler
    If Not IsNull(txtValue) Then
        CtrlToUpdate = txtValue
        'The following line is optional, but may be needed if you have calculated field on the subform
        'This line is now remarked because it caused issues with forms that were not ready to be saved, eg, some required fields were not entered yet
        'CtrlToUpdate.Parent.Refresh
    End If
    'OpenConfirmValueAndLog "Please confirm the revised SPECIES ID below", "Combo_Box", Forms.fsub_Tag_Tree, CtrlToUpdate, "tbl_Tags", "TSN", "Tag_ID", Me!Tag_ID, Nz(Me!cboTSN.OldValue, "Null"), Me!cboTSN, "tlu_Plants", "Latin_Name", "TSN"
    DoCmd.Close acForm, "frm_Pad_Species"
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
    DoCmd.Close acForm, "frm_Pad_Species"
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
    txtValue = ""
    txtLatin = ""
    txtCommon = ""
End Sub
