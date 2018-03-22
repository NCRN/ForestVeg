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
    ItemSuffix =6
    Left =7455
    Top =5445
    Right =10980
    Bottom =7830
    DatasheetGridlinesColor =12632256
    AfterDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0xe8be61080c64e440
    End
    RecordSource ="tbl_Sapling_Conditions"
    Caption ="Conditions and Pests"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1680
                    Width =1260
                    Height =255
                    ColumnWidth =2310
                    Name ="DBH_ID"
                    ControlSource ="Sapling_Condition_ID"

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =1680
                    Width =660
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Sapling_Data_ID"
                    ControlSource ="Sapling_Data_ID"

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =255
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =24
                    ListWidth =4320
                    Left =420
                    Top =60
                    Width =2820
                    Height =300
                    ColumnWidth =2865
                    FontSize =12
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboCondition"
                    ControlSource ="Condition"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Tree_Condition.Description, tlu_Tree_Condition.Category FROM tlu_Tree"
                        "_Condition ORDER BY tlu_Tree_Condition.Sequence;"
                    ColumnWidths ="2880;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =360
                    DatasheetCaption ="Conditions and Pests"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Top =60
                    Width =351
                    Height =291
                    TabIndex =3
                    Name ="cmd_Sapling_Condition_Delete"
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

Private Sub cboCondition_AfterUpdate()
'Validation checks to ensure that specific pests are only associated with the specific target species.

Dim intTSN As Long
Dim strTaxa As String

intTSN = Forms!frm_Events!fsub_Sapling_Data!fsub_Tag_Sapling!cboTSN.Value

'MsgBox intTSN

Select Case Me!cboCondition

Case "Beech bark disease" 'Fagus grandifolia
    If intTSN <> "19462" Then
        MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End If

Case "Butternut canker"
    
    Select Case intTSN
        Case 19250 'Juglans
            GoTo Exitsub
        Case 19254 'Juglans
            GoTo Exitsub
        Case 501306 'Carya
            GoTo Exitsub
        Case 19227 'Carya
            GoTo Exitsub
        Case 19231 'Carya
            GoTo Exitsub
        Case 19234 'Carya
            GoTo Exitsub
        Case 19235 'Carya
            GoTo Exitsub
        Case 19241 'Carya
            GoTo Exitsub
        Case 19243 'Carya
            GoTo Exitsub
        Case Else
        
        MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
        
    End Select
    
Case "Chestnut blight"
    
    Select Case intTSN
        Case 505160 'Castanea
            GoTo Exitsub
        Case 19454 'Castanea
            GoTo Exitsub
        Case 501318 'Castanea
            GoTo Exitsub
        Case 19457 'Castanea
            GoTo Exitsub
        Case Else
            MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End Select
    
Case "Dogwood anthracnose"
    
    strTaxa = DLookup("[Genus]", "tlu_Plants", "[TSN_Accepted] = " & intTSN)
    
    If strTaxa = "Cornus" Then ' Dogwood genus
        Exit Sub
    Else
        MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End If

Case "Emerald ash borer" 'Oleaceae Family
    strTaxa = DLookup("[Family]", "tlu_Plants", "[TSN_Accepted] = " & intTSN)
    
    If strTaxa = "Oleaceae" Then
        Exit Sub
    Else
        MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End If
    
Case "Hemlock scale"

    If intTSN <> 183397 Then 'Tsuga Canadensis
         MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End If

Case "Oak wilt" 'Oak family
    
    strTaxa = DLookup("[Family]", "tlu_Plants", "[TSN_Accepted] = " & intTSN)
    
    If strTaxa <> "Fagaceae" Then
        MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End If
    
    
Case "Thousand cankers disease"
    Select Case intTSN
        Case 19250 'Juglans
            GoTo Exitsub
        Case 19254 'Juglans
            GoTo Exitsub
        Case Else
             MsgBox "Please Check", vbCritical, "NCRN Vegetation Monitoring"
    End Select
        
Case Else
    
End Select
Exitsub:
    Exit Sub
    
End Sub


Private Sub cmd_Sapling_Condition_Delete_Click()
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
    Forms![frm_Events]![fsub_Sapling_Data]![chkConditions_Checked].Value = True
    Me.Parent.Refresh
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Sapling_Condtions", "Sapling_Condition_ID") = dbText Then
            Me!Sapling_Condition_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
