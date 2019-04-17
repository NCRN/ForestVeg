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
    Width =6300
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =495
    Top =3390
    Right =7620
    Bottom =5745
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xebd3667b958fe440
    End
    RecordSource ="SELECT tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID, tbl_Quadrat_Herbaceous"
        "_Data.Quadrat_Data_ID, tbl_Quadrat_Herbaceous_Data.TSN, tlu_Plants.Family, tlu_P"
        "lants.Genus, tlu_Plants.Species, tlu_Plants.Subspecies, tbl_Quadrat_Herbaceous_D"
        "ata.Percent_Cover, tbl_Quadrat_Herbaceous_Data.Browse FROM tbl_Quadrat_Herbaceou"
        "s_Data LEFT JOIN tlu_Plants ON tbl_Quadrat_Herbaceous_Data.TSN = tlu_Plants.TSN;"
    Caption ="sfrm_Quad_Seedlings"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
            Height =300
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =540
                    Width =645
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =0
                    Name ="lblTSN"
                    Caption ="Taxon"
                    FontName ="Calibri"
                    LayoutCachedLeft =540
                    LayoutCachedWidth =1185
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3960
                    Width =870
                    Height =300
                    FontSize =11
                    FontWeight =700
                    Name ="lblPercentCover"
                    Caption ="% Cover"
                    FontName ="Calibri"
                    LayoutCachedLeft =3960
                    LayoutCachedWidth =4830
                    LayoutCachedHeight =300
                End
                Begin Label
                    OverlapFlags =85
                    Left =5160
                    Width =780
                    Height =300
                    FontSize =11
                    FontWeight =700
                    Name ="Label19"
                    Caption ="Browse"
                    FontName ="Calibri"
                    LayoutCachedLeft =5160
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =300
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
                    Left =4140
                    Top =60
                    Width =480
                    Height =300
                    FontSize =11
                    Name ="tbxHerbPercentCover"
                    ControlSource ="Percent_Cover"
                    ValidationRule ="(>=0 And <=100) Or Is Null"
                    ValidationText ="Choose a % cover between 0 and 100 (inclusive)"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000d4000000020000000100000000000000000000001a00000001000000 ,
                        0x00000000dfa7a50001000000000000001b000000390000000100000000000000 ,
                        0xdfa7a50000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074007800740048006500720062005f00500065007200630065006e007400 ,
                        0x5f0043006f007600650072005d003d00300000000000490073004e0075006c00 ,
                        0x6c0028005b007400620078004800650072006200500065007200630065006e00 ,
                        0x740043006f007600650072005d00290000000000
                    End

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000dfa7a500190000005b00 ,
                        0x74007800740048006500720062005f00500065007200630065006e0074005f00 ,
                        0x43006f007600650072005d003d00300000000000000000000000000000000000 ,
                        0x000000000001000000000000000100000000000000dfa7a5001d000000490073 ,
                        0x004e0075006c006c0028005b0074006200780048006500720062005000650072 ,
                        0x00630065006e00740043006f007600650072005d002900000000000000000000 ,
                        0x000000000000000000000000
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Top =60
                    Width =456
                    Height =306
                    TabIndex =1
                    Name ="btnDelete"
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
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedTop =60
                    LayoutCachedWidth =456
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
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =40
                    ListWidth =6840
                    Left =540
                    Top =60
                    Width =3300
                    Height =300
                    FontSize =11
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cbxTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Rank_Name, tlu_Plants.C"
                        "ommon FROM tlu_Plants WHERE (((tlu_Plants.Favorite)=True) AND ((tlu_Plants.Targe"
                        "ted_Herb)=True)) ORDER BY tlu_Plants.Latin_Name; "
                    ColumnWidths ="0;2520;1080;3240"
                    OnEnter ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =540
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =360
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5040
                    Top =60
                    Width =900
                    Height =300
                    FontSize =11
                    TabIndex =3
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0048006500720062004200 ,
                        0x72006f007700730065005d0029003d00540072007500650000000000
                    End
                    Name ="cboHerbBrowse"
                    ControlSource ="Browse"
                    RowSourceType ="Value List"
                    RowSource ="\"Yes\";\"No\""
                    FontName ="Calibri"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001c0000004900 ,
                        0x73004e0075006c006c0028005b00630062006f00480065007200620042007200 ,
                        0x6f007700730065005d0029003d00540072007500650000000000000000000000 ,
                        0x0000000000000000000000
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
' FORM:         fsub_Quad_Herbaceous
' Level:        Application report
' Version:      1.01
'
' Description:  Form related functions & procedures for application
' Requires:     Keypad Utils module
'
' Source/date:  Bonnie Campbell, April 20, 2018
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 4/20/2018 - 1.01 - added documentation, error handling
'                                          field renaming cmd>btn, Label>lbl, txt>tbx
'                                          cmd_DBH_Keypad_Click() removed
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
' Adapted:      Bonnie Campbell, April 20, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/20/2018 - added error handling, documentation
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Quadrat_Herbaceous_Data", "Quadrat_Herbaceous_ID") = dbText Then
            Me!Quadrat_Herbaceous_ID = fxnGUIDGen
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Quad_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTSN_Enter
' Description:  combobox click actions
' Requires:     -
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 20, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/20/2018 - added error handling, documentation
' ---------------------------------
Private Sub cbxTSN_Enter()
On Error GoTo Err_Handler

    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim strSpeciesType As String
    Dim frmFormToUpdate As Form
    
    'set keypad form to launch & control on this form to be updated by it
    strKeypadFormName = "frm_Pad_Species"
    strControlToUpdate = "cbxTSN"
    
    'Choose TREE, SAPLING, SEEDLING, CWD, VINE or TARGETED HERB
    strSpeciesType = "TARGETED HERB"
  
    'launch keypad
    Set frmFormToUpdate = Me
    Call OpenSpeciespad(strKeypadFormName, frmFormToUpdate, strControlToUpdate, strSpeciesType)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDBH_Click[fsub_Quad_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxHerbPercentCover_Click
' Description:  textbox click actions
' Requires:     -
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/20/2018 - initial version
' ---------------------------------
Private Sub tbxHerbPercentCover_Click()
On Error GoTo Err_Handler
    
    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'set keypad form to launch & control on this form to be updated by it
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "tbxHerbPercentCover"
    
    'launch keypad
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxHerbPercentCover_Click[fsub_Quad_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxHerbPercentCover_OnExit
' Description:  textbox exit actions
' Note:         Even an empty sub for OnExit for a control will
'               evaluate & trigger Validation rules & Validation text
'               for the control
'               Without this subroutine the validation rule is not
'               evaluated and the text is not displayed.
' Requires:     -
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 20, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/20/2018 - initial version
' ---------------------------------
Private Sub tbxHerbPercentCover_OnExit()
On Error GoTo Err_Handler

    'runs tbx validation rule vs. entered value
    Dim ctl As Control
    Set ctl = tbxHerbPercentCover
    
    If ctl < 0 _
        Or ctl > 100 _
        Or ctl Is Null Then
        
        MsgBox "Please check your percent value." _
            & "Percentages must be between 0 and 100 (inclusive)", _
            vbOKOnly & "Invalid Percent Value"
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxHerbPercentCover_OnExit[fsub_Quad_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnKeypad_Click
' Description:  keypad button click actions
' Requires:     Keypad Utils module
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 20, 2018
' Revisions:
'   ML/GS - unknown - initial version (cmdHerb_Cover_Keypad_Click)
'   BLC - 4/20/2018 - added error handling, documentation
' ---------------------------------
Private Sub btnKeypad_Click()
On Error GoTo Err_Handler

    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'set keypad form to launch & control on this form to be updated by it
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "tbxHerbPercentCover"
    
    'launch keypad
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnKeypadDBH_Click[fsub_Quad_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDelete_Click
' Description:  delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 20, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/20/2018 - added error handling, documentation
'                     validate DBH
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler

    If MsgBox("You are about to DELETE all data for this tree for " _
        & "this herb event only." & vbNewLine & vbNewLine & "Is this OK?", _
        vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Handler
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
            "Error encountered (#" & Err.Number & " - btnDeleteDBH_Click[fsub_Quad_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub
