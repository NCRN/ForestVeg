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
    Width =3246
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =4740
    Top =4410
    Right =8265
    Bottom =6795
    DatasheetGridlinesColor =12632256
    AfterDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0xc94ae1967ba7e340
    End
    RecordSource ="SELECT tbl_Tree_Vines.Tree_Vine_ID, tbl_Tree_Vines.Tree_Data_ID, tbl_Tree_Vines."
        "TSN FROM tbl_Tree_Vines; "
    Caption ="Vines"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000201c0000e010000001000000 ,
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =366
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =1140
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="Tree_Data_ID"
                    ControlSource ="Tree_Data_ID"

                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =1140
                    TabIndex =3
                    Name ="Tree_Vine_ID"
                    ControlSource ="Tree_Vine_ID"

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =300
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =27
                    ListWidth =5040
                    Left =420
                    Top =60
                    Width =2459
                    Height =300
                    ColumnWidth =2850
                    FontSize =11
                    TabIndex =4
                    BoundColumn =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cboTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.Latin_Name, tlu_Plants.TSN FROM tlu_Plants WHERE (((tlu_Plants"
                        ".Vine)=True) AND ((tlu_Plants.Favorite)=True)) ORDER BY tlu_Plants.Latin_Name; "
                    ColumnWidths ="3600;1440"
                    StatusBarText ="Name of species found in tree"
                    OnEnter ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =420
                    LayoutCachedTop =60
                    LayoutCachedWidth =2879
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Top =60
                    Width =351
                    Height =291
                    TabIndex =1
                    Name ="cmd_Tree_Vine_Delete"
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =2940
                    Top =60
                    Width =306
                    Height =306
                    FontSize =12
                    FontWeight =700
                    Name ="cmdAdd_To_Quickfind"
                    Caption ="i"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Show Taxon Details"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =60
                    LayoutCachedWidth =3246
                    LayoutCachedHeight =366
                    Alignment =4
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
' MODULE:       fsub_Tree_Vines
' Level:        Application module
' Version:      1.01
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/16/2018 - 1.01 - added documentation, error handling
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Events
' ----------------

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/16/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tree_Vines", "Tree_Vines_ID") = dbText Then
            Me!Tree_Vines_ID = fxnGUIDGen
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Tree_Vines])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_AfterDelConfirm
' Description:  form actions after confirming a delete
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/16/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_AfterDelConfirm(Status As Integer)
On Error GoTo Err_Handler

    Me.Parent.Refresh
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterDelConfirm[fsub_Tree_Vines])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/16/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_AfterUpdate()
On Error GoTo Err_Handler

    Forms![frm_Events]![fsub_Tree_Data]![chkVinesChecked].value = True
    Me.Parent.Refresh
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterUpdate[fsub_Tree_Vines])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cboTSN_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/16/2018 - added documentation, error handling
' ---------------------------------
Private Sub cboTSN_Enter()
On Error GoTo Err_Handler

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
  strSpeciesType = "VINE"
  
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenSpeciespad(strKeypadFormName, frmFormToUpdate, strControlToUpdate, strSpeciesType)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cboTSN_Enter[fsub_Tree_Vines])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cmd_Tree_Vine_Delete_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/16/2018 - added documentation, error handling
' ---------------------------------
Private Sub cmd_Tree_Vine_Delete_Click()
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
            "Error encountered (#" & Err.Number & " - cmd_Tree_Vine_Delete_Click[fsub_Tree_Vines])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cmdAdd_To_Quickfind_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 16, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/16/2018 - added documentation, error handling
' ---------------------------------
Private Sub cmdAdd_To_Quickfind_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Plants"
    stLinkCriteria = "[TSN]=" & Me!cboTSN
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    'Form_frm_Field_Data_Foliage_Problems.Data_ID.DefaultValue = StringFromGUID(Me!Data_ID)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdAdd_To_Quickfind_Click[fsub_Tree_Vines])"
    End Select
    Resume Exit_Handler
End Sub

'Private Sub cboTSN_Enter()
'On Error GoTo Err_cmdOpenKeyPad_Click
'  'This routine requires the presence of the Keypad_Utils module.
'  Dim strKeypadFormName As String
'  Dim strControlToUpdate As String
'  Dim strSpeciesType As String
'  Dim frmFormToUpdate As Form
'
'  'The two lines below should be changed to reflect the name of the keypad to open
'  '    and the name of the control to be updated.
'  strKeypadFormName = "frm_Pad_Species"
'  strControlToUpdate = "cboTSN"
'  'Choose TREE, SAPLING, SEEDLING, CWD, VINE or TARGETED HERB
'  strSpeciesType = "VINE"
'
'  'The lines below should not usually be edited.
'  Set frmFormToUpdate = Me
'  Call OpenSpeciespad(strKeypadFormName, frmFormToUpdate, strControlToUpdate, strSpeciesType)
'
'Exit_cmdOpenKeyPad_Click:
'  Exit Sub
'Err_cmdOpenKeyPad_Click:
'  MsgBox Err.Description
'  Resume Exit_cmdOpenKeyPad_Click
'End Sub
