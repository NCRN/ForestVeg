Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =12660
    DatasheetFontHeight =9
    ItemSuffix =21
    Left =1410
    Top =5790
    Right =13770
    Bottom =9795
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xde59bba555ace340
    End
    RecordSource ="tbl_Tags_History"
    Caption ="Change Log"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin FormHeader
            Height =0
            BackColor =3751056
            Name ="FormHeader"
        End
        Begin Section
            Height =4020
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3524
                    Top =3570
                    Width =3030
                    Height =359
                    ColumnWidth =4200
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtTag_History_ID"
                    ControlSource ="Tags_History_ID"
                    StatusBarText ="MA. Field data table row identifier (Data_ID)"

                    LayoutCachedLeft =3524
                    LayoutCachedTop =3570
                    LayoutCachedWidth =6554
                    LayoutCachedHeight =3929
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1095
                            Top =3570
                            Width =2369
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag_Species_History"
                            Caption ="Tag_History_ID:"
                            LayoutCachedLeft =1095
                            LayoutCachedTop =3570
                            LayoutCachedWidth =3464
                            LayoutCachedHeight =3929
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2294
                    Top =1440
                    Width =10185
                    Height =1186
                    ColumnWidth =2370
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtHistory_Notes"
                    ControlSource ="Value_History_Notes"
                    StatusBarText ="Comments about this identification change"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740048006900730074006f00 ,
                        0x720079005f004e006f007400650073005d00290000000000
                    End

                    LayoutCachedLeft =2294
                    LayoutCachedTop =1440
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =2626
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b0074007800740048006900730074006f007200 ,
                        0x79005f004e006f007400650073005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1440
                            Width =2129
                            Height =661
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblSpecies_History_Notes"
                            Caption ="Please describe why you made this change"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1440
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =2101
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9854
                    Top =2715
                    Width =2625
                    Height =359
                    TabIndex =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="txtNetwork_User_Name"
                    ControlSource ="Network_User_Name"
                    StatusBarText ="The network user name of the person making the change"

                    LayoutCachedLeft =9854
                    LayoutCachedTop =2715
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =3074
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7425
                            Top =2715
                            Width =2369
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblNetwork_User_Name"
                            Caption ="Network User Name"
                            LayoutCachedLeft =7425
                            LayoutCachedTop =2715
                            LayoutCachedWidth =9794
                            LayoutCachedHeight =3074
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2279
                    Top =3150
                    Width =4260
                    Height =359
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtChange_Date"
                    ControlSource ="Change_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that species identification was changed for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =2279
                    LayoutCachedTop =3150
                    LayoutCachedWidth =6539
                    LayoutCachedHeight =3509
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =75
                            Top =3150
                            Width =2129
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblChange_Date"
                            Caption ="Date of Change"
                            LayoutCachedLeft =75
                            LayoutCachedTop =3150
                            LayoutCachedWidth =2204
                            LayoutCachedHeight =3509
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2279
                    Top =2715
                    Width =4275
                    Height =359
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0043006f006e0074006100 ,
                        0x630074005f00490044005d00290000000000
                    End
                    Name ="cboContact_ID"
                    ControlSource ="Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                        "Last_Name, tlu_Contacts.First_Name;"
                    ColumnWidths ="0;2880"
                    StatusBarText ="M. Contact identifier (Contact_ID)"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2279
                    LayoutCachedTop =2715
                    LayoutCachedWidth =6554
                    LayoutCachedHeight =3074
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500170000004900 ,
                        0x73004e0075006c006c0028005b00630062006f0043006f006e00740061006300 ,
                        0x74005f00490044005d0029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2715
                            Width =2129
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblContact_ID"
                            Caption ="Changed By"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2715
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =3074
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10860
                    Top =3420
                    Width =1605
                    Height =450
                    FontWeight =700
                    TabIndex =7
                    ForeColor =4754549
                    Name ="cmdAccept_Value_Change"
                    Caption ="Accept Change"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =3420
                    LayoutCachedWidth =12465
                    LayoutCachedHeight =3870
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9240
                    Top =3420
                    Width =1605
                    Height =450
                    TabIndex =6
                    ForeColor =3751056
                    Name ="cmdCancel_Value_Change"
                    Caption ="Cancel Change"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9240
                    LayoutCachedTop =3420
                    LayoutCachedWidth =10845
                    LayoutCachedHeight =3870
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2294
                    Top =975
                    Width =3945
                    Height =359
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtValue_New"
                    ControlSource ="Value_New"
                    StatusBarText ="New TSN of Specimen"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00740078007400560061006c0075006500 ,
                        0x5f004e00650077005d00290000000000
                    End

                    LayoutCachedLeft =2294
                    LayoutCachedTop =975
                    LayoutCachedWidth =6239
                    LayoutCachedHeight =1334
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500160000004900 ,
                        0x73004e0075006c006c0028005b00740078007400560061006c00750065005f00 ,
                        0x4e00650077005d002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =600
                            Top =975
                            Width =1589
                            Height =359
                            FontSize =14
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblValue_New"
                            Caption ="New Value"
                            LayoutCachedLeft =600
                            LayoutCachedTop =975
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =1334
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8249
                    Top =960
                    Width =4230
                    Height =359
                    FontSize =14
                    FontWeight =700
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="txtValue_Old"
                    ControlSource ="Value_Old"
                    StatusBarText ="Previous TSN of Specimen"

                    LayoutCachedLeft =8249
                    LayoutCachedTop =960
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =1319
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6705
                            Top =960
                            Width =1424
                            Height =359
                            FontSize =14
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblValue_Old"
                            Caption ="Old Value"
                            LayoutCachedLeft =6705
                            LayoutCachedTop =960
                            LayoutCachedWidth =8129
                            LayoutCachedHeight =1319
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =120
                    Width =1854
                    Height =479
                    FontSize =18
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =3751056
                    Name ="lblTag_Species_History_Header"
                    Caption ="Change Log"
                    GridlineColor =-2147483616
                    HorizontalAnchor =2
                    LayoutCachedLeft =120
                    LayoutCachedWidth =1974
                    LayoutCachedHeight =479
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =525
                    Width =12360
                    Height =315
                    FontSize =13
                    ForeColor =3751056
                    Name ="lblChange_Description"
                    Caption ="Please enter the revised MICROPLOT NUMBER below"
                    LayoutCachedLeft =120
                    LayoutCachedTop =525
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11219
                    Top =60
                    Width =1275
                    Height =299
                    FontWeight =700
                    TabIndex =10
                    BackColor =11056034
                    ForeColor =10921638
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cboTag_ID"
                    ControlSource ="Record_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag FROM tbl_Tags ORDER BY tbl_Tags.Tag; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="MA. Field data table row identifier (Data_ID)"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =11219
                    LayoutCachedTop =60
                    LayoutCachedWidth =12494
                    LayoutCachedHeight =359
                    ForeThemeColorIndex =1
                    ForeShade =65.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10440
                            Top =75
                            Width =690
                            Height =285
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =10921638
                            Name ="lblTag_ID"
                            Caption ="Tag"
                            LayoutCachedLeft =10440
                            LayoutCachedTop =75
                            LayoutCachedWidth =11130
                            LayoutCachedHeight =360
                            ForeThemeColorIndex =1
                            ForeShade =65.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =5760
                    Left =1980
                    Top =2160
                    Width =240
                    Height =315
                    FontSize =12
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboQuick_Comment"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Quick Comments\")) ORDER BY tlu_Enumerations.Sort_Order;"
                    ColumnWidths ="5760"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =2160
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =2475
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2160
                            Width =1860
                            Height =320
                            Name ="cboQuick_Comment_Label"
                            Caption ="Quick Comment ->"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1920
                            LayoutCachedHeight =2480
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =6300
                    Top =960
                    Width =366
                    Height =396
                    FontSize =10
                    FontWeight =900
                    TabIndex =11
                    ForeColor =0
                    Name ="cmdNew_Value_Keypad"
                    Caption ="1"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =960
                    LayoutCachedWidth =6666
                    LayoutCachedHeight =1356
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public ctlToReset As Control
Public frmReferrer As Form

Private Sub cboQuick_Comment_AfterUpdate()
    Me.txtHistory_Notes = LTrim(Me.txtHistory_Notes & " " & Me.cboQuick_Comment)
    Me.txtHistory_Notes.Requery
End Sub

Private Sub cmdAccept_Value_Change_Click()
On Error GoTo Err_cmdAccept_Value_Change_Click

    If Me.Dirty Then Me.Dirty = False
    ctlToReset.Value = Me!Value_New
    DoCmd.Close acForm, Me.Name, acSaveNo
    frmReferrer.SaveRecord
    'Redundant requeries below; improve code
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery
    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tags_History_Summary].Requery
    
Exit_cmdAccept_Value_Change_Click:
    Exit Sub
Err_cmdAccept_Value_Change_Click:
    MsgBox Err.Description
    Resume Exit_cmdAccept_Value_Change_Click
End Sub

Private Sub cmdCancel_Value_Change_Click()
On Error GoTo Err_cmdCancel_Value_Change_Click
        
    'Command below is not needed when for is called from button instead of BeforeUpdate
    'ctlToReset.Value = ctlToReset.OldValue
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
    MsgBox "Value was NOT changed", vbInformation, "Change cancelled"
    DoCmd.Close , , acSaveNo

Exit_cmdCancel_Value_Change_Click:
    Exit Sub
Err_cmdCancel_Value_Change_Click:
    DoCmd.SetWarnings True
    MsgBox Err.Description
    Resume Exit_cmdCancel_Value_Change_Click
End Sub

Private Sub cmdNew_Value_Keypad_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "txtValue_New"
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
'Generate string GUID for Tag_Species_History_ID
    If Me.NewRecord Then
        If GetDataType("tbl_Tags_History_Update", "Tag_History_ID") = dbText Then
            Me!Tag_History_ID = fxnGUIDGen
        End If
    End If
End Sub
