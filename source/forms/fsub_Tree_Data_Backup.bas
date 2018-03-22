Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13950
    DatasheetFontHeight =9
    ItemSuffix =52
    Left =1800
    Top =3840
    Right =15570
    Bottom =10155
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x4d5502714caae340
    End
    RecordSource ="tbl_Tree_Data"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Height =0
            BackColor =16768194
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =6480
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2384
                    Top =4650
                    Width =11505
                    Height =256
                    ColumnWidth =2055
                    FontSize =10
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =0
                    Name ="txtComments"
                    ControlSource ="Tree_Notes"
                    StatusBarText ="Notes about this sampling of this tree"

                    LayoutCachedLeft =2414
                    LayoutCachedTop =4395
                    LayoutCachedWidth =13919
                    LayoutCachedHeight =4651
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =60
                            Top =4650
                            Width =2249
                            Height =256
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BackColor =15527148
                            Name ="Label19"
                            Caption ="Comments:"
                            LayoutCachedLeft =90
                            LayoutCachedTop =4395
                            LayoutCachedWidth =2339
                            LayoutCachedHeight =4651
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5759
                    Top =1965
                    Width =210
                    Height =269
                    TabIndex =3
                    Name ="chkVines_Checked"
                    ControlSource ="Vines_Checked"
                    StatusBarText ="This tree was checked for vines"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =5789
                    LayoutCachedTop =1710
                    LayoutCachedWidth =5999
                    LayoutCachedHeight =1979
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9959
                    Top =1965
                    Width =210
                    Height =269
                    TabIndex =4
                    Name ="chkConditions_Checked"
                    ControlSource ="Conditions_Checked"
                    StatusBarText ="This tree was checked for disease/damage conditions"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =9989
                    LayoutCachedTop =1710
                    LayoutCachedWidth =10199
                    LayoutCachedHeight =1979
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =13619
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =5
                    Name ="chkFoliage_Conditions_Checked"
                    ControlSource ="Foliage_Conditions_Checked"
                    StatusBarText ="This tree was checked for foliage conditions"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =13649
                    LayoutCachedTop =1725
                    LayoutCachedWidth =13859
                    LayoutCachedHeight =1934
                End
                Begin Subform
                    OverlapFlags =85
                    BorderWidth =2
                    Left =60
                    Top =435
                    Width =13680
                    Height =1065
                    TabIndex =6
                    BorderColor =7633277
                    Name ="fsub_Tag_Tree"
                    SourceObject ="Form.fsub_Tag_Tree"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =435
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =1215
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    ListRows =20
                    ListWidth =6120
                    Left =1440
                    Top =60
                    Width =240
                    Height =315
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cboTag_Finder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, tbl_Tags.Tag_Status, IIf(IsNull([azimuth])"
                        ",\"\",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, tbl_Tags.Microplot_"
                        "Number AS MP, qry_Status_Tree_Current_Event.Tree_Status FROM tbl_Tags LEFT JOIN "
                        "qry_Status_Tree_Current_Event ON tbl_Tags.Tag_ID=qry_Status_Tree_Current_Event.T"
                        "ag_ID WHERE (((tbl_Tags.Location_ID)=[Forms]![frm_Events]![Location_ID])) ORDER "
                        "BY tbl_Tags.Tag_Status DESC , tbl_Tags.Tag; "
                    ColumnWidths ="0;936;1080;1224;936;1800"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =1365
                            Height =315
                            Name ="Label44"
                            Caption ="Select a tag ->"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =1425
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =90
                    Top =2220
                    Width =2190
                    Height =2340
                    TabIndex =8
                    Name ="fsub_Tree_DBH"
                    SourceObject ="Form.fsub_Tree_DBH"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1965
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =4305
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =204
                            Left =150
                            Top =1980
                            Width =1440
                            Height =240
                            FontSize =10
                            Name ="fsub_Tree_DBH Label"
                            Caption ="Stems (cm)"
                            EventProcPrefix ="fsub_Tree_DBH_Label"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1725
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =1965
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =119
                    Left =2370
                    Top =2220
                    Width =3660
                    Height =2340
                    TabIndex =9
                    Name ="fsub_Tree_Vines"
                    SourceObject ="Form.fsub_Tree_Vines"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =1965
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =4305
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =204
                            Left =2430
                            Top =1980
                            Width =600
                            Height =240
                            FontSize =10
                            Name ="fsub_Tree_Vines Label"
                            Caption ="Vines"
                            EventProcPrefix ="fsub_Tree_Vines_Label"
                            LayoutCachedLeft =2460
                            LayoutCachedTop =1725
                            LayoutCachedWidth =3060
                            LayoutCachedHeight =1965
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =119
                    Left =6150
                    Top =2220
                    Width =4080
                    Height =2340
                    TabIndex =10
                    Name ="fsub_Tree_Conditions"
                    SourceObject ="Form.fsub_Tree_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =1965
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =4305
                End
                Begin Subform
                    OverlapFlags =87
                    Left =10350
                    Top =2220
                    Width =3540
                    Height =2340
                    TabIndex =11
                    Name ="fsub_Tree_Foliage_Conditions"
                    SourceObject ="Form.fsub_Tree_Foliage_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =10380
                    LayoutCachedTop =1965
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =4305
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =204
                            Left =10410
                            Top =1980
                            Width =1620
                            Height =240
                            FontSize =10
                            Name ="fsub_Tree_Foliage_Conditions Label"
                            Caption ="Foliage Conditions"
                            EventProcPrefix ="fsub_Tree_Foliage_Conditions_Label"
                            LayoutCachedLeft =10440
                            LayoutCachedTop =1725
                            LayoutCachedWidth =12060
                            LayoutCachedHeight =1965
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1334
                    Top =1620
                    Width =2565
                    Height =314
                    FontSize =10
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"12\";\"0\""
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00430072006f0077006e005f0043006c00 ,
                        0x6100730073005d0029003d00540072007500650000000000
                    End
                    Name ="Crown_Class"
                    ControlSource ="Crown_Class"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description, tlu_Enumer"
                        "ations.Enum_Group FROM tlu_Enumerations WHERE (((tlu_Enumerations.Enum_Group)=\""
                        "Crown Class\")) ORDER BY tlu_Enumerations.Sort_Order; "
                    ColumnWidths ="0;1440"
                    StatusBarText ="Options: (1)open-grown (2)Dominant (3)Co-dominant (4)Intermediate (5)Overtopped"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1394
                    LayoutCachedTop =1290
                    LayoutCachedWidth =3959
                    LayoutCachedHeight =1604
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b00430072006f0077006e005f0043006c006100 ,
                        0x730073005d0029003d0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5759
                    Top =1620
                    Width =2820
                    Height =314
                    ColumnWidth =1875
                    FontSize =10
                    TabIndex =1
                    BorderColor =0
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0054007200650065005f00 ,
                        0x5300740061007400750073005d0029003d00540072007500650000000000
                    End
                    Name ="cboTree_Status"
                    ControlSource ="Tree_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Tree Status\")) ORDER BY tlu_Enumer"
                        "ations.Sort_Order; "
                    ColumnWidths ="3168"
                    StatusBarText ="Health status of this specimen"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =5759
                    LayoutCachedTop =1620
                    LayoutCachedWidth =8579
                    LayoutCachedHeight =1934
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                        0x73004e0075006c006c0028005b00630062006f0054007200650065005f005300 ,
                        0x740061007400750073005d0029003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =4680
                            Top =1620
                            Width =1019
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BackColor =15527148
                            Name ="Label17"
                            Caption ="Tree Status:"
                            LayoutCachedLeft =4680
                            LayoutCachedTop =1620
                            LayoutCachedWidth =5699
                            LayoutCachedHeight =1904
                        End
                    End
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =120
                    Top =1620
                    Width =1206
                    Height =306
                    FontSize =10
                    TabIndex =12
                    ForeColor =6108695
                    Name ="cmdOpen_Form_Crown_Class"
                    Caption ="Crown Class"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =180
                    LayoutCachedTop =1290
                    LayoutCachedWidth =1386
                    LayoutCachedHeight =1596
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    Left =6210
                    Top =1920
                    Width =2106
                    Height =306
                    FontSize =10
                    TabIndex =13
                    ForeColor =6108695
                    Name ="cmdOpen_Form_Conditions_and_Pests"
                    Caption ="Conditions and Pests"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =6240
                    LayoutCachedTop =1665
                    LayoutCachedWidth =8346
                    LayoutCachedHeight =1971
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Subform
                    OverlapFlags =215
                    Left =90
                    Top =5175
                    Width =13770
                    Height =1275
                    TabIndex =14
                    Name ="fsub_Tags_History_Summary"
                    SourceObject ="Form.fsub_Tags_History_Summary"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedLeft =90
                    LayoutCachedTop =5175
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =6450
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =204
                            Left =90
                            Top =4935
                            Width =1200
                            Height =315
                            FontSize =10
                            Name ="fsub_Tags_History_Summary Label"
                            Caption ="Tag History"
                            EventProcPrefix ="fsub_Tags_History_Summary_Label"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4680
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =4995
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2445
                    Top =30
                    Width =1845
                    FontSize =10
                    TabIndex =15
                    Name ="cmdTag_New_Specimen"
                    Caption ="Tag New Specimen"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2445
                    LayoutCachedTop =30
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =390
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =1950
                    Top =60
                    Width =270
                    Height =285
                    Name ="Label49"
                    Caption ="or"
                    LayoutCachedLeft =1950
                    LayoutCachedTop =60
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4710
                    Top =1920
                    Width =975
                    Height =285
                    FontSize =10
                    TabIndex =16
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =0
                    Name ="lblVines_Checked"
                    ControlSource ="=\"Completed\""
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000100 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00560069006e00650073005f0043006800650063006b00650064005d003c00 ,
                        0x3e00540072007500650000000000
                    End

                    LayoutCachedLeft =4710
                    LayoutCachedTop =1920
                    LayoutCachedWidth =5685
                    LayoutCachedHeight =2205
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100010000000000dfa7a500150000005b00 ,
                        0x560069006e00650073005f0043006800650063006b00650064005d003c003e00 ,
                        0x5400720075006500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8970
                    Top =1920
                    Width =929
                    Height =269
                    FontSize =10
                    TabIndex =17
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =0
                    Name ="lblConditions_Checked"
                    ControlSource ="=\"Completed\""
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000100 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0043006f006e0064006900740069006f006e0073005f004300680065006300 ,
                        0x6b00650064005d003c003e00540072007500650000000000
                    End

                    LayoutCachedLeft =8970
                    LayoutCachedTop =1920
                    LayoutCachedWidth =9899
                    LayoutCachedHeight =2189
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100010000000000dfa7a5001a0000005b00 ,
                        0x43006f006e0064006900740069006f006e0073005f0043006800650063006b00 ,
                        0x650064005d003c003e0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12570
                    Top =1935
                    Width =975
                    Height =285
                    FontSize =10
                    TabIndex =18
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =0
                    Name ="lblFoliage_Conditions_Checked"
                    ControlSource ="=\"Completed\""
                    ConditionalFormat = Begin
                        0x01000000a8000000010000000100000000000000000000002300000001000100 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0046006f006c0069006100670065005f0043006f006e006400690074006900 ,
                        0x6f006e0073005f0043006800650063006b00650064005d003c003e0054007200 ,
                        0x7500650000000000
                    End

                    LayoutCachedLeft =12570
                    LayoutCachedTop =1935
                    LayoutCachedWidth =13545
                    LayoutCachedHeight =2220
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100010000000000dfa7a500220000005b00 ,
                        0x46006f006c0069006100670065005f0043006f006e0064006900740069006f00 ,
                        0x6e0073005f0043006800650063006b00650064005d003c003e00540072007500 ,
                        0x6500000000000000000000000000000000000000000000
                    End
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

Private Sub cboTag_Finder_AfterUpdate()
    ' Find the record that matches the control, if record doesn't exist, create it.
    
    On Error GoTo HandleErrors
    
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cboTag_Finder.Column(0)
    strSearchField = "Tag_ID"
    
    If Me!cboTag_Finder.Column(2) = "Sapling" Then
        If MsgBox("You are upgrading a SAPLING to a TREE.  Is this OK?", vbOKCancel) = vbCancel Then GoTo ExitHere
    End If
        
    'Search for a matching record
    Set rstClone = Me.Recordset.Clone
    
    Do Until rstClone.EOF
        If rstClone(strSearchField) = strFind Then
            'Goto matching record and exit subroutine
            Me.Bookmark = rstClone.Bookmark
            GoTo ExitHere
        End If
        rstClone.MoveNext
    Loop
    'If we haven't found record and exited by now, create new record.
    DoCmd.GoToRecord , , acNewRec
    Tag_ID.Value = strFind
    DoCmd.RunCommand acCmdSaveRecord
    Me!fsub_Tag_Tree.Requery
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tag_Tree]!txtTag_Status = "Tree"
    Me!fsub_Tag_Tree.Requery
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery
    
ExitHere:
    Exit Sub
HandleErrors:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            rst.CancelUpdate 'I hope this is the correct fix.
        Case 3021 'record not found .... Mel says DOUBLE CHECK
            MsgBox ("Case 3021 error cboTagFinder code")
            DoCmd.GoToRecord , , acNewRec
            txtTag_ID.Value = Me!cboTag_Finder.Column(0)
            DoCmd.RunCommand acCmdSaveRecord
            Me!fsub_Tree_Data.Requery
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error encountered in procedure" & strProcName
            Exit Sub
    End Select
 
End Sub

Private Sub cboTag_Finder_Enter()
    Me!cboTag_Finder.Requery
End Sub

Private Sub chkConditions_Checked_AfterUpdate()
    lblConditions_Checked.Requery
End Sub

Private Sub chkFoliage_Conditions_Checked_AfterUpdate()
    lblFoliage_Conditions_Checked.Requery
End Sub

Private Sub chkVines_Checked_AfterUpdate()
    lblVines_Checked.Requery
End Sub

Private Sub cmdOpen_Form_Conditions_and_Pests_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Conditions_and_Pests"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Popup_Click:
    Exit Sub
Err_cmdOpen_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Popup_Click
End Sub

Private Sub cmdOpen_Form_Crown_Class_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Crown_Classes"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Popup_Click:
    Exit Sub
Err_cmdOpen_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Popup_Click
End Sub

Private Sub cmdTag_New_Specimen_Click()
On Error GoTo Err_cmdTag_New_Specimen_Click
    Dim strCriteria As String

    strCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Parent.Name, "txtLocation_ID")
    DoCmd.OpenForm "frm_Locations", , , strCriteria, , , "Filter by location"

Exit_cmdTag_New_Specimen_Click:
    Exit Sub
Err_cmdTag_New_Specimen_Click:
    MsgBox Err.Description
    Resume Exit_cmdTag_New_Specimen_Click
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tree_Data", "Tree_Data_ID") = dbText Then
            Me!Tree_Data_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
