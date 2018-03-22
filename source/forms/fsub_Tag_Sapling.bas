Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =13650
    DatasheetFontHeight =9
    ItemSuffix =30
    Left =1665
    Top =4365
    Right =15165
    Bottom =6780
    DatasheetGridlinesColor =15062992
    OrderBy ="[tbl_Tags].[Tag]"
    RecSrcDt = Begin
        0xbb20843c6eaee340
    End
    RecordSource ="tbl_Tags"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
            BackColor =16768194
            Name ="FormHeader"
        End
        Begin Section
            Height =853
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =479
                    Width =839
                    Height =360
                    FontSize =12
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtAzimuth"
                    ControlSource ="Azimuth"
                    StatusBarText ="Azimuth from plot center to specimen (true north)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =479
                    LayoutCachedWidth =9179
                    LayoutCachedHeight =839
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7260
                            Top =479
                            Width =1019
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblAzimuh"
                            Caption ="Azimuth:"
                            LayoutCachedLeft =7260
                            LayoutCachedTop =479
                            LayoutCachedWidth =8279
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10319
                    Top =479
                    Width =900
                    Height =360
                    FontSize =12
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtDistance"
                    ControlSource ="Distance"
                    StatusBarText ="Distance (m) from plot center to near EDGE of tree"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10319
                    LayoutCachedTop =479
                    LayoutCachedWidth =11219
                    LayoutCachedHeight =839
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9240
                            Top =479
                            Width =1019
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblDistance"
                            Caption ="Distance:"
                            LayoutCachedLeft =9240
                            LayoutCachedTop =479
                            LayoutCachedWidth =10259
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12359
                    Top =479
                    Width =839
                    Height =360
                    FontSize =12
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtMicroplot_Number"
                    ControlSource ="Microplot_Number"
                    StatusBarText ="The Microplot location of specimen"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b004d006900630072006f0070006c006f00 ,
                        0x74005f004e0075006d006200650072005d00290000000000
                    End

                    LayoutCachedLeft =12359
                    LayoutCachedTop =479
                    LayoutCachedWidth =13198
                    LayoutCachedHeight =839
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b004d006900630072006f0070006c006f007400 ,
                        0x5f004e0075006d006200650072005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11280
                            Top =479
                            Width =1019
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblMicroplot_Number"
                            Caption ="Microplot:"
                            LayoutCachedLeft =11280
                            LayoutCachedTop =479
                            LayoutCachedWidth =12299
                            LayoutCachedHeight =839
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8340
                    Top =60
                    Width =1620
                    Height =360
                    FontSize =12
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtStart_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking began on this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7079
                            Top =60
                            Width =1199
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblStart_Date"
                            Caption ="Start_Date:"
                            LayoutCachedLeft =7079
                            LayoutCachedTop =60
                            LayoutCachedWidth =8278
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =11639
                    Top =60
                    Width =1559
                    Height =360
                    FontSize =12
                    TabIndex =6
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtStop_Date"
                    ControlSource ="Stop_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking ended for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =11639
                    LayoutCachedTop =60
                    LayoutCachedWidth =13198
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10319
                            Top =60
                            Width =1260
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblStop_Date"
                            Caption ="Stop_Date:"
                            LayoutCachedLeft =10319
                            LayoutCachedTop =60
                            LayoutCachedWidth =11579
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin ComboBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =5040
                    Left =2880
                    Top =60
                    Width =3839
                    Height =360
                    FontSize =12
                    FontWeight =700
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN, tlu_Plants.Favorite, IIf([Latin_Name]=\"Kalmia latifolia\""
                        ",[Latin_Name] & \"***\",IIf([Latin_Name]=\"Lindera benzoin\",[Latin_Name] & \"**"
                        "*\",IIf([Latin_Name]=\"Ilex verticillata\",[Latin_Name] & \"***\",[Latin_Name]))"
                        ") AS Name, IIf([Tree]=True,\"Tree\",\"Shrub\") AS Habit FROM tlu_Plants WHERE (("
                        "(tlu_Plants.Tree)=True) AND ((tlu_Plants.Accepted_Found)=False)) OR (((tlu_Plant"
                        "s.Accepted_Found)=False) AND ((tlu_Plants.Shrub)=True)) ORDER BY tlu_Plants.Favo"
                        "rite, IIf([Latin_Name]=\"Kalmia latifolia\",[Latin_Name] & \"***\",IIf([Latin_Na"
                        "me]=\"Lindera benzoin\",[Latin_Name] & \"***\",IIf([Latin_Name]=\"Ilex verticill"
                        "ata\",[Latin_Name] & \"***\",[Latin_Name])));"
                    ColumnWidths ="0;0;3600;1440"
                    StatusBarText ="TSN of Specimen"
                    BeforeUpdate ="[Event Procedure]"
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =6719
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2039
                            Top =60
                            Width =780
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblSpecies"
                            Caption ="Taxon:"
                            LayoutCachedLeft =2039
                            LayoutCachedTop =60
                            LayoutCachedWidth =2819
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1545
                    Height =479
                    FontSize =16
                    FontWeight =700
                    TabIndex =7
                    Name ="txtTag"
                    ControlSource ="Tag"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =539
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =119
                    Top =540
                    Width =1455
                    Height =270
                    FontSize =9
                    TabIndex =8
                    ForeColor =0
                    Name ="cmdReplace_Tag"
                    Caption ="Replace Tag"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =119
                    LayoutCachedTop =540
                    LayoutCachedWidth =1574
                    LayoutCachedHeight =810
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
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2880
                    Top =479
                    Width =2340
                    Height =374
                    FontSize =12
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="txtTag_Status"
                    ControlSource ="Tag_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Tag Status\")) ORDER BY tlu_Enumerations.Sort_Order;"
                    StatusBarText ="Last sampled as tree or sapling?"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2880
                    LayoutCachedTop =479
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =853
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1679
                            Top =479
                            Width =1140
                            Height =360
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag_Status"
                            Caption ="Tag Status:"
                            LayoutCachedLeft =1679
                            LayoutCachedTop =479
                            LayoutCachedWidth =2819
                            LayoutCachedHeight =839
                        End
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

Private Sub cboTag_Status_BeforeUpdate(Cancel As Integer)
    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!cboTag_Status
    ChangeDescription = "Please confirm the revised TAG STATUS below"
    ChangeFieldType = "Combo_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenConfirmValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, "tbl_Tags", "Tag_Status", "Tag_ID", Me!Tag_ID, Nz(Me!cboTag_Status.OldValue, "Null"), Me!cboTag_Status, "", "", ""
End Sub

Private Sub cboTSN_BeforeUpdate(Cancel As Integer)
    'Dim frm As Form
    'Dim ctl As Control
    
    'Set frm = Me
    'Set ctl = Me!cboTSN
    
    'OpenChangeHistory frm, ctl, "tbl_Tags", "TSN", "Tag_ID", Me!Tag_ID, Me!cboTSN.OldValue, Me!cboTSN, "tlu_Plants", "Latin_Name", "TSN", dbLong
    
    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!cboTSN
    ChangeDescription = "Please confirm the revised SPECIES ID below"
    ChangeFieldType = "Combo_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenConfirmValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, "tbl_Tags", "TSN", "Tag_ID", Me!Tag_ID, Nz(Me!cboTSN.OldValue, "Null"), Me!cboTSN, "tlu_Plants", "Latin_Name", "TSN"
End Sub

Public Sub SaveRecord()
    DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub cmdReplace_Tag_Click()
On Error GoTo Err_cmdReplace_Tag_Click

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!txtTag
    ChangeDescription = "Please enter the new TAG NUMBER below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, "tbl_Tags", "Tag", "Tag_ID", Me!Tag_ID, Me!Tag.Value

Exit_cmdReplace_Tag_Click:
    Exit Sub
Err_cmdReplace_Tag_Click:
    MsgBox Err.Description
    Resume Exit_cmdReplace_Tag_Click
End Sub

Private Sub txtAzimuth_Click()
On Error GoTo Err_txtAzimuth_Click

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!txtAzimuth
    ChangeDescription = "Please enter the revised AZIMUTH below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, "tbl_Tags", "Azimuth", "Tag_ID", Me!Tag_ID, Nz(Me!Azimuth.Value, "Null")

Exit_txtAzimuth_Click:
    Exit Sub
Err_txtAzimuth_Click:
    MsgBox Err.Description
    Resume Exit_txtAzimuth_Click
End Sub

Private Sub txtDistance_Click()
On Error GoTo Err_txtDistance_Click

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!txtDistance
    ChangeDescription = "Please enter the revised DISTANCE below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, "tbl_Tags", "Distance", "Tag_ID", Me!Tag_ID, Nz(Me!Distance.Value, "Null")

Exit_txtDistance_Click:
    Exit Sub
Err_txtDistance_Click:
    MsgBox Err.Description
    Resume Exit_txtDistance_Click
End Sub

Private Sub txtMicroplot_Number_Click()
On Error GoTo Err_txtMicroplot_Number_Click

    Dim frm As Form
    Dim ctl As Control
    Dim ChangeDescription As String
    Dim ChangeFieldType As String
    
    Set frm = Me
    Set ctl = Me!txtMicroplot_Number
    ChangeDescription = "Please enter the revised MICROPLOT NUMBER below"
    ChangeFieldType = "Text_Box"
    
    'strChangeDescription,strChangeFieldType,frmFormToSave,ctlControlToReset,strTableName,strFieldName,strRecordIDFieldName,strRecordID,strOldValue
    OpenChangeValueAndLog ChangeDescription, ChangeFieldType, frm, ctl, "tbl_Tags", "Microplot_Number", "Tag_ID", Me!Tag_ID, Nz(Me!Microplot_Number.Value, "Null")

Exit_txtMicroplot_Number_Click:
    Exit Sub
Err_txtMicroplot_Number_Click:
    MsgBox Err.Description
    Resume Exit_txtMicroplot_Number_Click
End Sub
