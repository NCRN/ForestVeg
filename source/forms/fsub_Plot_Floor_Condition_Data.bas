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
    Cycle =1
    GridX =24
    GridY =24
    Width =4846
    DatasheetFontHeight =9
    ItemSuffix =18
    Left =1320
    Top =6330
    Right =7305
    Bottom =11370
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x816d9c7d10a7e340
    End
    RecordSource ="tbl_Plot_Floor_Condition_Data"
    BeforeUpdate ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =16768194
            Name ="FormHeader"
            AutoHeight =1
        End
        Begin Section
            Height =1433
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2670
                    Top =120
                    Width =1110
                    Height =374
                    FontSize =13
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0052006f0063006b005f00 ,
                        0x43006f007600650072005d0029003d00540072007500650000000000
                    End
                    Name ="cboRock_Cover"
                    ControlSource ="Rock_Cover"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Percent Broad\")) ORDER BY tlu_Enum"
                        "erations.Sort_Order; "
                    StatusBarText ="Percent of the plot covered by rocks"
                    ValidationText ="Choose a % cover between 0 and 100 (inclusive)"
                    GroupTable =10
                    RightPadding =38
                    BottomPadding =38
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2670
                    LayoutCachedTop =120
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =494
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001c0000004900 ,
                        0x73004e0075006c006c0028005b00630062006f0052006f0063006b005f004300 ,
                        0x6f007600650072005d0029003d00540072007500650000000000000000000000 ,
                        0x0000000000000000000000
                    End
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =120
                            Width =2445
                            Height =374
                            FontSize =13
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblRock_Cover"
                            Caption =" % Rock Cover"
                            GroupTable =10
                            BottomPadding =38
                            LayoutCachedLeft =165
                            LayoutCachedTop =120
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =494
                            LayoutGroup =1
                            GroupTable =10
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2670
                    Top =570
                    Width =1110
                    Height =375
                    FontSize =13
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x01000000a6000000010000000100000000000000000000002200000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0042006100720065005f00 ,
                        0x53006f0069006c005f0043006f007600650072005d0029003d00540072007500 ,
                        0x650000000000
                    End
                    Name ="cboBare_Soil_Cover"
                    ControlSource ="Bare_Soil_Cover"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Percent Broad\")) ORDER BY tlu_Enum"
                        "erations.Sort_Order; "
                    StatusBarText ="Percent of the plot covered by bare soil"
                    ValidationText ="Choose a % cover between 0 and 100 (inclusive)"
                    GroupTable =10
                    RightPadding =38
                    BottomPadding =38
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2670
                    LayoutCachedTop =570
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =945
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500210000004900 ,
                        0x73004e0075006c006c0028005b00630062006f0042006100720065005f005300 ,
                        0x6f0069006c005f0043006f007600650072005d0029003d005400720075006500 ,
                        0x000000000000000000000000000000000000000000
                    End
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =570
                            Width =2445
                            Height =375
                            FontSize =13
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label11"
                            Caption ="% Bare Soil Cover:"
                            GroupTable =10
                            BottomPadding =38
                            LayoutCachedLeft =165
                            LayoutCachedTop =570
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =945
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =10
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2670
                    Top =1020
                    Width =1110
                    Height =375
                    FontSize =13
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f005400720061006d007000 ,
                        0x6c00650064005d0029003d00540072007500650000000000
                    End
                    Name ="cboTrampled"
                    ControlSource ="Trampled"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Percent Broad\")) ORDER BY tlu_Enum"
                        "erations.Sort_Order; "
                    StatusBarText ="Percent of the plot trampled"
                    ValidationText ="Choose a % cover between 0 and 100 (inclusive)"
                    GroupTable =10
                    RightPadding =38
                    BottomPadding =38
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2670
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1395
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b00630062006f005400720061006d0070006c00 ,
                        0x650064005d0029003d0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    GroupTable =10
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =165
                            Top =1020
                            Width =2445
                            Height =375
                            FontSize =13
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label13"
                            Caption ="% Trampled:"
                            GroupTable =10
                            BottomPadding =38
                            LayoutCachedLeft =165
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2610
                            LayoutCachedHeight =1395
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =10
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
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Plot_Floor_Condition_Data", "Plot_Floor_Data_ID") = dbText Then
            Me!Plot_Floor_Data_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
