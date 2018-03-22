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
    AllowDeletions = NotDefault
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
    Width =12930
    DatasheetFontHeight =9
    ItemSuffix =18
    Left =6285
    Top =6900
    Right =19470
    Bottom =11055
    RecSrcDt = Begin
        0xde59bba555ace340
    End
    RecordSource ="tbl_Tags_History"
    Caption ="Species Change Log"
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
    DatasheetBackColor12 =16777215
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
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            SizeMode =3
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ComboBox
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin FormHeader
            Height =509
            BackColor =11056034
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Width =9219
                    Height =509
                    FontSize =18
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="lblTag_Species_History_Header"
                    Caption ="Change Log"
                    GridlineColor =-2147483616
                    HorizontalAnchor =2
                    LayoutCachedLeft =60
                    LayoutCachedWidth =9279
                    LayoutCachedHeight =509
                End
            End
        End
        Begin Section
            Height =3660
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9134
                    Top =180
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

                    LayoutCachedLeft =9134
                    LayoutCachedTop =180
                    LayoutCachedWidth =12164
                    LayoutCachedHeight =539
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =6705
                            Top =180
                            Width =2369
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag_Species_History"
                            Caption ="Tag_History_ID:"
                            LayoutCachedLeft =6705
                            LayoutCachedTop =180
                            LayoutCachedWidth =9074
                            LayoutCachedHeight =539
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
                    Top =1125
                    Width =10185
                    Height =1186
                    ColumnWidth =2370
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtHistory_Notes"
                    ControlSource ="Value_History_Notes"
                    StatusBarText ="Comments about this identification change"

                    LayoutCachedLeft =2294
                    LayoutCachedTop =1125
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =2311
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1125
                            Width =2129
                            Height =1186
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblSpecies_History_Notes"
                            Caption ="Please describe why you made this change"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1125
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =2311
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =9854
                    Top =2400
                    Width =2565
                    Height =359
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="txtNetwork_User_Name"
                    ControlSource ="Network_User_Name"
                    StatusBarText ="The network user name of the person making the change"

                    LayoutCachedLeft =9854
                    LayoutCachedTop =2400
                    LayoutCachedWidth =12419
                    LayoutCachedHeight =2759
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7425
                            Top =2400
                            Width =2369
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblNetwork_User_Name"
                            Caption ="Network User Name"
                            LayoutCachedLeft =7425
                            LayoutCachedTop =2400
                            LayoutCachedWidth =9794
                            LayoutCachedHeight =2759
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2279
                    Top =2835
                    Width =4260
                    Height =359
                    TabIndex =6
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
                    LayoutCachedTop =2835
                    LayoutCachedWidth =6539
                    LayoutCachedHeight =3194
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =75
                            Top =2835
                            Width =2129
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblChange_Date"
                            Caption ="Date of Change"
                            LayoutCachedLeft =75
                            LayoutCachedTop =2835
                            LayoutCachedWidth =2204
                            LayoutCachedHeight =3194
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2294
                    Top =180
                    Width =4200
                    Height =389
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    BackColor =11056034
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

                    LayoutCachedLeft =2294
                    LayoutCachedTop =180
                    LayoutCachedWidth =6494
                    LayoutCachedHeight =569
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =45
                            Top =165
                            Width =2100
                            Height =405
                            FontSize =16
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag_ID"
                            Caption ="Tag"
                            LayoutCachedLeft =45
                            LayoutCachedTop =165
                            LayoutCachedWidth =2145
                            LayoutCachedHeight =570
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2279
                    Top =2400
                    Width =4275
                    Height =359
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboContact_ID"
                    ControlSource ="Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & \", \" & [First_Name] AS Contact_N"
                        "ame FROM tlu_Contacts ORDER BY tlu_Contacts.Active, [Last_Name] & \", \" & [Firs"
                        "t_Name]; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="M. Contact identifier (Contact_ID)"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2279
                    LayoutCachedTop =2400
                    LayoutCachedWidth =6554
                    LayoutCachedHeight =2759
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2400
                            Width =2129
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblContact_ID"
                            Caption ="Changed By"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2400
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =2759
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10845
                    Top =3075
                    Width =1605
                    Height =450
                    FontWeight =700
                    TabIndex =8
                    ForeColor =4754549
                    Name ="cmdAccept_Species_Change"
                    Caption ="Accept Change"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10845
                    LayoutCachedTop =3075
                    LayoutCachedWidth =12450
                    LayoutCachedHeight =3525
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9225
                    Top =3075
                    Width =1605
                    Height =450
                    TabIndex =9
                    ForeColor =3751056
                    Name ="cmdCancel_Species_Change"
                    Caption ="Cancel Change"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9225
                    LayoutCachedTop =3075
                    LayoutCachedWidth =10830
                    LayoutCachedHeight =3525
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2294
                    Top =660
                    Width =4200
                    Height =359
                    FontWeight =700
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="txtValue_New"
                    ControlSource ="Value_New"
                    StatusBarText ="New TSN of Specimen"

                    LayoutCachedLeft =2294
                    LayoutCachedTop =660
                    LayoutCachedWidth =6494
                    LayoutCachedHeight =1019
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =600
                            Top =660
                            Width =1589
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblValue_New"
                            Caption ="New Value"
                            LayoutCachedLeft =600
                            LayoutCachedTop =660
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =1019
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8249
                    Top =645
                    Width =4200
                    Height =359
                    FontWeight =700
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="txtValue_Old"
                    ControlSource ="Value_Old"
                    StatusBarText ="Previous TSN of Specimen"

                    LayoutCachedLeft =8249
                    LayoutCachedTop =645
                    LayoutCachedWidth =12449
                    LayoutCachedHeight =1004
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6705
                            Top =645
                            Width =1424
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblValue_Old"
                            Caption ="Old Value"
                            LayoutCachedLeft =6705
                            LayoutCachedTop =645
                            LayoutCachedWidth =8129
                            LayoutCachedHeight =1004
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

Public ctlToReset As Control
Public frmReferrer As Form

Private Sub cmdAccept_Species_Change_Click()
On Error GoTo Err_cmdAccept_Species_Change_Click

    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close acForm, Me.Name, acSaveNo
    frmReferrer.SaveRecord
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery
    
Exit_cmdAccept_Species_Change_Click:
    Exit Sub
Err_cmdAccept_Species_Change_Click:
    MsgBox Err.Description
    Resume Exit_cmdAccept_Species_Change_Click
End Sub

Private Sub cmdCancel_Species_Change_Click()
On Error GoTo Err_cmdCancel_Species_Change_Click
        
    ctlToReset.Value = ctlToReset.OldValue
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
    MsgBox "Species ID was NOT changed", vbInformation, "Change cancelled"
    DoCmd.Close , , acSaveNo

Exit_cmdCancel_Species_Change_Click:
    Exit Sub
Err_cmdCancel_Species_Change_Click:
    DoCmd.SetWarnings True
    MsgBox Err.Description
    Resume Exit_cmdCancel_Species_Change_Click
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Generate string GUID for Tag_Species_History_ID
    If Me.NewRecord Then
        If GetDataType("tbl_Tags_History", "Tag_History_ID") = dbText Then
            Me!Tag_History_ID = fxnGUIDGen
        End If
    End If
End Sub
