Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8610
    DatasheetFontHeight =9
    ItemSuffix =21
    Left =2055
    Top =8790
    Right =10920
    Bottom =11805
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x612a1090ccade340
    End
    RecordSource ="SELECT tbl_Tasks.Task_ID, tbl_Tasks.Location_ID, tbl_Tasks.Event_ID, tbl_Tasks.T"
        "ask_Date, tbl_Tasks.Task_Contact_ID, tbl_Tasks.Task_Notes, tbl_Tasks.Task_Status"
        ", tbl_Tasks.Followup_Date, tbl_Tasks.Followup_Contact_ID, tbl_Tasks.Followup_Not"
        "es FROM tbl_Tasks ORDER BY tbl_Tasks.Task_Date DESC; "
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
    OrderByOnLoad =0
    OrderByOnLoad =0
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
            BorderLineStyle =0
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
        Begin Subform
            BorderColor =12632256
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
        Begin FormHeader
            Height =0
            BackColor =15527148
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =3033
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =750
                    Top =60
                    Height =315
                    FontSize =10
                    Name ="cboTask_Status"
                    ControlSource ="Task_Status"
                    RowSourceType ="Value List"
                    RowSource ="Active;Complete;Inactive"
                    StatusBarText ="Status of the task"
                    AllowValueListEdits =0

                    LayoutCachedLeft =750
                    LayoutCachedTop =60
                    LayoutCachedWidth =2190
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =660
                            Height =315
                            FontSize =10
                            FontWeight =700
                            Name ="Label8"
                            Caption ="Status"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =720
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2625
                    Top =450
                    Height =252
                    FontSize =10
                    TabIndex =1
                    Name ="txtTask_Date"
                    ControlSource ="Task_Date"
                    Format ="Short Date"
                    StatusBarText ="Task creation date"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =2625
                    LayoutCachedTop =450
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =702
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2070
                            Top =450
                            Width =525
                            Height =252
                            FontSize =10
                            Name ="Label9"
                            Caption ="Date"
                            LayoutCachedLeft =2070
                            LayoutCachedTop =450
                            LayoutCachedWidth =2595
                            LayoutCachedHeight =702
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =765
                    Top =765
                    Width =7845
                    Height =948
                    ColumnWidth =6210
                    FontSize =10
                    TabIndex =2
                    Name ="txtTask_Notes"
                    ControlSource ="Task_Notes"
                    StatusBarText ="Description of Task"

                    LayoutCachedLeft =765
                    LayoutCachedTop =765
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =1713
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =765
                            Width =660
                            Height =315
                            FontSize =10
                            Name ="Label11"
                            Caption ="Notes"
                            LayoutCachedLeft =60
                            LayoutCachedTop =765
                            LayoutCachedWidth =720
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =4860
                    Top =450
                    Width =1923
                    Height =252
                    FontSize =10
                    TabIndex =3
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboTask_Contact_ID"
                    ControlSource ="Task_Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                        "Last_Name, tlu_Contacts.First_Name; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="Observer identifier"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =4860
                    LayoutCachedTop =450
                    LayoutCachedWidth =6783
                    LayoutCachedHeight =702
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4110
                            Top =450
                            Width =720
                            Height =252
                            FontSize =10
                            Name ="Label12"
                            Caption ="Contact"
                            LayoutCachedLeft =4110
                            LayoutCachedTop =450
                            LayoutCachedWidth =4830
                            LayoutCachedHeight =702
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =75
                    Top =420
                    Width =1800
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label13"
                    Caption ="Task Details"
                    LayoutCachedLeft =75
                    LayoutCachedTop =420
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =705
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2625
                    Top =1770
                    Height =252
                    FontSize =10
                    TabIndex =4
                    Name ="txtFollowup_Date"
                    ControlSource ="Followup_Date"
                    Format ="Short Date"
                    StatusBarText ="Task creation date"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =2625
                    LayoutCachedTop =1770
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =2022
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2070
                            Top =1770
                            Width =525
                            Height =252
                            FontSize =10
                            Name ="Label15"
                            Caption ="Date"
                            LayoutCachedLeft =2070
                            LayoutCachedTop =1770
                            LayoutCachedWidth =2595
                            LayoutCachedHeight =2022
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =765
                    Top =2085
                    Width =7830
                    Height =945
                    FontSize =10
                    TabIndex =5
                    Name ="txtFollowup_Notes"
                    ControlSource ="Followup_Notes"
                    StatusBarText ="Description of Task"

                    LayoutCachedLeft =765
                    LayoutCachedTop =2085
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =3030
                    ConditionalFormat14 = Begin
                        0x010000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2085
                            Width =660
                            Height =315
                            FontSize =10
                            Name ="Label17"
                            Caption ="Notes"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2085
                            LayoutCachedWidth =720
                            LayoutCachedHeight =2400
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =4860
                    Top =1770
                    Width =1923
                    Height =252
                    FontSize =10
                    TabIndex =6
                    BackColor =-2147483643
                    BorderColor =0
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboFollowup_Contact_ID"
                    ControlSource ="Followup_Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                        "Last_Name, tlu_Contacts.First_Name; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="Observer identifier"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =4860
                    LayoutCachedTop =1770
                    LayoutCachedWidth =6783
                    LayoutCachedHeight =2022
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4110
                            Top =1770
                            Width =720
                            Height =252
                            FontSize =10
                            Name ="Label19"
                            Caption ="Contact"
                            LayoutCachedLeft =4110
                            LayoutCachedTop =1770
                            LayoutCachedWidth =4830
                            LayoutCachedHeight =2022
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =75
                    Top =1740
                    Width =1800
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label20"
                    Caption ="Follow-up Details"
                    LayoutCachedLeft =75
                    LayoutCachedTop =1740
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =2025
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

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tasks", "Task_ID") = dbText Then
            Me!Task_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub
