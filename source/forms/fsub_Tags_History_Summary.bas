Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11820
    DatasheetFontHeight =9
    ItemSuffix =23
    Left =1500
    Top =9885
    Right =14790
    Bottom =11745
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xb6c68d90e9aee340
    End
    RecordSource ="qFsub_Tag_Note_History"
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
        Begin Line
            BorderLineStyle =0
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
            AutoHeight =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =315
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =15
                    Width =2385
                    Height =269
                    ColumnWidth =2760
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtChange_Date"
                    ControlSource ="Change_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that species identification was changed for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =60
                    LayoutCachedTop =15
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =284
                End
                Begin Line
                    OverlapFlags =93
                    Left =60
                    Top =300
                    Width =11760
                    Name ="Line16"
                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    BorderWidth =1
                    OverlapFlags =87
                    TextFontCharSet =204
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2505
                    Top =15
                    Width =9315
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="txtTag_History_Description"
                    ControlSource ="=[Change_Desc]"

                    LayoutCachedLeft =2505
                    LayoutCachedTop =15
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =300
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

Private Sub cmdDeleteTagHistory_Click()
On Error GoTo Err_cmdDeleteTagHistory_Click


    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDeleteRecord

Exit_cmdDeleteTagHistory_Click:
    Exit Sub

Err_cmdDeleteTagHistory_Click:
    MsgBox Err.Description
    Resume Exit_cmdDeleteTagHistory_Click
    
End Sub
