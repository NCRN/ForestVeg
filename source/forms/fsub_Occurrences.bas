Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8760
    DatasheetFontHeight =11
    ItemSuffix =29
    Left =11895
    Top =5295
    Right =21360
    Bottom =9960
    DatasheetGridlinesColor =15062992
    OrderBy ="[qSum_Active_Trees_Shrubs_Herbs_Vines_by_Event].[Date], [qSum_Active_Trees_Shrub"
        "s_Herbs_Vines_by_Event].[Plot_Name], [qSum_Active_Trees_Shrubs_Herbs_Vines_by_Ev"
        "ent].[Habit_Class]"
    RecSrcDt = Begin
        0xa71c6b6f13c8e340
    End
    RecordSource ="qSum_Active_Trees_Shrubs_Herbs_Vines_by_Event"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin Section
            Height =4560
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =177
                    IMESentenceMode =3
                    Left =3900
                    Top =60
                    Height =315
                    ColumnOrder =0
                    FontSize =10
                    Name ="txtTSN"
                    ControlSource ="TSN"
                    StatusBarText ="ITIS TSN"
                    FontName ="Arial"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =177
                            TextAlign =3
                            Left =2100
                            Top =60
                            Width =1740
                            Height =315
                            FontSize =10
                            Name ="Label1"
                            Caption ="TSN:"
                            FontName ="Arial"
                            LayoutCachedLeft =2100
                            LayoutCachedTop =60
                            LayoutCachedWidth =3840
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =3900
                    Top =480
                    Height =315
                    ColumnWidth =1395
                    ColumnOrder =1
                    TabIndex =1
                    ForeColor =8210719
                    Name ="txtPlot_Name"
                    ControlSource ="Plot_Name"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =480
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =795
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =480
                            Width =1740
                            Height =315
                            Name ="Label14"
                            Caption ="Plot"
                            LayoutCachedLeft =2100
                            LayoutCachedTop =480
                            LayoutCachedWidth =3840
                            LayoutCachedHeight =795
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =840
                    Height =315
                    ColumnWidth =1575
                    ColumnOrder =3
                    TabIndex =2
                    Name ="txtEvent_Date"
                    ControlSource ="Date"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =840
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1155
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =840
                            Width =1740
                            Height =315
                            Name ="Label15"
                            Caption ="Date"
                            LayoutCachedLeft =2100
                            LayoutCachedTop =840
                            LayoutCachedWidth =3840
                            LayoutCachedHeight =1155
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =1200
                    Height =315
                    ColumnWidth =1890
                    ColumnOrder =5
                    TabIndex =3
                    Name ="txtHabit_Class"
                    ControlSource ="Habit_Class"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1515
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =1200
                            Width =1740
                            Height =315
                            Name ="Label17"
                            Caption ="Habit/Class"
                            LayoutCachedLeft =2100
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3840
                            LayoutCachedHeight =1515
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =1560
                    Height =315
                    ColumnWidth =1935
                    ColumnOrder =6
                    TabIndex =4
                    Name ="txtOccurence_Count"
                    ControlSource ="Occurence_Count"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1875
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =1560
                            Width =1740
                            Height =315
                            Name ="Label18"
                            Caption ="Occurences"
                            LayoutCachedLeft =2100
                            LayoutCachedTop =1560
                            LayoutCachedWidth =3840
                            LayoutCachedHeight =1875
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =420
                    Height =315
                    ColumnWidth =825
                    ColumnOrder =2
                    TabIndex =5
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =420
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =735
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7020
                    Top =60
                    Height =315
                    ColumnWidth =945
                    TabIndex =6
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"

                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =375
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Plot_Name_Click()
On Error GoTo Err_Plot_Name_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Locations"
    stLinkCriteria = "[Location_ID]=" & StringFromGUID(Me!txt_Location_ID)
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Plot_Name_Click:
    Exit Sub

Err_Plot_Name_Click:
    MsgBox Err.Description
    Resume Exit_Plot_Name_Click
End Sub

Private Sub txtEvent_Date_Click()
On Error GoTo Err_txtEvent_Date_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Events"
    stLinkCriteria = "[Event_ID]=" & StringFromGUID(Me!txtEvent_ID)
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_txtEvent_Date_Click:
    Exit Sub

Err_txtEvent_Date_Click:
    MsgBox Err.Description
    Resume Exit_txtEvent_Date_Click
End Sub

Private Sub txtPlot_Name_Click()
On Error GoTo Err_txtPlot_Name_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Locations"
    stLinkCriteria = "[Location_ID]=" & StringFromGUID(Me!txtLocation_ID)
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_txtPlot_Name_Click:
    Exit Sub

Err_txtPlot_Name_Click:
    MsgBox Err.Description
    Resume Exit_txtPlot_Name_Click
End Sub
