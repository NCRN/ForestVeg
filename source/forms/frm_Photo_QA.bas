Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7260
    DatasheetFontHeight =10
    ItemSuffix =40
    Left =6375
    Top =2820
    Right =13365
    Bottom =7560
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7bb4fcf19148e440
    End
    RecordSource ="tbl_QA_Photos"
    Caption ="Append Utilities"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =5205
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =240
                    Top =1620
                    Width =6840
                    Height =2400
                    FontSize =14
                    Name ="Label39"
                    Caption ="No Errors Detected"
                    LayoutCachedLeft =240
                    LayoutCachedTop =1620
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =4020
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =480
                    Width =7125
                    Height =210
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Please document your assessment of these photos"
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =7185
                    LayoutCachedHeight =690
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =-15
                    Width =7275
                    Height =465
                    FontSize =18
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label19"
                    Caption ="Photo QA Report"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =-15
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =465
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5820
                    Top =4140
                    Width =1259
                    Height =734
                    FontWeight =700
                    ForeColor =0
                    Name ="cmd_close_form"
                    Caption ="Submit Report"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5820
                    LayoutCachedTop =4140
                    LayoutCachedWidth =7079
                    LayoutCachedHeight =4874
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =10798077
                    HoverThemeColorIndex =5
                    HoverTint =40.0
                    PressedColor =0
                    PressedThemeColorIndex =0
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4320
                    Top =4200
                    Height =345
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtEvent_Date1"
                    ControlSource ="Event_Date1"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedTop =4200
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4545
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =3300
                            Top =4200
                            Width =1020
                            Height =240
                            Name ="Label33"
                            Caption ="Event_Date1"
                            LayoutCachedLeft =3300
                            LayoutCachedTop =4200
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =4440
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4320
                    Top =4620
                    Height =345
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtEvent_Date2"
                    ControlSource ="Event_Date2"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedTop =4620
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =4965
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =3300
                            Top =4620
                            Width =1020
                            Height =240
                            Name ="Label34"
                            Caption ="Event_Date2"
                            LayoutCachedLeft =3300
                            LayoutCachedTop =4620
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =4860
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =240
                    Top =930
                    TabIndex =3
                    BorderColor =10921638
                    Name ="chkError_Detected"
                    ControlSource ="Error_Detected"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =930
                    LayoutCachedWidth =500
                    LayoutCachedHeight =1170
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =470
                            Top =900
                            Width =1185
                            Height =240
                            Name ="Label35"
                            Caption ="Error_Detected"
                            LayoutCachedLeft =470
                            LayoutCachedTop =900
                            LayoutCachedWidth =1655
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =240
                    Top =1560
                    Width =6840
                    Height =2454
                    ColumnWidth =9960
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtError_Description"
                    ControlSource ="Error_Description"
                    StatusBarText ="Description of the error"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =4014
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1260
                            Width =2400
                            Height =240
                            Name ="Label36"
                            Caption ="Error Description"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1260
                            LayoutCachedWidth =2640
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2040
                    Top =4140
                    Height =345
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtAD_Name"
                    ControlSource ="AD_Name"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =4140
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =4485
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =240
                            Top =4140
                            Width =780
                            Height =240
                            Name ="Label37"
                            Caption ="AD_Name"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4140
                            LayoutCachedWidth =1020
                            LayoutCachedHeight =4380
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2040
                    Top =4560
                    Height =345
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Location_ID)"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =4560
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =4905
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =240
                            Top =4560
                            Width =930
                            Height =240
                            Name ="Label38"
                            Caption ="Location_ID"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4560
                            LayoutCachedWidth =1170
                            LayoutCachedHeight =4800
                        End
                    End
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


Private Sub chkError_Detected_AfterUpdate()
    If Me.chkError_Detected = True Then
        Me.txtError_Description.Visible = True
    Else
      Me.txtError_Description.Visible = False
    End If
End Sub

Private Sub cmd_Close_Form_Click()
On Error GoTo Err_cmd_close_form_Click
    Forms("frm_Photos").txtMatch_Votes.Requery
    Forms("frm_Photos").txtNoMatch_Votes.Requery
    DoCmd.Close

Exit_cmd_close_form_Click:
    Exit Sub
Err_cmd_close_form_Click:
    MsgBox Err.Description
    Resume Exit_cmd_close_form_Click
End Sub
