Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =2820
    Top =5295
    Right =9255
    Bottom =8760
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x2e1f8472d703e440
    End
    Caption ="Utilities"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin Section
            Height =3480
            BackColor =15921906
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Width =6060
                    Height =480
                    FontSize =18
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUtilities_Header"
                    Caption ="Utilities and Configuration Tools"
                    GridlineColor =10921638
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =480
                    ThemeFontIndex =-1
                    BackThemeColorIndex =0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1680
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdData_QA"
                    Caption ="QA/QC"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the QA/QC Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Data_QA"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2040
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =240
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdAppend"
                    Caption =" Append Data"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the Append Data Switchboard ti Import Field Data"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Append_Switchboard"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =2040
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =240
                    Top =660
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    Name ="cmdRelink_Tables"
                    Caption =" Relink Tables"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Reset the link to the backend database"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Connect_Tables"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =660
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =1920
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4560
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdClose_Utilities"
                    Caption ="Close"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =10798077
                    HoverThemeColorIndex =5
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =1680
                    Top =660
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdBackup_BE"
                    Caption ="Create Backup"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Create a Backup of the Backend Database"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =660
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1920
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3120
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdLookups"
                    Caption ="Lookups"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the QA/QC Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Lookups"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cmdLookups\" Event=\"OnClick\" xmlns=\"http://schemas.micros"
                                "oft.com/office/accessservices/2009/11/application\" xmlns:a=\"http://schemas.mic"
                                "rosoft.com/office/accessservices"
                        End
                        Begin
                            Comment ="_AXL:/2009/11/forms\"><Statements><Action Name=\"OpenForm\"><Argument Name=\"For"
                                "mName\">frm_Lookups</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3120
                    LayoutCachedTop =2040
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
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

Private Sub cmdBackup_BE_Click()
        fxnMakeBackup
End Sub
