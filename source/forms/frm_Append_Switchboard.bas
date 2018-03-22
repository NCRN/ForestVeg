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
    ItemSuffix =32
    Left =7860
    Top =3255
    Right =15120
    Bottom =8280
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf1f79facd5fde240
    End
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =5040
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =480
                    Width =7080
                    Height =210
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Utilities for appending data collected in the field to the master database"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Width =7260
                    Height =420
                    FontSize =18
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label19"
                    Caption ="Append Utilities"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =360
                    Top =4200
                    Width =839
                    Height =734
                    FontWeight =700
                    ForeColor =0
                    Name ="cmd_close_form"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =360
                    LayoutCachedTop =4200
                    LayoutCachedWidth =1199
                    LayoutCachedHeight =4934
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =360
                    Top =1680
                    Width =840
                    Height =735
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="cmd_Select"
                    Caption ="Select"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =360
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
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
                Begin Label
                    OverlapFlags =87
                    Left =1260
                    Top =1920
                    Width =4680
                    Height =210
                    Name ="lblSelect"
                    Caption ="Select the tables to be imported"
                    LayoutCachedLeft =1260
                    LayoutCachedTop =1920
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2130
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =360
                    Top =2520
                    Width =840
                    Height =735
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="cmd_Append"
                    Caption ="Append"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =360
                    LayoutCachedTop =2520
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =3255
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
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
                Begin Label
                    OverlapFlags =87
                    Left =1260
                    Top =2760
                    Width =4680
                    Height =210
                    Name ="lblAppend"
                    Caption ="Append the selected tables to the master database"
                    LayoutCachedLeft =1260
                    LayoutCachedTop =2760
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2970
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =360
                    Top =3360
                    Width =840
                    Height =735
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="cmd_Delete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =360
                    LayoutCachedTop =3360
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =4095
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
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
                Begin Label
                    OverlapFlags =87
                    Left =1260
                    Top =3600
                    Width =4680
                    Height =210
                    Name ="lblDelete"
                    Caption ="Delete the temporary import tables"
                    LayoutCachedLeft =1260
                    LayoutCachedTop =3600
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =3810
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =360
                    Top =840
                    Width =840
                    Height =735
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =0
                    Name ="cmd_Backup"
                    Caption ="Backup"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =360
                    LayoutCachedTop =840
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =1575
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
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
                Begin Label
                    OverlapFlags =87
                    Left =1260
                    Top =1080
                    Width =4680
                    Height =210
                    Name ="lblBackup"
                    Caption ="Make a dated backup copy of the database backend (optional)"
                    LayoutCachedLeft =1260
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1290
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

Private Sub cmd_Backup_Click()
    fxnMakeBackup
End Sub

Private Sub cmd_Select_Click()
On Error GoTo Err_cmd_Select_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Select_Import_Tables"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmd_Select_Click:
    Exit Sub
Err_cmd_Select_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Select_Click
End Sub

Private Sub cmd_Append_Click()
On Error GoTo Err_cmd_Append_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Append_Data"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmd_Append_Click:
    Exit Sub
Err_cmd_Append_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Append_Click
End Sub

Private Sub cmd_Delete_Click()
On Error GoTo Err_cmd_Delete_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Delete_Tables"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmd_Delete_Click:
    Exit Sub
Err_cmd_Delete_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Delete_Click
End Sub

Private Sub cmd_Close_Form_Click()
On Error GoTo Err_cmd_close_form_Click

    DoCmd.Close

Exit_cmd_close_form_Click:
    Exit Sub
Err_cmd_close_form_Click:
    MsgBox Err.Description
    Resume Exit_cmd_close_form_Click
End Sub
