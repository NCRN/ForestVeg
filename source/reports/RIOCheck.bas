Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    PageHeader =1
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11535
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =375
    Top =2115
    DatasheetGridlinesColor =14276557
    Filter ="[tbl_Tags].[Tag_Status] = \"Retired (In Office)\""
    OrderBy ="[tbl_Tags].[Tag]"
    RecSrcDt = Begin
        0xc9937ddd895be540
    End
    RecordSource ="tbl_Tags"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0x680100006801000068010000680100000000000060030000b001000000000000 ,
        0x0c0000004800000000000000a20700000100000001000000
    End
    FilterOnLoad =255
    FitToPage =1
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Tag_Status"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1215
            BackColor =14277338
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Width =3675
                    Height =405
                    FontSize =14
                    BorderColor =8355711
                    Name ="lblTitle"
                    Caption ="RIO Check"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedWidth =3675
                    LayoutCachedHeight =405
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9540
                    Top =60
                    Width =1920
                    Height =330
                    ColumnOrder =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPrepareDate"
                    ControlSource ="=Format(Now(),\"mmm dd\"\", \"\"yyyy hh:nn\")"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =9540
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =390
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            Left =8580
                            Top =60
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPrepared"
                            Caption ="Prepared"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =8580
                            LayoutCachedTop =60
                            LayoutCachedWidth =9525
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin Label
                    Left =60
                    Top =420
                    Width =10260
                    Height =615
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblDirections"
                    Caption ="The following tags are currently listed as Retired (In Office) -- aka RIO.  Plac"
                        "e a check in the box IF the tag actually IS in the office. When complete sign, d"
                        "ate, scan && file sheet(s) on the server."
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1035
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10440
                    Top =840
                    Width =960
                    Height =300
                    ColumnOrder =0
                    FontWeight =500
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =2500134
                    Name ="tbxRIOTagCount"
                    ControlSource ="=Nz([TempVars](\"TotalRIOs\"),\"UNKNOWN\")"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =10440
                    LayoutCachedTop =840
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =1140
                    ForeTint =85.0
                    Begin
                        Begin Label
                            Left =8220
                            Top =840
                            Width =2205
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblRIOTagCount"
                            Caption ="Total # of RIO Tags >>"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =8220
                            LayoutCachedTop =840
                            LayoutCachedWidth =10425
                            LayoutCachedHeight =1155
                        End
                    End
                End
                Begin Line
                    Top =1200
                    Width =11520
                    Name ="lnSplitterHdr"
                    GridlineColor =10921638
                    LayoutCachedTop =1200
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1200
                End
            End
        End
        Begin PageHeader
            Height =360
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9375
                    Width =2160
                    Height =330
                    FontSize =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text19"
                    ControlSource ="=Format(Now(),\"mmm dd\"\", \"\"yyyy hh:nn\")"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =9375
                    LayoutCachedWidth =11535
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =8820
                            Width =840
                            Height =300
                            FontSize =10
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label20"
                            Caption ="Prepared"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =8820
                            LayoutCachedWidth =9660
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin Label
                    Width =1125
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="Label18"
                    Caption ="RIO Check"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedWidth =1125
                    LayoutCachedHeight =360
                    ForeTint =100.0
                End
                Begin Line
                    Top =300
                    Width =11520
                    Name ="lnSplitter"
                    GridlineColor =10921638
                    LayoutCachedTop =300
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =300
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="GroupHeader0"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =360
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =720
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxRIOTag"
                    ControlSource ="Tag"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =300
                End
                Begin Rectangle
                    Left =60
                    Top =105
                    Width =180
                    Height =180
                    BorderColor =10921638
                    Name ="rctCheckbox"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =105
                    LayoutCachedWidth =240
                    LayoutCachedHeight =285
                End
            End
        End
        Begin PageFooter
            Height =480
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Left =10860
                    Width =180
                    Height =345
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblOf"
                    Caption ="/"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =10860
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10440
                    Width =360
                    Height =330
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=[Page]"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =10440
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11100
                    Width =420
                    Height =330
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPages"
                    ControlSource ="=[Pages]"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =11100
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =330
                End
                Begin Label
                    Left =180
                    Top =60
                    Width =6090
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCheckedByOn"
                    Caption ="Checked by                                                        Date"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =6270
                    LayoutCachedHeight =375
                End
                Begin Line
                    Left =1320
                    Top =330
                    Width =4320
                    Name ="lnCheckedBy"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =330
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =330
                End
                Begin Line
                    Left =6240
                    Top =300
                    Name ="lnDate"
                    GridlineColor =10921638
                    LayoutCachedLeft =6240
                    LayoutCachedTop =300
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       RIOCheck
' Level:        Application report module
' Version:      1.00
'
' Description:  RIO (retired in office) tag check report related functions & procedures
'
' Source/date:  Bonnie Campbell, October 1, 2019
' Adapted:      -
' Revisions:    BLC - 10/1/2019 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

' ----------------
'  Events
' ----------------

' ---------------------------------
' Sub:          Report_Open
' Description:  report opening actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

Debug.Print Nz([TempVars]("TotalRIOs"), "UNKNOWN")

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[RIOCheck report])"
    End Select
    Resume Exit_Handler
End Sub
