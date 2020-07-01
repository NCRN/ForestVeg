Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    TabularFamily =127
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =10
    ItemSuffix =61
    Left =300
    Top =585
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xd308d50e9a7de540
    End
    RecordSource ="SELECT DISTINCT l.Plot_Name, e.Event_Date, l.Panel, t.Tag, t.Microplot_Number,  "
        "ba.Stems, ba.SumBasalArea_cm2, ba.Equiv_DBH_cm,  sd.Sapling_Status, d.DBH, sd.Sa"
        "pling_Data_ID, sd.Status, t.Tag_Status, sd.Habit, t.Azimuth, t.Distance, t.Azimu"
        "th/t.Distance AS Azi_Dist,  t.Tag_Notes, p.Latin_Name, e.Event_ID, l.Location_ID"
        " FROM ((((((tbl_Locations l  INNER JOIN tbl_Events e ON l.Location_ID = e.Locati"
        "on_ID) INNER JOIN qCalc_Basal_Area_per_Sapling ba ON e.Event_ID = ba.Event_ID)  "
        "INNER JOIN tbl_Tags t ON ba.FirstOfTag_ID = t.Tag_ID)  INNER JOIN tbl_Sapling_Da"
        "ta sd ON t.Tag_ID = sd.Tag_ID)  LEFT JOIN tlu_Plants p ON t.TSN = p.TSN)  LEFT J"
        "OIN tbl_Sapling_DBH d ON sd.Sapling_Data_ID = d.Sapling_Data_ID) WHERE  ba.Equiv"
        "_DBH_cm>=10 AND d.Live = True ORDER BY l.Plot_Name, e.Event_Date, t.Tag;"
    Caption ="rSub_Event_Monster_Saplings"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000c02100002001000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            ControlSource ="Plot_Name"
        End
        Begin BreakLevel
            GroupOn =3
            ControlSource ="Event_Date"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Microplot_Number"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Tag"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1080
            Name ="ReportHeader"
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    BackStyle =1
                    Top =720
                    Width =8640
                    Height =300
                    Name ="boxHdrDBH"
                    LayoutCachedTop =720
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =1020
                End
                Begin Rectangle
                    BackStyle =1
                    Width =8640
                    Height =360
                    BackColor =13434879
                    Name ="boxHdrMP"
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =360
                End
                Begin Rectangle
                    BackStyle =1
                    Top =360
                    Width =8640
                    Height =360
                    BackColor =16776935
                    Name ="boxHdrTag"
                    GridlineWidthLeft =0
                    GridlineWidthRight =0
                    LayoutCachedTop =360
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =720
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Top =420
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrTag"
                    Caption ="Tag"
                    FontName ="Arial"
                    LayoutCachedTop =420
                    LayoutCachedWidth =840
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =900
                    Top =420
                    Width =1440
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrClass"
                    Caption ="Tag Status"
                    FontName ="Arial"
                    LayoutCachedLeft =900
                    LayoutCachedTop =420
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =900
                    Top =60
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrAziDist"
                    Caption ="Azi/Dist"
                    FontName ="Arial"
                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Top =60
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrMP"
                    Caption ="MP"
                    FontName ="Arial"
                    LayoutCachedTop =60
                    LayoutCachedWidth =840
                    LayoutCachedHeight =285
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =4020
                    Top =720
                    Width =4440
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrSaplingStatus"
                    Caption ="Sapling Status"
                    FontName ="Arial"
                    LayoutCachedLeft =4020
                    LayoutCachedTop =720
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =945
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =6960
                    Top =420
                    Width =1485
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrTagNotes"
                    Caption ="Tag Notes"
                    FontName ="Arial"
                    LayoutCachedLeft =6960
                    LayoutCachedTop =420
                    LayoutCachedWidth =8445
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =2460
                    Top =420
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblEquivDBH"
                    Caption ="Equiv DBH"
                    FontName ="Arial"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =420
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =3420
                    Top =420
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblStems"
                    Caption ="Stems"
                    FontName ="Arial"
                    LayoutCachedLeft =3420
                    LayoutCachedTop =420
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =4320
                    Top =420
                    Width =2400
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblLatinName"
                    Caption ="Latin Name"
                    FontName ="Arial"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =420
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =2460
                    Top =720
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblDBH"
                    Caption ="DBH"
                    FontName ="Arial"
                    LayoutCachedLeft =2460
                    LayoutCachedTop =720
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =945
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
            End
        End
        Begin PageHeader
            Height =15
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Name ="Line14"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            BreakLevel =2
            BackColor =13434879
            Name ="GroupHeader2"
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Width =780
                    Height =288
                    FontSize =9
                    FontWeight =600
                    BackColor =0
                    ForeColor =4210752
                    Name ="tbxMP"
                    ControlSource ="Microplot_Number"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =840
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1140
                    Width =780
                    Height =288
                    TabIndex =1
                    ForeColor =4210752
                    Name ="tbxAziDist"
                    ControlSource ="Azi_Dist"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            BreakLevel =3
            BackColor =16776935
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Width =780
                    Height =288
                    FontSize =9
                    FontWeight =500
                    ForeColor =4210752
                    Name ="tbxTag"
                    ControlSource ="Tag"

                    LayoutCachedWidth =780
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =900
                    Height =288
                    TabIndex =1
                    ForeColor =4210752
                    Name ="tbxClass"
                    ControlSource ="Tag_Status"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6720
                    Width =1740
                    Height =288
                    TabIndex =2
                    ForeColor =4210752
                    Name ="tbxTagNotes"
                    ControlSource ="Tag_Notes"

                    LayoutCachedLeft =6720
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    FontItalic = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Width =2220
                    Height =288
                    TabIndex =3
                    ForeColor =4210752
                    Name ="tbxLatinName"
                    ControlSource ="Latin_Name"

                    LayoutCachedLeft =4260
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Width =780
                    Height =288
                    TabIndex =4
                    ForeColor =4210752
                    Name ="tbxStems"
                    ControlSource ="Stems"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Width =780
                    Height =288
                    TabIndex =5
                    ForeColor =4210752
                    Name ="tbxEquivDBH"
                    ControlSource ="Equiv_DBH_cm"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4020
                    Width =4440
                    Height =288
                    ForeColor =4210752
                    Name ="tbxSaplingStatus"
                    ControlSource ="Sapling_Status"

                    LayoutCachedLeft =4020
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2520
                    Width =780
                    Height =288
                    TabIndex =1
                    ForeColor =4210752
                    Name ="tbxDBH"
                    ControlSource ="DBH"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =288
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
            End
        End
        Begin PageFooter
            Height =15
            Name ="PageFooterSection"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =0
                    BorderColor =12632256
                    Name ="Line15"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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
' REPORT:       rSub_Event_Monster_Saplings
' Level:        Application report
' Version:      1.01
'
' Description:  Report related functions & procedures for application
'
' Source/date:  Bonnie Campbell, April 5, 2018
' Revisions:    BLC - 4/5/2018 - 1.00 - initial version
'               BLC - 4/12/2018 - 1.01 - added NoData event
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  report open actions
' Assumptions:  -
' Parameters:   Cancel - whether open action(s) should be cancelled (boolean)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 12, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/12/2018 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[rpt_Event_Monster_Saplings])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Detail_Format
' Description:  report format actions
' Assumptions:  -
' Parameters:   Cancel - whether format action should be cancelled (boolean)
'               FormatCount - number of times a section (in this case the detail section)
'                             is formatted (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 12, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/12/2018 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    'show/hide label
    'Me.lblNoData.Visible = Not Me.Report.HasData
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[rpt_rSub_Event_UnsampledTags])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Report_NoData
' Description:  report no data actions
' Assumptions:  -
' Parameters:   Cancel - whether no data action(s) should be cancelled (boolean)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 12, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/12/2018 - initial version
' ---------------------------------
Private Sub Report_NoData(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.Recordset.RecordCount = 0 Then
        'lblNoData.Visible = False
    Else
        'lblNoData.Visible = False
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_NoData[rpt_Event_Monster_Saplings])"
    End Select
    Resume Exit_Handler
End Sub
