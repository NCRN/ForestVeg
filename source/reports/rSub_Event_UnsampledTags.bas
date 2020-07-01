Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =10
    ItemSuffix =46
    Left =810
    Top =750
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf9628c6e967de540
    End
    RecordSource ="SELECT * FROM ( SELECT t.Tag_ID, t.Tag, t.Tag_Status, IIf(IsNull([azimuth]),\"\""
        ",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, t.Microplot_Number AS MP"
        ", t.Location_ID, LEN(t.Tag)*10/LEN(t.TAG) AS RecordOrder FROM (tbl_Tags AS t  LE"
        "FT JOIN qry_Status_Sapling_Current_Event sce ON t.Tag_ID = sce.Tag_ID)  LEFT JOI"
        "N qry_Status_Tree_Current_Event tce ON t.Tag_ID = tce.Tag_ID WHERE  sce.Event_ID"
        " Is Null AND tce.Event_ID Is Null AND t.Tag_Status NOT IN ('Retired (In Office)'"
        ", 'Inactive (In Field)', 'Inactive (Lost)') GROUP BY LEN(t.Tag)*1/LEN(t.Tag), t."
        "Tag_Status, t.Tag, t.Tag_ID, IIf(IsNull([azimuth]),\"\",[Azimuth] & \" / \" & [d"
        "istance] & \"m\"),  t.Microplot_Number, t.Location_ID ORDER BY t.Tag_Status, t.T"
        "ag ) grp1  UNION ALL  SELECT * FROM ( SELECT t.Tag_ID, t.Tag, t.Tag_Status, IIf("
        "IsNull([azimuth]),\"\",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, t."
        "Microplot_Number AS MP, t.Location_ID,  LEN(t.Tag)*10^5/LEN(t.Tag)  AS RecordOrd"
        "er FROM (tbl_Tags AS t  LEFT JOIN qry_Status_Sapling_Current_Event sce ON t.Tag_"
        "ID = sce.Tag_ID)  LEFT JOIN qry_Status_Tree_Current_Event tce ON t.Tag_ID = tce."
        "Tag_ID WHERE  sce.Event_ID Is Null AND tce.Event_ID Is Null AND t.Tag_Status IN "
        "('Retired (In Office)', 'Inactive (In Field)', 'Inactive (Lost)') GROUP BY LEN(t"
        ".Tag)*10^5/LEN(t.Tag), t.Tag_Status, t.Tag, t.Tag_ID, IIf(IsNull([azimuth]),\"\""
        ",[Azimuth] & \" / \" & [distance] & \"m\"),  t.Microplot_Number, t.Location_ID O"
        "RDER BY t.Tag_Status, t.Tag ) grp2 ORDER BY RecordOrder, t.Tag_Status, t.Tag;"
    Caption ="rSub_Event_UnsampledTags"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000e01000003c00000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
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
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            NewRowOrCol =1
            Height =0
            BackColor =15590879
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =220
            BackColor =15590879
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =840
                    Height =220
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrTag"
                    Caption ="Tag"
                    FontName ="Arial"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedWidth =840
                    LayoutCachedHeight =220
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =900
                    Width =1440
                    Height =220
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrTagStatus"
                    Caption ="Tag Status"
                    FontName ="Arial"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =900
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =220
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =2520
                    Width =840
                    Height =220
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrAziDist"
                    Caption ="Azi/Dist"
                    FontName ="Arial"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =220
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =3360
                    Width =840
                    Height =220
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrMP"
                    Caption ="MP"
                    FontName ="Arial"
                    TopPadding =0
                    BottomPadding =0
                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =220
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            NewRowOrCol =2
            Height =60
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Width =780
                    Height =0
                    ForeColor =4210752
                    Name ="tbxTag"
                    ControlSource ="Tag"

                    LayoutCachedWidth =780
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =840
                    Height =0
                    TabIndex =1
                    ForeColor =4210752
                    Name ="tbxTagStatus"
                    ControlSource ="Tag_Status"

                    LayoutCachedLeft =840
                    LayoutCachedWidth =2280
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2400
                    Width =1008
                    Height =0
                    TabIndex =2
                    ForeColor =4210752
                    Name ="tbxAziDist"
                    ControlSource ="Azi_Dist"

                    LayoutCachedLeft =2400
                    LayoutCachedWidth =3408
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3420
                    Width =780
                    Height =0
                    TabIndex =3
                    ForeColor =4210752
                    Name ="tbxMP"
                    ControlSource ="MP"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4200
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =255
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
' REPORT:       rSub_Event_UnsampledTags
' Level:        Application report
' Version:      1.00
'
' Description:  Report related functions & procedures for application
'
' Source/date:  Bonnie Campbell, April 5, 2018
' Revisions:    BLC - 4/5/2018 - 1.00 - initial version
' =================================

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
' Source/date:  Bonnie Campbell, April 5, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/5/2018 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    
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
