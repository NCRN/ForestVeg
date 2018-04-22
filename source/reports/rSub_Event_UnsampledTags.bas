Version =20
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
    Width =8640
    DatasheetFontHeight =10
    ItemSuffix =46
    Left =6930
    Top =4755
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x51b40f584e19e540
    End
    RecordSource ="SELECT t.Tag_ID, t.Tag, t.Tag_Status,  IIf(IsNull([azimuth]),\"\",[Azimuth] & \""
        " / \" & [distance] & \"m\") AS Azi_Dist,  t.Microplot_Number AS MP, t.Location_I"
        "D FROM (tbl_Tags t  LEFT JOIN qry_Status_Sapling_Current_Event ON t.Tag_ID = qry"
        "_Status_Sapling_Current_Event.Tag_ID)  LEFT JOIN qry_Status_Tree_Current_Event O"
        "N t.Tag_ID = qry_Status_Tree_Current_Event.Tag_ID WHERE ( ((qry_Status_Sapling_C"
        "urrent_Event.Event_ID) Is Null)  AND ((qry_Status_Tree_Current_Event.Event_ID) I"
        "s Null)) ORDER BY t.Tag_Status, t.Tag;"
    Caption ="rSub_Event_UnsampledTags"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000c02100000f00000001000000 ,
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
            Height =240
            BackColor =15590879
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrTag"
                    Caption ="Tag"
                    FontName ="Arial"
                    LayoutCachedWidth =840
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =900
                    Width =1440
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrTagStatus"
                    Caption ="Tag Status"
                    FontName ="Arial"
                    LayoutCachedLeft =900
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =2520
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrAziDist"
                    Caption ="Azi/Dist"
                    FontName ="Arial"
                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =3360
                    Width =840
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrMP"
                    Caption ="MP"
                    FontName ="Arial"
                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =15
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Name ="Line14"
                End
                Begin TextBox
                    CanGrow = NotDefault
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
