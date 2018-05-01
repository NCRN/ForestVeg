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
    Left =1320
    Top =2190
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8e5b0749ae1ae540
    End
    RecordSource ="SELECT ev.Event_Date, l.Plot_Name, ev.Event_ID, e.Enum_Code FROM tbl_Events ev, "
        "tlu_Enumerations e, tbl_Locations l WHERE e.Enum_Group = \"Quadrat Number\" AND "
        "l.Location_ID = ev.Location_ID AND e.Enum_Code NOT IN (SELECT q.Quadrat_Number F"
        "ROM tbl_Quadrat_Data q  WHERE q.Event_ID = ev.Event_ID) ORDER BY ev.Event_ID, e."
        "Enum_Code;"
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
            Height =225
            BackColor =15590879
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Width =1440
                    Height =225
                    FontSize =8
                    FontWeight =800
                    ForeColor =5855577
                    Name ="lblHdrQuadrat"
                    Caption ="Quadrat"
                    FontName ="Arial"
                    LayoutCachedWidth =1440
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
            Height =14
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Height =0
                    ForeColor =4210752
                    Name ="tbxQuadrat"
                    ControlSource ="Enum_Code"

                    LayoutCachedWidth =1440
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
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
' REPORT:       rSub_Event_UnsampledQuadrats
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
            "Error encountered (#" & Err.Number & " - Detail_Format[rpt_rSub_Event_UnsampledQuadrats])"
    End Select
    Resume Exit_Handler
End Sub
