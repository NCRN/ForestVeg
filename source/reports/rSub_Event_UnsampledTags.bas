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
    ItemSuffix =43
    Left =435
    Top =3300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf40ffa5a7317e540
    End
    Caption ="rSub_Event_UnsampledTags"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xf0000000f0000000190100000301000000000000c02100009402000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
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
            Height =285
            BackColor =15590879
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =60
                    Width =2955
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="lblTrees"
                    Caption ="T R E E"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =3015
                    LayoutCachedHeight =285
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =4380
                    Width =3405
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="lblSapling"
                    Caption ="S A P L I N G"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4380
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =285
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
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =60
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Subform
                    Left =120
                    Width =4320
                    Height =60
                    Name ="rSub_rSub_Tree_UnsampledTags"
                    SourceObject ="Report.rSub_Event_rSub_Tree_UnsampledTags"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =60
                End
                Begin Subform
                    Left =4320
                    Width =4320
                    Height =60
                    TabIndex =1
                    Name ="rSub_rSub_Sapling_UnsampledTags"
                    SourceObject ="Report.rSub_Event_rSub_Sapling_UnsampledTags"

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =60
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
