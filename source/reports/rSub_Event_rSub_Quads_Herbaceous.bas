Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5460
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =2445
    Top =2325
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa14d2202a130e540
    End
    RecordSource ="SELECT tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID, StringFromGUID([Quadrat_Data"
        "_ID]) AS Quadrat_Data_txt, tbl_Quadrat_Herbaceous_Data.TSN, tbl_Quadrat_Herbaceo"
        "us_Data.Percent_Cover, [Percent_Cover] & \" %\" AS Perc_Cover_txt, tlu_Plants.La"
        "tin_Name, tbl_Quadrat_Herbaceous_Data.Browse FROM tbl_Quadrat_Herbaceous_Data LE"
        "FT JOIN tlu_Plants ON tbl_Quadrat_Herbaceous_Data.TSN = tlu_Plants.TSN;"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f000000055010000f00000000000000064140000f000000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =120
                    Width =2820
                    Name ="tbxLatinName"
                    ControlSource ="Latin_Name"
                    ConditionalFormat = Begin
                        0x0100000060000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000
                    End

                    ConditionalFormat14 = Begin
                        0x010000000000
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3120
                    Width =780
                    TabIndex =1
                    Name ="Percent_Cover"
                    ControlSource ="Perc_Cover_txt"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0050006500720063005f0043006f007600650072005f007400780074005d00 ,
                        0x3d002200300020002500220000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400160000005b00 ,
                        0x50006500720063005f0043006f007600650072005f007400780074005d003d00 ,
                        0x22003000200025002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4140
                    Width =660
                    TabIndex =2
                    Name ="txtBrowse"
                    ControlSource ="Browse"
                    ConditionalFormat = Begin
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00740078007400420072006f0077007300 ,
                        0x65005d00290000000000
                    End

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400130000004900 ,
                        0x73004e0075006c006c0028005b00740078007400420072006f00770073006500 ,
                        0x5d002900000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =300
                    Width =1680
                    Height =225
                    FontWeight =700
                    BackColor =721136
                    ForeColor =16777215
                    Name ="lblMissingID"
                    Caption ="M I S S I N G  I D"
                    LayoutCachedLeft =300
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =225
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
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
' REPORT:       rSub_Event_rSub_Quads_Herbaceous
' Level:        Application report
' Version:      1.01
'
' Description:  Report related functions & procedures for application
'
' Source/date:  Bonnie Campbell, October 24, 2018
' Revisions:    BLC - 10/24/2018 - 1.00 - initial version
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
' Source/date:  Bonnie Campbell, October 24, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/24/2018 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    'turn on label if missing sapling ID (tbxLatinName)
    'visible IF there is no data (if no latin name = False, returns True & displays)
    lblMissingID.Visible = IIf(Len(tbxLatinName) > 0, False, True)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[rpt_rSub_Event_rSub_Quads_Herbaceous])"
    End Select
    Resume Exit_Handler
End Sub
