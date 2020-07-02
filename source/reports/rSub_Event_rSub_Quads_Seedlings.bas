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
    Width =5220
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =495
    Top =3810
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdcf87604a130e540
    End
    RecordSource ="SELECT tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID, StringFromGUID([Quadrat_Data_"
        "ID]) AS Quadrat_Data_txt, tbl_Quadrat_Seedlings_Data.TSN, tlu_Plants.Family, tlu"
        "_Plants.Genus, tlu_Plants.Species, tbl_Quadrat_Seedlings_Data.Height, [genus] & "
        "\" \" & [species] AS SciName, [Height] & \" cm\" AS Height_txt, tlu_Plants.Latin"
        "_Name, [Browsable] & \"/\" & [Browsed] AS Browse FROM tbl_Quadrat_Seedlings_Data"
        " LEFT JOIN tlu_Plants ON tbl_Quadrat_Seedlings_Data.TSN = tlu_Plants.TSN;"
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
            Height =240
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =120
                    Width =2520
                    Name ="Genus"
                    ControlSource ="Latin_Name"
                    ConditionalFormat = Begin
                        0x0100000060000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x010000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3420
                    Width =720
                    TabIndex =1
                    Name ="txt_Height"
                    ControlSource ="Height_txt"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2640
                    Width =780
                    TabIndex =2
                    Name ="txtBrowse"
                    ControlSource ="Browse"

                    LayoutCachedLeft =2640
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =240
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =240
                    Width =1680
                    Height =225
                    FontWeight =700
                    BackColor =721136
                    ForeColor =16777215
                    Name ="lblMissingID"
                    Caption ="M I S S I N G  I D"
                    LayoutCachedLeft =240
                    LayoutCachedWidth =1920
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

    'turn on label if missing seedling ID (Genus)
    'visible IF there is no data (if no genus = False, returns True & displays)
    lblMissingID.visible = IIf(Len(Genus) > 0, False, True)
    
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
