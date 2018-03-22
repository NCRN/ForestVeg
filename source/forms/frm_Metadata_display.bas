Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =2
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =6
    Left =6090
    Top =1395
    Right =16155
    Bottom =3465
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x73558bf3552be340
    End
    RecordSource ="tbl_Db_Meta"
    Caption ="Database Purpose and Metadata"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin Section
            Height =2100
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =420
                    Width =9840
                    Height =768
                    Name ="txtDb_Desc"
                    ControlSource ="Db_Desc"
                    StatusBarText ="M. Description of the database purpose (Db_Desc)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =60
                            Width =9840
                            Height =240
                            Name ="Label0"
                            Caption ="Description of the purpose of the database:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1320
                    Width =1920
                    Height =240
                    Name ="Label2"
                    Caption ="Local Metadata File"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2100
                    Top =1320
                    Width =7860
                    Height =210
                    Name ="lblLocalMetadataFile"
                    Caption ="None"
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1680
                    Width =1920
                    Height =240
                    Name ="Label4"
                    Caption ="NPS Data Store Metadata"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2100
                    Top =1680
                    Width =7860
                    Height =210
                    Name ="lblNPSDataStoreMetadata"
                    Caption ="None"
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
Option Explicit

Private Sub Form_Load()
Dim varLocalMetadataFileName As Variant
Dim varNPSDataStoreMetadataURL As Variant
Dim strCaptionLocal As String

varLocalMetadataFileName = fxnGetLocalMetadataFileName

If fxnLocalMetadataExists(varLocalMetadataFileName) Then
    strCaptionLocal = varLocalMetadataFileName
    Me!lblLocalMetadataFile.HyperlinkAddress = varLocalMetadataFileName
    HyperlinkFormat True, Me!lblLocalMetadataFile
Else
    If Not IsNull(varLocalMetadataFileName) Then
        strCaptionLocal = "Unable to find file " & varLocalMetadataFileName
    Else
        strCaptionLocal = "None"
    End If
End If

Me!lblLocalMetadataFile.Caption = strCaptionLocal

varNPSDataStoreMetadataURL = fxnGetNPSDataStoreMetadataLink

If fxnNPSDataStoreMetadataExists(varNPSDataStoreMetadataURL) Then
    Me!lblNPSDataStoreMetadata.Caption = varNPSDataStoreMetadataURL
    Me!lblNPSDataStoreMetadata.HyperlinkAddress = varNPSDataStoreMetadataURL
    HyperlinkFormat True, Me!lblNPSDataStoreMetadata
End If

End Sub
