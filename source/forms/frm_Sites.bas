Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9960
    DatasheetFontHeight =9
    ItemSuffix =11
    Left =4440
    Top =2175
    Right =14685
    Bottom =5775
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x125596ccd408e340
    End
    RecordSource ="tbl_Sites"
    Caption =" Sites"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
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
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
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
        Begin ComboBox
            SpecialEffect =2
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
        Begin Section
            Height =3606
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1446
                    Top =120
                    Width =8388
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtSite_ID"
                    ControlSource ="Site_ID"
                    StatusBarText ="M. Site identifier (Site_ID)"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1200
                            Height =225
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblLocation_ID"
                            Caption ="Site ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =1440
                    Top =480
                    Width =1020
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboUnit_Code"
                    ControlSource ="Unit_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"Unit"
                        " Code\" ORDER BY Enum_Code; "
                    ColumnWidths ="720;5040"
                    StatusBarText ="NPS Unit code"
                    DefaultValue ="=[Forms]![frm_Switchboard]![cPark]"
                    FontName ="Arial"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =1230
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblUnitCode"
                            Caption ="NPS Unit"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =2520
                    Width =3540
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtGIS_Location_ID"
                    ControlSource ="GIS_Location_ID"
                    StatusBarText ="MA. Link to GIS feature, equivalent to NPS_Location_ID (GIS_Loc_ID)"
                    FontName ="MS Sans Serif"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2520
                            Width =1215
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label38"
                            Caption ="GIS Location ID"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =480
                    Width =3300
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtSite_Name"
                    ControlSource ="Site_Name"
                    StatusBarText ="M. Unique name or code for a site (Site_Name)"
                    FontName ="MS Sans Serif"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2640
                            Top =480
                            Width =1155
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label49"
                            Caption ="Site Name"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =1680
                    Width =8400
                    Height =603
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtSite_Notes"
                    ControlSource ="Site_Notes"
                    StatusBarText ="MA. General notes on the site (Site_Notes)"
                    FontName ="MS Sans Serif"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1680
                            Width =1155
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label51"
                            Caption ="Site Notes"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Line
                    OverlapFlags =85
                    Left =120
                    Top =840
                    Width =9720
                    Name ="Line1"
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
                Begin Line
                    OverlapFlags =85
                    Left =120
                    Top =2400
                    Width =9720
                    Name ="Line3"
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =8400
                    Top =3120
                    Height =300
                    FontWeight =700
                    TabIndex =7
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                End
                Begin Line
                    OverlapFlags =85
                    Left =120
                    Top =2880
                    Width =9720
                    Name ="Line6"
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =6480
                    Top =3120
                    Height =300
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdAddSite"
                    Caption ="Add New Site"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1440
                    Top =960
                    Width =8400
                    Height =603
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtSite_Desc"
                    ControlSource ="Site_Desc"
                    StatusBarText ="M. Description for a site (Site_Desc)"
                    FontName ="MS Sans Serif"
                    Tag ="<data>"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =960
                            Width =1185
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label10"
                            Caption ="Site Description"
                            FontName ="MS Sans Serif"
                        End
                    End
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

' =================================
' Description:  Sites entry form
' Data source:  tbl_Sites
' Data access:  edit, add, delete
' Pages:        none
' Functions:    none
' References:   fxnGUIDGen
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cmdAddSite_Click()
DoCmd.GoToRecord acActiveDataObject, Me.Name, acNewRec
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'if the form checks out and a primary key is needed, generate the pk
If fxnFormCheck(Me) Then
    If IsNull(Me!txtSite_ID) Then
        If GetDataType("tbl_Sites", "Site_ID") = dbText Then
            Me!txtSite_ID = fxnGUIDGen
        End If
    End If
Else
    Cancel = True
End If
End Sub

Private Sub Form_Close()
Dim strFormName As String

On Error Resume Next

'requery any controls that need to reflect new site values
strFormName = "frm_Locations"
If IsLoaded(strFormName) Then
    Forms(strFormName)!cboSite_ID.Requery
End If
End Sub

Private Sub Form_Current()
'generate the primary key if we are using string GUIDs
If Me.NewRecord Then
    If GetDataType("tbl_Sites", "Site_ID") = dbText Then
        Me!txtSite_ID = fxnGUIDGen
    End If
End If
End Sub
