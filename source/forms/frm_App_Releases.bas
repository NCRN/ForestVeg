Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8160
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =5340
    Top =2085
    Right =13500
    Bottom =10110
    DatasheetGridlinesColor =12632256
    OrderBy ="tsys_App_Releases.Release_date DESC"
    RecSrcDt = Begin
        0x16cf15acd1cee240
    End
    RecordSource ="tsys_App_Releases"
    Caption =" Application Releases"
    BeforeInsert ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =8040
            BackColor =9677753
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1560
                    Top =120
                    Width =5340
                    Height =252
                    ColumnWidth =1440
                    FontSize =9
                    Name ="txtRelease_ID"
                    ControlSource ="Release_ID"
                    StatusBarText ="Unique identifier for the release"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1032
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblRelease_ID"
                            Caption ="Release ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6360
                    Top =840
                    Width =1680
                    Height =252
                    ColumnWidth =1140
                    FontSize =9
                    TabIndex =3
                    Name ="txtRelease_date"
                    ControlSource ="Release_date"
                    Format ="Short Date"
                    StatusBarText ="Date of the release"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5160
                            Top =840
                            Width =1185
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="lblRelease_date"
                            Caption ="Release date"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =840
                    Width =1557
                    Height =252
                    ColumnWidth =972
                    FontSize =9
                    TabIndex =2
                    Name ="txtVersion_number"
                    ControlSource ="Version_number"
                    Format ="General Number"
                    StatusBarText ="Version control number"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1452
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblVersion_number"
                            Caption ="Version number"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1200
                    Width =3300
                    Height =252
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =4
                    Name ="txtFile_name"
                    ControlSource ="File_name"
                    StatusBarText ="Filename, used to identify older versions of the database"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =924
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblFile_name"
                            Caption ="File name"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5940
                    Top =1200
                    Width =2106
                    Height =252
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cboRelease_by"
                    ControlSource ="Release_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Last_Name] & (\", \"+[First_Name]) AS FullName FROM tlu_Contacts ORDER B"
                        "Y Last_Name, First_Name; "
                    StatusBarText ="Person who made the release"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4920
                            Top =1200
                            Width =1044
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblRelease_by"
                            Caption ="Release by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =2280
                    Width =6480
                    Height =1020
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =10
                    Name ="txtRelease_notes"
                    ControlSource ="Release_notes"
                    StatusBarText ="Release notes, which may include a summary of revisions"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =1332
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblRelease_notes"
                            Caption ="Release notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =480
                    Width =6480
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"200\""
                    Name ="cboDatabase_title"
                    ControlSource ="Database_title"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tsys_App_Releases.Database_title FROM tsys_App_Releases GROUP BY tsys_App"
                        "_Releases.Database_title ORDER BY tsys_App_Releases.Database_title; "
                    StatusBarText ="Title of the database"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1272
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblDatabase_title"
                            Caption ="Database title"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =120
                    Top =3480
                    Width =7920
                    Height =4380
                    TabIndex =11
                    Name ="subBugs"
                    SourceObject ="Form.fsub_Bug_Reports"
                    LinkChildFields ="Release_ID"
                    LinkMasterFields ="Release_ID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =3240
                            Width =1116
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="lblBugs"
                            Caption ="Known bugs"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1560
                    Width =1800
                    FontSize =9
                    TabIndex =6
                    Name ="txtAuthor_phone"
                    ControlSource ="Author_phone"
                    StatusBarText ="Phone number of application author"
                    FontName ="Arial"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1800
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1215
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label17"
                            Caption ="Author phone"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5160
                    Top =1560
                    Width =2880
                    FontSize =9
                    TabIndex =7
                    Name ="txtAuthor_email"
                    ControlSource ="Author_email"
                    StatusBarText ="Email address of application author"
                    FontName ="Arial"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =1560
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =1800
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3960
                            Top =1560
                            Width =1155
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label18"
                            Caption ="Author email"
                            FontName ="Arial"
                            LayoutCachedLeft =3960
                            LayoutCachedTop =1560
                            LayoutCachedWidth =5115
                            LayoutCachedHeight =1830
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5160
                    Top =1920
                    Width =2880
                    FontSize =9
                    TabIndex =9
                    Name ="txtAuthor_org_name"
                    ControlSource ="Author_org_name"
                    StatusBarText ="Name of organization for author's place of work"
                    FontName ="Arial"

                    LayoutCachedLeft =5160
                    LayoutCachedTop =1920
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =2160
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3600
                            Top =1920
                            Width =1545
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Author org. name"
                            FontName ="Arial"
                            LayoutCachedLeft =3600
                            LayoutCachedTop =1920
                            LayoutCachedWidth =5145
                            LayoutCachedHeight =2190
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1920
                    Width =1800
                    FontSize =9
                    TabIndex =8
                    Name ="cboAuthor_org"
                    ControlSource ="Author_org"
                    RowSourceType ="Table/Query"
                    StatusBarText ="Organization (NPS Unit code) for the author's place of work"
                    FontName ="Arial"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2160
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1440
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label19"
                            Caption ="Author org code"
                            FontName ="Arial"
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
' FORM NAME:    frm_App_Releases
' Description:  Standard module to update release information for the application
' Data source:  tsys_App_Releases
' Data access:  edits, additions, moving between records, deletions
' Pages:        none
' Functions:    none
' References:   CorrectText, GetDataType, fxnGUIDGen, AddAppProperty
' Source/date:  John R. Boetsch?, ?
' Revisions:    Simon Kingston, 2006 - added code to generate string GUID if needed as primary key
'               Simon Kingston, Feb. 2007 - added code to update contact info. if contact picked from tlu_Contacts; added documentation
' =================================

Private Sub cboRelease_by_AfterUpdate()
' Description:  Allows contact info. to be selected from previous entries instead of entering by hand repeatedly
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, 2/26/2007
' Revisions:
Dim strSQL As String
Dim rst As DAO.Recordset
Dim strName As String

On Error GoTo Error_Handler

If Not IsNothing(Me!cboRelease_by) Then
    strName = CorrectText(Me!cboRelease_by)

    strSQL = "SELECT Work_phone, Work_extension, Email_address, Organization FROM tlu_Contacts "
    strSQL = strSQL & "WHERE CorrectText(Last_Name & ', ' & First_Name)=" & Chr(34) & strName & Chr(34) & ";"
    
    Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenForwardOnly)
    'only replace if there's a match to an existing name and phone, email, and organization are currently blank
    If Not (rst.EOF And rst.BOF) And IsNull(Me!txtAuthor_phone) And IsNull(Me!txtAuthor_email) And IsNull(Me!txtAuthor_org_name) Then
        Me!txtAuthor_phone = rst!Work_Phone & (" x" + rst!Work_Extension)
        Me!txtAuthor_email = rst!Email_Address
        Me!txtAuthor_org_name = rst!Organization
        Me!txtAuthor_phone.Requery
        Me!txtAuthor_email.Requery
        Me!txtAuthor_org_name.Requery
    End If
End If

Exit_Handler:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

Error_Handler:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Handler

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Create the GUID primary key value if using a string GUID
    If IsNull(Me!Release_ID) Then
        If GetDataType("tsys_App_Releases", "Release_ID") = dbText Then
            Me!Release_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure


End Sub

Private Sub Form_Close()
Dim varLastReleaseDate As Variant

varLastReleaseDate = DMax("[Release_date]", "tsys_App_Releases")

'update the application title to the latest title recorded
If Not IsNull(varLastReleaseDate) Then
    Call AddAppProperty("AppTitle", dbText, DLookup("[Database_title]", "tsys_App_Releases", "[Release_date]=#" & varLastReleaseDate & "#"))
    Application.RefreshTitleBar
End If
End Sub
