Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =48
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =9
    ItemSuffix =39
    Left =3855
    Top =1515
    Right =18180
    Bottom =11505
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2680758ff389e340
    End
    Caption =" Data Summary Basics"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =9900
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =24
                    Left =4500
                    Top =600
                    Width =9780
                    Height =420
                    FontSize =14
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbQuery"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MSysObjects.Name, MSysObjects.Type, \"GetQueryDescription([Name])\" AS De"
                        "scription, * FROM MSysObjects WHERE (((MSysObjects.Name) Like \"qSumB_*\") AND ("
                        "(MSysObjects.Type)=5)) ORDER BY MSysObjects.Name;"
                    ColumnWidths ="5760"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    OnNotInList ="[Event Procedure]"

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =4500
                    LayoutCachedTop =600
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =1020
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =1080
                    Width =14160
                    Height =7320
                    TabIndex =1
                    Name ="subResults"

                    LayoutCachedLeft =120
                    LayoutCachedTop =1080
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =8400
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Width =14400
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUtilities_Header"
                    Caption ="Basic Summary Tools"
                    FontName ="Calibri"
                    GridlineColor =10921638
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =540
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    ForeThemeColorIndex =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10260
                    Top =8520
                    Width =2580
                    Height =1260
                    FontSize =13
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="cmdOpen_Advanced_Tools"
                    Caption ="Advanced Summary Tools "
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the Advanced Data Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =10260
                    LayoutCachedTop =8520
                    LayoutCachedWidth =12840
                    LayoutCachedHeight =9780
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =12960
                    Top =8520
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdClose_Utilities"
                    Caption ="Close"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =12960
                    LayoutCachedTop =8520
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =9780
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =10798077
                    HoverThemeColorIndex =5
                    HoverTint =40.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1140
                    Top =8580
                    Width =1260
                    Height =509
                    FontSize =13
                    FontWeight =700
                    TabIndex =4
                    ForeColor =0
                    Name ="cmdExport_to_Excel"
                    Caption ="To Excel"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Export this results to Excel"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =8580
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =9089
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =1140
                    Top =9180
                    Width =1260
                    Height =509
                    FontSize =13
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="cmdExport_to_Text"
                    Caption ="To Text"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Export the results to a text file"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =9180
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =9689
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =120
                    Top =8520
                    Width =960
                    Height =1260
                    FontSize =12
                    FontWeight =700
                    BackColor =855309
                    ForeColor =16777215
                    Name ="Label36"
                    Caption ="       \015\012Export Data"
                    FontName ="Calibri"
                    LayoutCachedLeft =120
                    LayoutCachedTop =8520
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =9780
                    BackThemeColorIndex =0
                    BackTint =95.0
                    ForeThemeColorIndex =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =215
                    Left =120
                    Top =8520
                    Width =2340
                    Height =1260
                    BorderColor =855309
                    Name ="Box37"
                    LayoutCachedLeft =120
                    LayoutCachedTop =8520
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =9780
                    BorderThemeColorIndex =0
                    BorderTint =95.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =120
                    Top =600
                    Width =4320
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label38"
                    Caption ="What would you like to know?   --->"
                    FontName ="Calibri"
                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1020
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

Private Sub cmdExport_to_Excel_Click()
    On Error GoTo Err_Handler

    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    strQryName = Me.cmbQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xlsx"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls*)", "*.xls*")
    'DoCmd.TransferSpreadsheet acOutputQuery, 10, strQryName, strSaveFile, True
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatXLSX, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure
End Sub

Private Sub cmdExport_to_Text_Click()
    On Error GoTo Err_Handler

    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    strQryName = Me.cmbQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".txt"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.txt)", "*.txt")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatTXT, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdGetting_Started_Click()
    subResults.visible = True
    cmbQuery.SetFocus
   ' cmdGetting_Started.Visible = False
    cmbQuery.Dropdown
End Sub

Private Sub cmdOpen_Advanced_Tools_Click()
    On Error GoTo Err_Handler

    'record what the current record is so we can go back to that record on return
    DoCmd.Close acForm, "frm_Data_Summary_Basic"
    DoCmd.OpenForm "frm_Data_Summary_Advanced"
        
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' =================================
' FORM NAME:    frm_Summary_Tool
' Description:  Standard form for summarizing and exploring project data
' Data source:  unbound
' Data access:  edit only, no additions or deletions
' Pages:        none
' Functions:    fxnFilterRecords
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, Jan 2010
' Revisions:    <name, date - desc>
' =================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close the form if the switchboard is not open
    If fxnSwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If
    
cmbQuery.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbQuery_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Me.ActiveControl.Undo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbQuery_AfterUpdate()
    On Error GoTo Err_Handler

    ' Exit if no query selected
    If IsNull(Me.cmbQuery) Then
        'Me.txtUnfilteredFlag = ""
        'Me.txtUnfilteredFlag.ForeColor = 0          'black
        'Me.txtUnfilteredFlag.BackColor = 8454143    'yellow
        Me.subResults.SourceObject = ""
        GoTo Exit_Procedure
    End If

    Dim qdf As DAO.QueryDef
    Dim qdfs As DAO.QueryDefs
    Set qdfs = DBEngine(0)(0).QueryDefs

    On Error GoTo Err_Handler
    ' Bind the subform to the newly-selected object
    Me.subResults.Enabled = True
    Me.subResults.visible = True
    Me.subResults.SourceObject = "Query." & Me.cmbQuery.Value

    ' Set focus to the subform to allow scrolling, etc.
    Me.subResults.SetFocus

Exit_Procedure:
    On Error Resume Next
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdOpenBrowser_Click()
    On Error GoTo Err_Handler

    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.subResults
    ' Open to a blank record - to distinguish from opening to the selected record in the subform
    DoCmd.OpenForm "frm_Data_Browser", , , , acFormAdd, , "off"

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdRequery_Click()
    On Error GoTo Err_Handler

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    ' Requery the selected record in the recordset, and update the subform
    Me.subResults.Requery
    Me.subResults.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub


' =================================
' The next set of procedures relate to manipulating the selected query/results

Private Sub cmdPivotTable_Click()
    On Error GoTo Err_Handler

    ' Open the selected query as a pivot table after checking that a query is selected
    If IsNull(Me.cmbQuery) = False Then
        DoCmd.OpenQuery Me.cmbQuery.Value, acViewPivotTable, acReadOnly
        DoCmd.Maximize
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cmbQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdExportExcel_Click()
    On Error GoTo Err_Handler

    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    strQryName = Me.cmbQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xls"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls)", "*.xls")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatXLS, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdExportText_Click()
    On Error GoTo Err_Handler

    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cmbQuery) Then GoTo Exit_Procedure

    strQryName = Me.cmbQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".txt"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.txt)", "*.txt")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatTXT, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub
