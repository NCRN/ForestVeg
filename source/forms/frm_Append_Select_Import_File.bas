Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10620
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =75
    Top =195
    Right =10695
    Bottom =3240
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcaee3e1e0fb2e340
    End
    RecordSource ="tsys_Import_File"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =780
            BackColor =15527148
            Name ="FormHeader"
            AutoHeight =255
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =120
                    Top =120
                    Width =4365
                    Height =480
                    FontSize =16
                    Name ="Label5"
                    Caption ="Select and Import Data Tables"
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =600
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8700
                    Top =120
                    Width =1800
                    FontWeight =700
                    ForeColor =0
                    Name ="btnImportLog"
                    Caption ="View Import Log"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8700
                    LayoutCachedTop =120
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin Line
                    OverlapFlags =87
                    Top =600
                    Width =10500
                    Name ="lnHdr"
                    LayoutCachedTop =600
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =600
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1740
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =480
                    Width =2340
                    FontWeight =600
                    Name ="lblImportFile"
                    ControlSource ="Import_Name"

                    LayoutCachedLeft =180
                    LayoutCachedTop =480
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =780
                    Width =5520
                    Height =360
                    TabIndex =1
                    LeftMargin =30
                    TopMargin =43
                    Name ="txt_Import_File_Name"
                    ControlSource ="Import_File_Name"

                    LayoutCachedLeft =180
                    LayoutCachedTop =780
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1140
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =1260
                    Width =9480
                    Height =360
                    TabIndex =2
                    LeftMargin =30
                    TopMargin =43
                    Name ="txt_Import_File"
                    ControlSource ="Import_File_Loc"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =1620
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =9720
                    Top =1260
                    Width =780
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="btnBrowse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9720
                    LayoutCachedTop =1260
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1620
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1095
                    Top =60
                    Width =1905
                    Height =299
                    FontSize =8
                    TabIndex =4
                    BackColor =13434879
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200220000000000
                    End
                    Name ="cbxImporter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contact_ID, Last_Name, First_Name, Last_Name & ', ' & First_Name AS Pick_"
                        "List FROM tlu_Contacts WHERE Active = True ORDER BY Last_Name;"
                    ColumnWidths ="0;0;0;2160"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Choose who is importing data"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1095
                    LayoutCachedTop =60
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =359
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =180
                            Top =60
                            Width =855
                            Height =299
                            FontWeight =700
                            Name ="lblImporter"
                            Caption ="Importer"
                            ControlTipText ="Choose who is importing data"
                            LayoutCachedLeft =180
                            LayoutCachedTop =60
                            LayoutCachedWidth =1035
                            LayoutCachedHeight =359
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =60
                    Width =1860
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="btnUpdateContacts"
                    Caption ="Update Contact List"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Go to the contact list & update information"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =540
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =180
                    Top =60
                    Width =1860
                    FontWeight =700
                    ForeColor =0
                    Name ="btnImportTables"
                    Caption ="Import Tables"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9720
                    Top =60
                    Width =780
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =255
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =2100
                    Top =60
                    Width =3900
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="btnSkipImport"
                    Caption ="Skip Import && Use Already  Imported Tables"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin Line
                    OverlapFlags =85
                    Width =10500
                    Name ="lnFooter"
                    LayoutCachedWidth =10500
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7440
                    Top =60
                    Width =2100
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="btnDeleteTables"
                    Caption ="Delete Import Tables"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open delete existing import table(s) form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7440
                    LayoutCachedTop =60
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    HoverColor =65280
                    PressedColor =0
                    PressedThemeColorIndex =0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
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
' MODULE:       frm_Append_Select_Import_File
' Level:        Application module
' Version:      1.04
'
' Description:  field data import related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 8/31/2019 - 1.01 - added documentation, error handling, option explicit,
'               BLC   - 9/3/2019  - 1.02 - add EOF/BOF checks before recordcounts
'               BLC   - 9/20/2019 - 1.03 - add importer to capture who is importing events
'               BLC   - 9/24/2019 - 1.04 - add delete import tables button, updated importer sort order
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Events
'---------------------
' ----------------
'  Form
' ----------------
' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, August 31, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 8/31/2019 - added documentation, error handling
'   BLC  - 9/20/2019 - populated importer list
'   BLC  - 9/24/2019 - updated importer sort order
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim sql As String
    
    'set importer contacts
    sql = "SELECT Contact_ID, Last_Name, First_Name, Last_Name & ', ' & First_Name AS Pick_List FROM tlu_Contacts " & _
            "WHERE Active = True ORDER BY Last_Name, First_Name;"
            
    'Debug.Print sql
    
    With cbxImporter
        .RowSource = sql
        .ColumnCount = 4
        .BoundColumn = 1
        .ColumnWidths = "0;0;0;1.5in;"
    End With
    
    'default
    Me.btnBrowse.Enabled = False
    Me.btnImportTables.Enabled = False
    Me.btnSkipImport.Enabled = False
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnUpdateContacts_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 20, 2019
' Adapted:      -
' Revisions:
'   BLC  - 9/20/2019 - initial version
' ---------------------------------
Private Sub btnUpdateContacts_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "frm_Contacts", acNormal, , , acFormEdit, acWindowNormal
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUpdateContacts_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnBrowse
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, August 31, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 8/31/2019 - added documentation, error handling
'   BLC  - 9/20/2019 - renamed btnBrowse from cmd_Browse, enabled import files button
' ---------------------------------
Private Sub btnBrowse_Click()
On Error GoTo Err_Handler

    Dim varImportFileName As Variant
    Dim arrFile() As String
    
    'Select the file to import
    varImportFileName = GetImportFile()
    
    If IsNull(varImportFileName) Then
        Exit Sub
    Else
        Me!txt_Import_File = varImportFileName
        Me.btnImportTables.Enabled = True
    End If
    
    arrFile = Split(varImportFileName, "\")
    Me!txt_Import_File_Name = arrFile(UBound(arrFile))

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnImportTables_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, August 31, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 8/31/2019 - added documentation, error handling
'   BLC  - 9/3/2019  - add BOF/EOF check before move first/last to
'                      get accurate recordcount
'   BLC  - 9/20/2019 - renamed to btnImportTables from cmd_Import_Tables
'                      added pseudoevent processing to update pseudoevent IDs
' ---------------------------------
Private Sub btnImportTables_Click()
On Error GoTo Err_Handler

    'On Error Resume Next
    
    Dim rsImportTablesList As DAO.Recordset
    Dim dbImport As DAO.Database 'Database to import
    Dim dbMain  As DAO.Database
    Dim rsImportLog As DAO.Recordset
    Dim rsNew As DAO.Recordset
    Dim td As TableDef 'Table Defs in DB
    Dim strTableToImport As String 'Name of a table to import
    Dim strTableToImport_NewName As String 'New name for the imported table
    Dim strDate As String
    Dim strPath As String
    Dim strImportFileDate As String
    Dim strImportFileRole As String
    Dim intRC As Integer
        
    'Dim strImportMsg As String
    
    DoCmd.Hourglass True
    
    'processing
    Application.SysCmd acSysCmdSetStatus, "Processing import tables..."
    
    'Pull the filename to be imported from a text box on the form
    strPath = Me!txt_Import_File.Value
    strImportFileRole = Mid(strPath, InStrAtPos(strPath, "_", 0) + 1, InStrAtPos(strPath, ".", 0) - InStrAtPos(strPath, "_", 0) - 1)
    strImportFileDate = Mid(strPath, InStrAtPos(strPath, "_", CharacterCount(strPath, "_") - 2) + 1, 8)
Debug.Print strImportFileRole

    'Open the database that contains the objects for import
    Set dbImport = DBEngine.Workspaces(0).OpenDatabase(strPath, True)
    strDate = Date
    Set dbMain = CurrentDb
    Set rsImportLog = dbMain.OpenRecordset("tsys_Import_Log")
    Set rsImportTablesList = dbMain.OpenRecordset("tsys_Import_Tables")
    
    'Populate the RS
    If Not rsImportTablesList.BOF And rsImportTablesList.EOF Then
        rsImportTablesList.MoveLast
        rsImportTablesList.MoveFirst
    End If
    intRC = rsImportTablesList.RecordCount
       
    'Loop through tsys_Import_Tables to see of the table should be imported
    Do Until rsImportTablesList.EOF
          
         'For each table in the importing data set check to see if:
         'the name matches the import table selected
            For Each td In dbImport.TableDefs
                strTableToImport = td.Name
 Debug.Print td.Name
                If strTableToImport = rsImportTablesList![Table_Name] Then
                    'If the name matches and the import box is checked then:
                    If rsImportTablesList![Import] = True Then
                        'Rename the import table
                        strTableToImport_NewName = "_" & strTableToImport & "_Import_" & strImportFileDate & "_" & strImportFileRole
                        
                        'processing
                        Application.SysCmd acSysCmdSetStatus, "Importing " & strTableToImport_NewName & "..."
                        
                        Dim tdefMain As TableDef
StartOver:
                        'Loop through the main data set to see if the new import table name is already taken.
                        For Each tdefMain In dbMain.TableDefs
                            Dim Counter As Integer
                           
                            If strTableToImport_NewName = tdefMain.Name Then
                                'If the name has already been taken then:
                                If Left(Right(tdefMain.Name, 2), 1) = "_" Then
                                    'Assign a new sequential number to the duplicate table name
                                    Dim iLength As Integer
                                    iLength = Len(strTableToImport_NewName)
                                    Dim strTdefTemp As String
                                    Counter = Right(tdefMain.Name, 1)
                                                                                                                        
                                    strTdefTemp = Left(strTableToImport_NewName, (iLength - 2))
                                    '
                                    strTableToImport_NewName = strTdefTemp & "_" & Counter + 1
                                    'counter2 = counter2 + 1
                                    dbMain.TableDefs.Refresh
                                Else
                                    'otherwise
                                    strTableToImport_NewName = strTableToImport_NewName & "_1" '& counter
                                    dbMain.TableDefs.Refresh
                                End If
                                GoTo StartOver:
                            End If
                        Next tdefMain
                        
                        DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, strTableToImport, strTableToImport_NewName, False
                        'IF Importing Events Table then delete events in the temporary table which already exist in the main database. Added mel 9/27/2010.
                        'If strTableToImport = "tbl_Events" Then
                        '    Dim strDeleteExistingEventsQry As String
                        '    strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        '    & "FROM [" & strTableToImport_NewName & "] " _
                        '    & "INNER JOIN tbl_Events ON [" & strTableToImport_NewName & "].Event_ID = tbl_Events.Event_ID;"
                        '    dbMain.Execute strDeleteExistingEventsQry
                        'End If
                                                
                        Dim strDeleteExistingEventsQry As String
                        Dim EventDeleteQuery As String      'store until AFTER pseudoevents are addressed
                        Dim ImportEventTableName As String  'store for handling pseudoevents

Debug.Print strTableToImport

                        Select Case strTableToImport
                            
                            Case "tbl_Events"
                                'strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* "
                                Dim strWHERE As String
                                strWHERE = ""
                                Select Case strImportFileRole
                                    Case "PRIMARY"
                                    Case "SECONDARY"
                                        strWHERE = " WHERE YEAR(e.Event_Date) < Year(Now()-1)"
                                End Select
                                EventDeleteQuery = "DELETE i.* " _
                                    & "FROM [" & strTableToImport_NewName & "] i " _
                                    & "INNER JOIN tbl_Events e ON i.Event_ID = e.Event_ID" _
                                    & strWHERE _
                                    & ";"
                                    
                                'dbMain.Execute strDeleteExistingEventsQry
                                ImportEventTableName = strTableToImport_NewName
Debug.Print strDeleteExistingEventsQry
                            Case "tbl_Tree_Data", "tbl_Sapling_Data", "tbl_Quadrat_Data", "tbl_Plot_Floor_Condition_Data", "xref_Event_Contacts", "tbl_CWD_Data"
                                strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                                    & "FROM [" & strTableToImport_NewName & "] " _
                                    & "LEFT JOIN [_tbl_Events" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                                    & "] ON [" & strTableToImport_NewName & "].[Event_ID] = [_tbl_Events_Import_" & strImportFileDate _
                                    & "_" & strImportFileRole & "].[Event_ID] " _
                                    & "WHERE (([_tbl_Events_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Event_ID]) Is Null);"
                                'dbMain.Execute strDeleteExistingEventsQry
                            
                            Case "tbl_Tree_DBH", "tbl_Tree_Conditions", "tbl_Tree_Foliage_Conditions", "tbl_Tree_Vines"
                                strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                                    & "FROM [" & strTableToImport_NewName & "] " _
                                    & "LEFT JOIN [_tbl_Tree_Data" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                                    & "] ON [" & strTableToImport_NewName & "].[Tree_Data_ID] = [_tbl_Tree_Data_Import_" & strImportFileDate _
                                    & "_" & strImportFileRole & "].[Tree_Data_ID] " _
                                    & "WHERE (([_tbl_Tree_Data_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Tree_Data_ID]) Is Null);"
                                'dbMain.Execute strDeleteExistingEventsQry
                            
                            Case "tbl_Sapling_DBH"
                                strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                                    & "FROM [" & strTableToImport_NewName & "] " _
                                    & "LEFT JOIN [_tbl_Sapling_Data" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                                    & "] ON [" & strTableToImport_NewName & "].[Sapling_Data_ID] = [_tbl_Sapling_Data_Import_" & strImportFileDate _
                                    & "_" & strImportFileRole & "].[Sapling_Data_ID] " _
                                    & "WHERE (([_tbl_Sapling_Data_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Sapling_Data_ID]) Is Null);"
                                'dbMain.Execute strDeleteExistingEventsQry
                            
                            Case "tbl_Quadrat_Seedlings_Data", "tbl_Quadrat_Herbaceous_Data"
                                strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                                    & "FROM [" & strTableToImport_NewName & "] " _
                                    & "LEFT JOIN [_tbl_Quadrat_Data" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                                    & "] ON [" & strTableToImport_NewName & "].[Quadrat_Data_ID] = [_tbl_Quadrat_Data_Import_" & strImportFileDate _
                                    & "_" & strImportFileRole & "].[Quadrat_Data_ID] " _
                                    & "WHERE (([_tbl_Quadrat_Data_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Quadrat_Data_ID]) Is Null);"
                                'dbMain.Execute strDeleteExistingEventsQry
                            
                            Case "tbl_Tags", "tbl_Tasks"
                                strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                                    & "FROM [" & strTableToImport_NewName & "] " _
                                    & "LEFT JOIN [_tbl_Events" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                                    & "] ON [" & strTableToImport_NewName & "].[Location_ID] = [_tbl_Events_Import_" & strImportFileDate _
                                    & "_" & strImportFileRole & "].[Location_ID] " _
                                    & "WHERE (([_tbl_Events_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Location_ID]) Is Null);"
                                'dbMain.Execute strDeleteExistingEventsQry
                            
                            Case "tbl_Tags_History"
                                strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                                    & "FROM [" & strTableToImport_NewName & "] " _
                                    & "LEFT JOIN [_tbl_Tags" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                                    & "] ON [" & strTableToImport_NewName & "].[Record_ID] = [_tbl_Tags_Import_" & strImportFileDate _
                                    & "_" & strImportFileRole & "].[Tag_ID] " _
                                    & "WHERE (([_tbl_Tags_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Tag_ID]) Is Null);"
                                'dbMain.Execute strDeleteExistingEventsQry
                        End Select
 Debug.Print strTableToImport & " Delete Query: " & strDeleteExistingEventsQry
                        If Not IsNothing(strDeleteExistingEventsQry) = True Then _
                            dbMain.Execute strDeleteExistingEventsQry
                    Else
                        GoTo Next_Record
                    End If
                Else
                    GoTo Next_Record:
                End If
    
        'Create New Record in Import Log with Table Name, Import Date, and Record Count
        rsImportLog.AddNew
        rsImportLog![Table_Name] = strTableToImport_NewName
        rsImportLog![Import_Date] = strDate
            
        Set rsNew = dbMain.OpenRecordset(strTableToImport_NewName)
        Dim intRecCount As Integer
        
        'get accurate recordcount
        If Not (rsNew.BOF And rsNew.EOF) Then
            rsNew.MoveLast
            rsNew.MoveFirst
        End If
        rsImportLog![Import_Records] = rsNew.RecordCount
        rsImportLog.Update
        Set rsNew = Nothing
    
Next_Record:
        Next
        rsImportTablesList.MoveNext
        
    Loop
    
    Dim response As String
    
    If intImport2 = 2 Then
    
        DoCmd.Close
        DoCmd.OpenForm ("frm_Append_Select_Import_Tables")
           
        Exit Sub
        
    End If
    
    'copy original import table
    DoCmd.CopyObject CurrentDb.Name, "_ORIG" & ImportEventTableName, acTable, ImportEventTableName
    
    'handle pseudoevents BEFORE appends --> updates EventIDs in import tables, archives & deletes tbl_Events EventIDs
    UpdatePseudoEventIDs ImportEventTableName
Debug.Print "e delete qry: " & EventDeleteQuery
    'delete existing events BEFORE appends
    If Not IsNothing(EventDeleteQuery) = True Then _
        dbMain.Execute EventDeleteQuery
    
    response = MsgBox("Import Complete! Would you like to proceed with appending data?", vbYesNo, "Import Data Tables")
        
    If response = vbYes Then
        Dim ImportFile As String
        ImportFile = Me.txt_Import_File_Name
        DoCmd.Close
    
        DoCmd.OpenForm "frm_Append_Append_Data", , , , , , ImportFile
    Else
        DoCmd.Close
    End If
    
    intImport2 = 0
    
Exit_Handler:
    'cleanup
    DoCmd.Hourglass False
    Application.SysCmd acSysCmdClearStatus
    Set dbMain = Nothing
    Set dbImport = Nothing
    Set td = Nothing
    Set rsNew = Nothing
    Set rsImportLog = Nothing
    Set rsImportTablesList = Nothing
    
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnImportTables_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSkipImport
' Description:  button click actions
' Assumptions:  tables have already been imported
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 31, 2019
' Adapted:      -
' Revisions:
'   BLC  - 8/31/2019 - initial version
' ---------------------------------
Private Sub btnSkipImport_Click()
On Error GoTo Err_Handler

    DoCmd.Close
    DoCmd.OpenForm "frm_Append_Append_Data"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSkipImport_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDeleteTables_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 24, 2019
' Adapted:      -
' Revisions:
'   BLC  - 9/24/2019 - initial version
' ---------------------------------
Private Sub btnDeleteTables_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "frm_Append_Delete_Tables", acNormal, , , acFormEdit, acWindowNormal
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDeleteTables_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnImportLog_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, August 31, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 8/31/2019 - added documentation, error handling
'   BLC  - 9/20/2019 - renamed to btnImportLog vs cmd_Import_Log
' ---------------------------------
Private Sub btnImportLog_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Import_Log"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnImportLog_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClose
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, August 31, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 8/31/2019 - added documentation, error handling
'   BLC  - 9/20/2019 - renamed to btnClose from cmd_Close
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxImporter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 20, 2019
' Adapted:      -
' Revisions:
'   BLC  - 9/20/2019 - initial version
' ---------------------------------
Private Sub cbxImporter_AfterUpdate()
On Error GoTo Err_Handler

    SetTempVar "ImportContact", cbxImporter.Value
    
    'enable file selection if there's a contact selected
    If Not IsNothing(cbxImporter.Value) Then
        Me.btnBrowse.Enabled = True
        Me.btnSkipImport.Enabled = True
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxImporter_AfterUpdate[frm_Append_Select_Import_File form])"
    End Select
    Resume Exit_Handler
End Sub
