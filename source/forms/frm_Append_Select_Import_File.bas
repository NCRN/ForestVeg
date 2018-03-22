Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10620
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =1995
    Top =5760
    Right =12360
    Bottom =7650
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcaee3e1e0fb2e340
    End
    RecordSource ="tsys_Import_File"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
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
        Begin FormHeader
            Height =780
            BackColor =15527148
            Name ="FormHeader"
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
                    Name ="cmd_Import_Log"
                    Caption ="View Import Log"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =120
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Line
                    OverlapFlags =87
                    Top =600
                    Width =10500
                    Name ="Line9"
                    LayoutCachedTop =600
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =600
                End
            End
        End
        Begin Section
            Height =1380
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =2340
                    Name ="Import_Name"
                    ControlSource ="Import_Name"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =360
                    Width =5520
                    TabIndex =1
                    Name ="txt_Import_File_Name"
                    ControlSource ="Import_File_Name"

                    LayoutCachedLeft =180
                    LayoutCachedTop =360
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =180
                    Top =660
                    Width =9480
                    TabIndex =2
                    Name ="txt_Import_File"
                    ControlSource ="Import_File_Loc"

                    LayoutCachedLeft =180
                    LayoutCachedTop =660
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =900
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9720
                    Top =600
                    Width =780
                    TabIndex =3
                    Name ="cmd_Browse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =600
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =960
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =1020
                    Width =1860
                    Height =300
                    TabIndex =4
                    Name ="cmd_Import_Tables"
                    Caption ="Import Tables"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1320
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2220
                    Top =1020
                    Width =780
                    Height =300
                    TabIndex =5
                    Name ="cmd_Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1320
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmd_Browse_Click()

  Dim varImportFileName As Variant
  Dim arrFile() As String
  
  'Select the file to import
  varImportFileName = GetImportFile()
  
  If IsNull(varImportFileName) Then
        Exit Sub
    Else
        Me!txt_Import_File = varImportFileName
    End If
    
    arrFile = Split(varImportFileName, "\")
    Me!txt_Import_File_Name = arrFile(UBound(arrFile))
End Sub

Private Sub cmd_Close_Click()
On Error GoTo Err_cmd_Close_Click

    DoCmd.Close

Exit_cmd_Close_Click:
    Exit Sub
Err_cmd_Close_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Close_Click
End Sub

Private Sub cmd_Import_Tables_Click()

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

'Pull the filename to be imported from a text box on the form
strPath = Me!txt_Import_File.Value
strImportFileRole = Mid(strPath, InStrAtPos(strPath, "_", 0) + 1, InStrAtPos(strPath, ".", 0) - InStrAtPos(strPath, "_", 0) - 1)
strImportFileDate = Mid(strPath, InStrAtPos(strPath, "_", CharacterCount(strPath, "_") - 2) + 1, 8)

'Open the database that contains the objects for import
Set dbImport = DBEngine.Workspaces(0).OpenDatabase(strPath, True)
strDate = Date
Set dbMain = CurrentDb
Set rsImportLog = dbMain.OpenRecordset("tsys_Import_Log")
Set rsImportTablesList = dbMain.OpenRecordset("tsys_Import_Tables")

'Populate the RS
rsImportTablesList.MoveFirst
intRC = rsImportTablesList.RecordCount
   
'Loop through tsys_Import_Tables to see of the table should be imported

Do Until rsImportTablesList.EOF
      
     'For each table in the importing data set check to see if:
        'the name matches the import table selected
        For Each td In dbImport.TableDefs
            strTableToImport = td.Name
                                
                If strTableToImport = rsImportTablesList![Table_Name] Then
                    'If the name matches and the import box is checked then:
                    If rsImportTablesList![Import] = True Then
                        'Rename the import table
                        strTableToImport_NewName = "_" & strTableToImport & "_Import_" & strImportFileDate & "_" & strImportFileRole
                            Dim tdefMain As TableDef
StartOver:
                            'Loop through the main data set to see if the new import table name is already taken.
                            For Each tdefMain In dbMain.TableDefs
                                Dim counter As Integer
                               
                                If strTableToImport_NewName = tdefMain.Name Then
                                    'If the name has already been taken then:
                                    If Left(Right(tdefMain.Name, 2), 1) = "_" Then
                                        'Assign a new sequential number to the duplicate table name
                                        Dim iLength As Integer
                                        iLength = Len(strTableToImport_NewName)
                                        Dim strTdefTemp As String
                                        counter = Right(tdefMain.Name, 1)
                                                                                                                            
                                        strTdefTemp = Left(strTableToImport_NewName, (iLength - 2))
                                        '
                                        strTableToImport_NewName = strTdefTemp & "_" & counter + 1
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
                    Select Case strTableToImport
                    Dim strDeleteExistingEventsQry As String
                    Case "tbl_Events"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "INNER JOIN tbl_Events ON [" & strTableToImport_NewName & "].Event_ID = tbl_Events.Event_ID;"
                        dbMain.Execute strDeleteExistingEventsQry
                    Case "tbl_Tree_Data", "tbl_Sapling_Data", "tbl_Quadrat_Data", "tbl_Plot_Floor_Condition_Data", "xref_Event_Contacts", "tbl_CWD_Data"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "LEFT JOIN [_tbl_Events" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                        & "] ON [" & strTableToImport_NewName & "].[Event_ID] = [_tbl_Events_Import_" & strImportFileDate _
                        & "_" & strImportFileRole & "].[Event_ID] " _
                        & "WHERE (([_tbl_Events_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Event_ID]) Is Null);"
                        dbMain.Execute strDeleteExistingEventsQry
                    Case "tbl_Tree_DBH", "tbl_Tree_Conditions", "tbl_Tree_Foliage_Conditions", "tbl_Tree_Vines"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "LEFT JOIN [_tbl_Tree_Data" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                        & "] ON [" & strTableToImport_NewName & "].[Tree_Data_ID] = [_tbl_Tree_Data_Import_" & strImportFileDate _
                        & "_" & strImportFileRole & "].[Tree_Data_ID] " _
                        & "WHERE (([_tbl_Tree_Data_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Tree_Data_ID]) Is Null);"
                        dbMain.Execute strDeleteExistingEventsQry
                    Case "tbl_Sapling_DBH"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "LEFT JOIN [_tbl_Sapling_Data" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                        & "] ON [" & strTableToImport_NewName & "].[Sapling_Data_ID] = [_tbl_Sapling_Data_Import_" & strImportFileDate _
                        & "_" & strImportFileRole & "].[Sapling_Data_ID] " _
                        & "WHERE (([_tbl_Sapling_Data_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Sapling_Data_ID]) Is Null);"
                        dbMain.Execute strDeleteExistingEventsQry
                    Case "tbl_Quadrat_Seedlings_Data", "tbl_Quadrat_Herbaceous_Data"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "LEFT JOIN [_tbl_Quadrat_Data" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                        & "] ON [" & strTableToImport_NewName & "].[Quadrat_Data_ID] = [_tbl_Quadrat_Data_Import_" & strImportFileDate _
                        & "_" & strImportFileRole & "].[Quadrat_Data_ID] " _
                        & "WHERE (([_tbl_Quadrat_Data_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Quadrat_Data_ID]) Is Null);"
                        dbMain.Execute strDeleteExistingEventsQry
                    Case "tbl_Tags", "tbl_Tasks"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "LEFT JOIN [_tbl_Events" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                        & "] ON [" & strTableToImport_NewName & "].[Location_ID] = [_tbl_Events_Import_" & strImportFileDate _
                        & "_" & strImportFileRole & "].[Location_ID] " _
                        & "WHERE (([_tbl_Events_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Location_ID]) Is Null);"
                        dbMain.Execute strDeleteExistingEventsQry
                    Case "tbl_Tags_History"
                        strDeleteExistingEventsQry = "DELETE [" & strTableToImport_NewName & "].* " _
                        & "FROM [" & strTableToImport_NewName & "] " _
                        & "LEFT JOIN [_tbl_Tags" & "_Import_" & strImportFileDate & "_" & strImportFileRole _
                        & "] ON [" & strTableToImport_NewName & "].[Record_ID] = [_tbl_Tags_Import_" & strImportFileDate _
                        & "_" & strImportFileRole & "].[Tag_ID] " _
                        & "WHERE (([_tbl_Tags_Import_" & strImportFileDate & "_" & strImportFileRole & "].[Tag_ID]) Is Null);"
                        dbMain.Execute strDeleteExistingEventsQry
                    End Select
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
    rsImportLog![Import_Records] = rsNew.RecordCount
    rsImportLog.Update
    Set rsNew = Nothing

Next_Record:
    Next
    rsImportTablesList.MoveNext
    
Loop

Dim Response As String

If intImport2 = 2 Then

    DoCmd.Close
    DoCmd.OpenForm ("frm_Append_Select_Import_Tables")
       
    Exit Sub
    
End If

Response = MsgBox("Import Complete! Would you like to proceed with appending data?", vbYesNo, "Import Data Tables")
    
    If Response = vbYes Then
        DoCmd.Close
        DoCmd.OpenForm ("frm_Append_Append_Data")
    Else
        DoCmd.Close
    End If

intImport2 = 0

Set dbMain = Nothing
Set dbImport = Nothing
Set td = Nothing
Set rsNew = Nothing
Set rsImportLog = Nothing
Set rsImportTablesList = Nothing

End Sub

Private Sub cmd_Import_Log_Click()
On Error GoTo Err_cmd_Import_Log_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Import_Log"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_cmd_Import_Log_Click:
    Exit Sub
Err_cmd_Import_Log_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Import_Log_Click
End Sub
