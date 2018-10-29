Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6120
    DatasheetFontHeight =9
    ItemSuffix =4
    Left =3660
    Top =3675
    Right =9780
    Bottom =6960
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xed7c66ab750be340
    End
    Caption ="Select Tables to Link"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Section
            Height =3300
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    OverlapFlags =87
                    MultiSelect =1
                    IMESentenceMode =3
                    Left =240
                    Top =420
                    Width =5640
                    Height =2100
                    Name ="lstTables"
                    RowSourceType ="Table/Query"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =240
                            Top =180
                            Width =5250
                            Height =240
                            Name ="Label1"
                            Caption ="Select the tables to link from the list below and then click the Link button"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1500
                    Top =2760
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdLink"
                    Caption ="Link"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3240
                    Top =2760
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
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
' Description:  Table selection form
' Data source:  unbound
' Data access:  NA
' Pages:        none
' Functions:    none
' References:   XML_Read, CorrectText
' Source/date:  Simon Kingston, Sept. - Dec., 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cmdCancel_Click()
DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdLink_Click()
'add the selected tables to the list of linked tables for the NRDT app.
Dim varItem As Variant
Dim strTableName As String
Dim strFileName As String
Dim strLinkType As String
Dim tdf As DAO.TableDef
Dim dbExternal As DAO.Database
Dim tdfExternal As DAO.TableDef
Dim strDescription As String
Dim strSQL As String
Dim strValues As String

On Error GoTo Err_Handler

If Me!lstTables.ItemsSelected.Count = 0 Then
    MsgBox "There are no tables selected.", vbExclamation, "No Tables Selected"
Else
    strFileName = XML_Read("FileName", Nz(Me.OpenArgs, ""))
    strLinkType = CorrectText(XML_Read("LinkType", Nz(Me.OpenArgs, "")))
    
    ' Enumerate through selected items.
    For Each varItem In Me!lstTables.ItemsSelected
        strTableName = Me!lstTables.ItemData(varItem)
        
        If Not IsNull(DLookup("Name", "MSysObjects", "Name=" & CorrectText(strTableName))) Then
            MsgBox "A table with the name " & strTableName & " already exists in the database.  That table will not be linked.", vbInformation + vbOKOnly, "Cannot Link Table"
        Else
            Set dbExternal = DBEngine.OpenDatabase(strFileName)
            Set tdfExternal = dbExternal.TableDefs(strTableName)
            strDescription = Nz(tdfExternal.Properties("Description"), "")
            'add table link
            Set tdf = CurrentDb.CreateTableDef(strTableName)
            tdf.SourceTableName = strTableName
            tdf.connect = ";DATABASE=" & strFileName
            CurrentDb.TableDefs.Append tdf
            
            'add table link record to link table
            strSQL = "INSERT INTO tsys_Link_Tables (Link_type,Link_table"
            If Len(strDescription) > 0 Then
                strSQL = strSQL & ",Description_text"
                strValues = "," & CorrectText(strDescription)
            Else
                strValues = ""
            End If
            strSQL = strSQL & ") VALUES (" & strLinkType & "," & CorrectText(strTableName)
            strSQL = strSQL & strValues & ");"
'            Debug.Print strSQL
            CurrentDb.Execute strSQL
        End If
    Next varItem
    
    MsgBox "Tables linked successfully!", vbOKOnly, "Tables Linked"
    DoCmd.Close acForm, Me.Name, acSaveNo
End If

Exit_Handler:
    On Error Resume Next
    dbExternal.Close
    Set dbExternal = Nothing
    Set tdfExternal = Nothing
    Set tdf = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3270
            strDescription = ""
            Resume Next
        Case Else
            MsgBox Err.Number & " - " & Err.Description
            Resume Exit_Handler
    End Select

End Sub

Private Sub Form_Load()
'populate a list of non-system tables in the selected database file
Dim strSQL As String
Dim strFileName As String

On Error GoTo Error_Handler

strFileName = XML_Read("FileName", Nz(Me.OpenArgs, ""))

strSQL = "SELECT Name FROM MSysObjects IN " & CorrectText(strFileName) & " WHERE Type=1 and Name NOT LIKE 'MSys*' and Name NOT LIKE 'tSys*' ORDER BY Name;"
Me!lstTables.RowSource = strSQL
Me!lstTables.Requery

Exit_Handler:
    Exit Sub

Error_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Error encountered"
    Resume Exit_Handler

End Sub
