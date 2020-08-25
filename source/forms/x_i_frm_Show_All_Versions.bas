Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8040
    DatasheetFontHeight =10
    ItemSuffix =25
    Left =2595
    Top =765
    Right =10635
    Bottom =5310
    DatasheetGridlinesColor =12632256
    Filter ="[version_key_number]=4"
    RecSrcDt = Begin
        0x8464019b33ffe240
    End
    RecordSource ="qry_master_version_by_number"
    Caption ="Show All Versions"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4560
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =480
                    Width =1080
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Protocol version key number (maintained in SOP #10)"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =480
                            Width =1860
                            Height =255
                            FontWeight =700
                            Name ="version_key_number_Label"
                            Caption ="Version Key Number"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =840
                    Width =1080
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="version_key_date"
                    ControlSource ="version_key_date"
                    Format ="Short Date"
                    StatusBarText ="Date of protocol version key number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =840
                            Width =1560
                            Height =255
                            FontWeight =700
                            Name ="version_key_date_Label"
                            Caption ="Version Key Date"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =1200
                    Width =1080
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="narrative_version"
                    ControlSource ="narrative_version"
                    Format ="Fixed"
                    StatusBarText ="Version of protocol narrative"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1200
                            Width =1560
                            Height =255
                            FontWeight =700
                            Name ="narrative_version_Label"
                            Caption ="Narrative Version"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1380
                    Top =1560
                    Width =4380
                    Height =420
                    ColumnWidth =3000
                    TabIndex =4
                    Name ="version_comments"
                    ControlSource ="version_comments"
                    StatusBarText ="Comments regarding version, if any"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1560
                            Width =1080
                            Height =420
                            FontWeight =700
                            Name ="version_comments_Label"
                            Caption ="Version Comments"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =4380
                    Top =120
                    Width =3300
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label10"
                    Caption ="Master Version Key Information"
                End
                Begin Subform
                    OverlapFlags =87
                    Left =120
                    Top =2520
                    Width =7860
                    Height =1800
                    TabIndex =5
                    Name ="Subform_Versions"
                    SourceObject ="Form.frm_sub_Show_All_Versions"
                    LinkChildFields ="version_key_number"
                    LinkMasterFields ="version_key_number"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =2280
                            Width =1965
                            Height =240
                            FontWeight =700
                            Name ="Label16"
                            Caption ="SOPs:"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2175
                    Left =1200
                    Top =120
                    Width =2220
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="project_ID"
                    ControlSource ="project_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_projects.project_ID, tlu_projects.project_name FROM tlu_projects; "
                    ColumnWidths ="0;2175"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =120
                            Width =900
                            Height =245
                            FontWeight =700
                            Name ="Project ID_Label"
                            Caption ="Project ID"
                            EventProcPrefix ="Project_ID_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3720
                    Top =1020
                    Height =300
                    TabIndex =6
                    Name ="ButtonPrint"
                    Caption ="Print Listing"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3720
                    Top =600
                    Width =1470
                    Height =300
                    TabIndex =7
                    Name ="ButtonAdd"
                    Caption ="Add New Version"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5880
                    Top =780
                    Width =1020
                    Height =300
                    TabIndex =8
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

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

Private Sub Form_Close()

  If CurrentProject.AllForms("frm_Version_List").IsLoaded = True Then
    Forms![frm_Version_List].Requery
  End If

End Sub

Private Sub Form_Load()
 On Error GoTo Err_Form_Load

        Dim intFID As Integer
        Dim db As DAO.Database
        Dim NewMaster As DAO.Recordset
        Dim NewSOP As DAO.Recordset
        Dim strSQL As String
        Dim Response As Integer
        Dim varInput As Variant
        Dim intIndex As Integer
        Dim intCount As Integer
        
    ' Check for no records
    If Me.RecordsetClone.RecordCount = 0 Then
       Response = MsgBox("There are no existing versions, do you want to initialize?", 4, "Add new version")
       If Response = 7 Then
         Cancel = True
         GoTo Exit_Form_Load
         Else
           varInput = InputBox("Enter the number of SOPs for this protocol.", "Initialize Protocol Versions", "10")
           Do Until IsNumeric(varInput)
             varInput = InputBox("You must enter a numeric value.", "Initialize Protocol Versions", "10")
           Loop
           intIndex = CInt(varInput)
           varInput = InputBox("Enter project ID, if known.", "Initialize Protocol Versions", "1")
           Do Until IsNumeric(varInput)
             varInput = InputBox("You must enter a numeric value.", "Initialize Protocol Versions", "1")
           Loop
           Set db = CurrentDb
           Set NewMaster = db.OpenRecordset("tbl_master_version") ' Open recordset for new SOP records
           NewMaster.AddNew
           NewMaster![project_ID] = CInt(varInput)
           NewMaster![version_key_number] = 1  ' This will be the first version key
           NewMaster![version_key_date] = Now()     ' Today's date
           NewMaster![narrative_version] = 1   ' Version 1
           NewMaster.Update ' Save parent record so child records can be added
           NewMaster.Close
        
           ' Add new child records for SOPs

           Set NewSOP = db.OpenRecordset("tbl_SOP_version") ' Open recordset for new SOP records
           intCount = 0   ' Initialize counter
           Do Until intCount = intIndex
             NewSOP.AddNew
             NewSOP![version_key_number] = 1
             intCount = intCount + 1  ' Increment counter
             NewSOP![SOP_number] = intCount
             NewSOP![SOP_version_number] = 1
             NewSOP![active_flag] = -1
             NewSOP.Update
           Loop
           NewSOP.Close
           Me!project_ID.Locked = False  ' Let them change the project ID on initialization.
           Me.Requery ' Refresh form SOPs from table
         End If   ' End if for inputbox response
       End If  ' End if for no records check
Exit_Form_Load:
    Exit Sub

Err_Form_Load:
    MsgBox Err.Description
    Resume Exit_Form_Load
   
    
End Sub
Private Sub ButtonPrint_Click()
On Error GoTo Err_ButtonPrint_Click

    Dim stDocName As String
    Dim strWHERE As String

    strWHERE = "version_key_number = " & Me![version_key_number]
    stDocName = "rpt_Show_All_Versions"
    DoCmd.OpenReport stDocName, acPreview, , strWHERE

Exit_ButtonPrint_Click:
    Exit Sub

Err_ButtonPrint_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPrint_Click
    
End Sub
Private Sub ButtonAdd_Click()
On Error GoTo Err_ButtonAdd_Click

        Dim intFID As Integer
        Dim db As DAO.Database
        Dim Versions As DAO.Recordset
        Dim VersionsByDate As DAO.Recordset
        Dim CurrentSOP As DAO.Recordset
        Dim NewSOP As DAO.Recordset
        Dim strSQL As String
        Dim Response As Integer
        
        DoCmd.GoToRecord , , acNewRec   ' Add new parent record for new version
        Set db = CurrentDb
        strSQL = "SELECT * FROM [tbl_master_version] ORDER BY [version_key_date]"
        Set VersionsByDate = db.OpenRecordset(strSQL)
        VersionsByDate.MoveLast ' Check if version exists for current date
        If Format(VersionsByDate![version_key_date], "short date") = Format(Date, "short date") Then
          Response = MsgBox("There already exists a version with today's date.  Do you want to continue?", 4, "Version Add")
          If Response = 7 Then
            GoTo Exit_ButtonAdd_Click
          End If
        End If
        VersionsByDate.Close
        
        strSQL = "SELECT * FROM [tbl_master_version] ORDER BY [version_key_number]"
        Set Versions = db.OpenRecordset(strSQL)
        Versions.MoveLast  ' Move to current version - max version key number
        Me![project_ID] = Versions![project_ID]
        intFID = Versions.RecordCount   ' Set new key
        Me![version_key_number] = intFID + 1
        Me![version_key_date] = Now()     ' Today's date
        Me![narrative_version] = Versions![narrative_version] ' Fill new record from last version
        DoCmd.RunCommand acCmdSaveRecord ' Save parent record so child records can be added
        Versions.Close
        
        ' Add new child records for SOPs
        strSQL = "SELECT * FROM [tbl_SOP_version] WHERE [version_key_number] = " & intFID
        Set CurrentSOP = db.OpenRecordset(strSQL) ' Open recordset for current SOPs
        Set NewSOP = db.OpenRecordset("tbl_SOP_version") ' Open recordset for new SOP records
        While CurrentSOP.EOF = False
            NewSOP.AddNew
            NewSOP![version_key_number] = intFID + 1
            NewSOP![SOP_number] = CurrentSOP![SOP_number]
            NewSOP![SOP_version_number] = CurrentSOP![SOP_version_number]
            NewSOP![active_flag] = CurrentSOP![active_flag]
            NewSOP.Update
            CurrentSOP.MoveNext
        Wend
        CurrentSOP.Close
        NewSOP.Close
        Me!Subform_Versions.Form.AllowAdditions = True  ' Its a new version, so we
        Me!Subform_Versions.Form.AllowEdits = True      ' need to allow updates.
        Me!narrative_version.Locked = False
        Me!project_ID.Locked = True  ' Dont let them change the project ID on new records.
        Me.Requery ' Refresh form SOPs from table
        Me.Filter = ""  ' Clear filter so new record will display.

Exit_ButtonAdd_Click:
    Exit Sub

Err_ButtonAdd_Click:
    MsgBox Err.Description
    Resume Exit_ButtonAdd_Click
    
End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
