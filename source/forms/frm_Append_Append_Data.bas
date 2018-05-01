Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =1080
    Top =-225
    Right =12435
    Bottom =11160
    DatasheetGridlinesColor =12632256
    OrderBy ="Append_Order"
    RecSrcDt = Begin
        0x79f24c49d2b3e340
    End
    RecordSource ="tsys_Append_Tables"
    Caption ="Append Data"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
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
            BorderLineStyle =0
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
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =3600
            BackColor =5394044
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Width =2520
                    Height =480
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label2"
                    Caption ="Append Data"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8040
                    Top =120
                    Width =1320
                    Height =600
                    FontWeight =700
                    Name ="cmd_AppendLog"
                    Caption ="View Append Log"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =120
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =720
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    Enabled = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =2880
                    Left =5715
                    Top =2880
                    Width =3360
                    ColumnOrder =2
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cmbo_Select_Event"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Events.Event_ID, [Plot_Name] & \" \" & \" \" & [Event_Date] AS PickSt"
                        "ring FROM tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID=tbl_E"
                        "vents.Location_ID WHERE (((Year([Event_Date]))=Year(Now()))) ORDER BY tbl_Events"
                        ".Event_Date DESC; "
                    ColumnWidths ="0;2880"
                    LayoutCachedLeft =5715
                    LayoutCachedTop =2880
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =3120
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =660
                            Top =2880
                            Width =4980
                            Height =240
                            FontWeight =700
                            ForeColor =9868950
                            Name ="cmbo_Event_LU_Label"
                            Caption ="Select the Event to Append  to in Master Database -->"
                            ControlTipText ="Select the Event from the main data set that you wish to append the secondary ta"
                                "blet  data to"
                            LayoutCachedLeft =660
                            LayoutCachedTop =2880
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =3120
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5715
                    Top =2280
                    Width =3360
                    ColumnOrder =0
                    TabIndex =1
                    Name ="cmbo_Select_Import_Event_Table"
                    RowSourceType ="Value List"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =5715
                    LayoutCachedTop =2280
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =2520
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =660
                            Top =2280
                            Width =4980
                            Height =240
                            FontWeight =700
                            ForeColor =9868950
                            Name ="Label22"
                            Caption ="Select Events Table to import from Secondary Tablet -->"
                            LayoutCachedLeft =660
                            LayoutCachedTop =2280
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =2520
                        End
                    End
                End
                Begin ComboBox
                    Enabled = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5715
                    Top =2580
                    Width =3360
                    ColumnOrder =1
                    TabIndex =2
                    Name ="cmbo_Select_Import_Events"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;1440"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =5715
                    LayoutCachedTop =2580
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =660
                            Top =2580
                            Width =4980
                            Height =240
                            FontWeight =700
                            ForeColor =9868950
                            Name ="Label24"
                            Caption ="Select the Event to Append from Secondary Tablet -->"
                            LayoutCachedLeft =660
                            LayoutCachedTop =2580
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =2820
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8040
                    Top =780
                    Width =1320
                    Height =600
                    FontWeight =700
                    TabIndex =4
                    Name ="cmd_ViewUpdateLog"
                    Caption ="View Update Log"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =780
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1380
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =540
                    Top =600
                    Width =6420
                    Height =725
                    ColumnOrder =3
                    TabIndex =5
                    Name ="optframe_Step1Append"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =540
                    LayoutCachedTop =600
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =1325
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =215
                            Left =660
                            Top =480
                            Width =1440
                            Height =420
                            FontSize =16
                            FontWeight =700
                            BackColor =5394044
                            ForeColor =8454143
                            Name ="Label29"
                            Caption ="Step 1"
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =4620
                            Top =900
                            Width =1140
                            Height =390
                            FontWeight =700
                            OptionValue =1
                            Name ="Toggle31"
                            Caption ="One Tablet"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =5760
                            Top =900
                            Width =1140
                            Height =390
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            Name ="Toggle32"
                            Caption ="Two Tablets"

                            LayoutCachedLeft =5760
                            LayoutCachedTop =900
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =1290
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =600
                    Top =960
                    Width =3960
                    Height =240
                    FontWeight =700
                    ForeColor =8454143
                    Name ="Label33"
                    Caption ="How many tablets were the data collected on?"
                End
                Begin OptionGroup
                    Enabled = NotDefault
                    OverlapFlags =255
                    Left =540
                    Top =1560
                    Width =8817
                    Height =1920
                    ColumnOrder =4
                    TabIndex =6
                    Name ="optframe_Step2Append"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =540
                    LayoutCachedTop =1560
                    LayoutCachedWidth =9357
                    LayoutCachedHeight =3480
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =247
                            Left =660
                            Top =1440
                            Width =1500
                            Height =360
                            FontSize =16
                            FontWeight =700
                            BackColor =5394044
                            ForeColor =8454143
                            Name ="Label35"
                            Caption ="Step 2"
                        End
                        Begin ToggleButton
                            OverlapFlags =127
                            Left =4620
                            Top =1800
                            Width =1140
                            Height =390
                            FontWeight =700
                            OptionValue =1
                            Name ="Toggle37"
                            Caption ="Tablet One"

                            LayoutCachedLeft =4620
                            LayoutCachedTop =1800
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =2190
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                        Begin ToggleButton
                            OverlapFlags =119
                            Left =5760
                            Top =1800
                            Width =1140
                            Height =390
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            Name ="Toggle38"
                            Caption ="Tablet Two"

                            LayoutCachedLeft =5760
                            LayoutCachedTop =1800
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =2190
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =600
                    Top =1860
                    Width =3720
                    Height =240
                    FontWeight =700
                    ForeColor =8454143
                    Name ="Label39"
                    Caption ="On which tablet was this data collected?"
                    LayoutCachedLeft =600
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2100
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    Left =660
                    Top =3180
                    Width =7020
                    Height =240
                    FontWeight =700
                    ForeColor =9868950
                    Name ="Lbl_Step2_Finish"
                    Caption ="Click 'Append Data' below and then repeat Step 2 for each event to be appended"
                    ControlTipText ="Select the Event from the main data set that you wish to append the secondary ta"
                        "blet  data to"
                    LayoutCachedLeft =660
                    LayoutCachedTop =3180
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3420
                End
            End
        End
        Begin Section
            Height =360
            BackColor =15527148
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1245
                    Top =60
                    Width =2535
                    ColumnWidth =4260
                    Name ="txt_Table_Name"
                    ControlSource ="Table_Name"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =60
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =60
                            Width =960
                            Height =240
                            Name ="Label0"
                            Caption ="Table Name:"
                            LayoutCachedLeft =165
                            LayoutCachedTop =60
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =3840
                    Top =60
                    TabIndex =1
                    Name ="chk_Append"
                    ControlSource ="Append"
                    DefaultValue ="0"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =300
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =5760
                    Left =5580
                    Top =60
                    Width =4380
                    TabIndex =2
                    Name ="cmbo_Append_Table"
                    ControlSource ="Append_Table"
                    RowSourceType ="Value List"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =4080
                            Top =60
                            Width =1440
                            Height =240
                            Name ="Label5"
                            Caption ="Append data from:"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =60
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =300
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =600
            BackColor =5394044
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6240
                    Top =120
                    Width =1620
                    FontWeight =700
                    Name ="cmd_Append_Event_Data"
                    Caption ="Append Data"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8039
                    Top =120
                    Width =1319
                    FontWeight =700
                    TabIndex =1
                    Name ="cmd_Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8039
                    LayoutCachedTop =120
                    LayoutCachedWidth =9358
                    LayoutCachedHeight =480
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub cmbo_Append_Table_GotFocus()
Dim db As DAO.Database
Dim tdef As TableDef

Set db = CurrentDb
Me!cmbo_Append_Table.RowSource = ""
For Each tdef In db.TableDefs
    Dim iTableName As Long
        iTableName = Len(Me!txt_Table_Name.Value)
    Dim strTableName As String
        strTableName = Me!txt_Table_Name.Value
    Dim strAppTableName As String
   
    If Left(tdef.Name, 1) = "_" Then
        strAppTableName = Right(Left(tdef.Name, iTableName + 1), iTableName)
        If strAppTableName = strTableName Then
            Me!cmbo_Append_Table.AddItem tdef.Name
        End If
    Else
        GoTo NextRecord:
    End If
NextRecord:
Next
Me!cmbo_Append_Table.Requery

'Cleanup
Set db = Nothing
Set tdef = Nothing

End Sub

Private Sub cmbo_Select_Import_Event_Table_AfterUpdate()

Dim strTableName As String
Dim EventSQL As String

If Me!cmbo_Select_Import_Event_Table = "" Or IsNull(Me!cmbo_Select_Import_Event_Table) Then
    Exit Sub

Else
    strTableName = Me!cmbo_Select_Import_Event_Table.Value
    
    EventSQL = "SELECT [" & strTableName & "].Event_ID, [" & strTableName & "].Location_ID, " _
        & "[tbl_Locations].[Plot_Name] &" & """  """ & "& [" & strTableName & "].[Event_Date] " _
        & "AS [Pick String] " _
        & "FROM [" & strTableName & "] " _
        & "LEFT JOIN tbl_Locations " _
        & "ON [" & strTableName & "].Location_ID = tbl_Locations.Location_ID " _
        & "ORDER BY Event_Date DESC;"
       
       'MsgBox EventSQL
      
    Me!cmbo_Select_Import_Events.RowSource = EventSQL
           
End If
End Sub

Private Sub cmbo_Select_Import_Event_Table_GotFocus()

Dim db As DAO.Database
Dim tdef As TableDef

Set db = CurrentDb

Me!cmbo_Select_Import_Event_Table.RowSource = ""

For Each tdef In db.TableDefs
    If Left(tdef.Name, 11) = "_tbl_Events" Then
        Me!cmbo_Select_Import_Event_Table.AddItem tdef.Name
    Else
        GoTo NextRecord:
    End If
NextRecord:
Next
Me!cmbo_Select_Import_Event_Table.Requery
Set db = Nothing
Set tdef = Nothing

End Sub

Private Sub cmbo_Select_Import_Events_GotFocus()
    Me!cmbo_Select_Import_Events.Requery
End Sub

Private Sub cmd_Append_Event_Data_Click()

Dim db As DAO.Database
Set db = CurrentDb

Dim rsMain As DAO.Recordset  'rsMain is the master dataset in the database to which data is being appended
Dim rsAppend As DAO.Recordset 'rsAppend is the dataset with new records to be appended to the
Dim rsForm As DAO.Recordset 'the recordset for the append form.
'Dim rsAppendLog As DAO.Recordset 'the recordset that stores the information about the records appended to each table

Dim strMain As String
Dim strAppend As String

'This is required
DoCmd.RunCommand acCmdSaveRecord

Set rsForm = Me.RecordsetClone

rsForm.MoveFirst

Do Until rsForm.EOF
    
    'Cycle through the tables to see which ones have been chosen
    'to have new data appended.
    If rsForm![Append] = True Then
    
        'rsMain is the main dataset in the database
         strMain = rsForm![Table_Name]
        
        Set rsMain = db.OpenRecordset(strMain)
        
            'Get the length of the table name to check and make sure that the Main Table and the Append Table names match
                Dim iLength As Long
                Dim iLength2 As Long
                iLength = Len(rsForm![Table_Name])
                iLength2 = iLength + 1
                
                'Root of the Append Table name
                Dim strAppTableName As String
                strAppTableName = Right(Left(rsForm![Append_Table], iLength2), iLength)
                              
            'Check to make sure that an Append Table is specified if the Append box is checked
            If rsForm![Append_Table] = "" Or IsNull(rsForm![Append_Table]) Then
                MsgBox "Make sure that you have properly selected all of the data you wish to append!", vbCritical, "Append Data"
                Exit Sub
            'Check to make sure that the Append Table Name matches the Main Table Name
            ElseIf strAppTableName <> rsForm![Table_Name] Then
                MsgBox "Make sure you have properly selected the data set to append to " & rsForm![Table_Name] & ".", vbCritical, "Append Data"
                Exit Sub
            End If
        
        strAppend = rsForm![Append_Table]
        Dim strAppendTableName As String
        strAppendTableName = strAppend
        
        'Capture the imported events table to use when appending new tree and sapling data.
        If strMain = "tbl_Events" Then
            Dim rsEvents As DAO.Recordset
            Set rsEvents = db.OpenRecordset(strAppend)
        End If
        
        'Check to see if the table is tbl_Locations. If so, send it to a special functionto update the locations table with newly collected slope and aspect.
        If strMain = "tbl_Locations" Then
            Dim rsLoc As DAO.Recordset
            Set rsLoc = db.OpenRecordset(strAppend)
            'send it to a specail function to check to see if anything needs updating. If so, update it and return.
            fxnUpdateLocInfo rsLoc, strAppend
                        
        End If
        
        
    'Determine if the data being appended is for tbl_Tags.
    'If it is send it to a special function to update the data in these tables tables prior to appending new data.
    
        If strMain = "tbl_Tags" Then
            Set rsAppend = db.OpenRecordset(strAppend)
            UpdateTags rsMain, rsAppend, rsEvents, strAppendTableName
            GoTo NextRecord:
        End If
        
 'If you are appending records to an existing Event_ID:
 'First figure out if you are appending data from the Main Tablet or Secondary Tablet
        
        'If it is from the secondary tablet run it through this code to replace Event_IDs
        If Me!optframe_Step1Append.Value = 2 Then
            If Me!optframe_Step2Append.Value = 2 Then
                If Me!cmbo_Select_Event = "" Or IsNull(Me!cmbo_Select_Event) Then
                    MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
                    Me!cmbo_Select_Event.SetFocus
                    Exit Sub
                ElseIf Me!cmbo_Select_Import_Event_Table = "" Or IsNull(Me!cmbo_Select_Import_Event_Table) Then
                        MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
                        Me!cmbo_Select_Import_Event_Table.SetFocus
                        Exit Sub
                ElseIf Me!cmbo_Select_Import_Events = "" Or IsNull(Me!cmbo_Select_Import_Events) Then
                            MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
                            Me!cmbo_Select_Import_Events.SetFocus
                            Exit Sub
                End If
                        
    'Declare and set the variables for the Event ID's in both the Main (master) dataset
    'as well as in the imported data set
    
                Dim GUIDMain As String
                Dim GUIDReplace As String
                
                GUIDMain = Me!cmbo_Select_Event.Column(0)
                GUIDReplace = Me!cmbo_Select_Import_Events.Column(0)
    
 'Check to see if the table contains an Event_ID field
 
                Dim boolEvent As Boolean
                boolEvent = False
            
                Dim tdef As DAO.TableDef
                Dim lCount As Long
                Dim lCtr As Long
                Dim strFieldName As String
                Set tdef = db.TableDefs(strAppend)
        
                With tdef
                    lCount = .Fields.Count
                        For lCtr = 0 To lCount - 1
                            strFieldName = .Fields(lCtr).Name
                                              
                            If strFieldName = "Event_ID" Then
                                boolEvent = True
                            End If
         
                        Next
                End With
 
 'if the table contains an Event_ID pass the data set to the update event id function
 
                If boolEvent = True Then
   
    'Pass the appending data set, the master GUID as well as the GUID that needs to be
    'replaced to the function to update the Event ID to the Master Event ID
                    Dim strUpdateEventIDSQL As String
                    Dim qdefUpdateEventID As QueryDef
                    Dim strTableName As String
                    Dim strFindUpdate As String
                
                    strFindUpdate = GUIDReplace
                    strTableName = Me!cmbo_Append_Table.Value
                
               'Only select those records where the EventID = the GUID to be replaced
                
                    strUpdateEventIDSQL = "SELECT [" & strAppend & "].*" _
                    & "FROM [" & strAppend & "] " _
                    & "WHERE ((([" & strAppend & "].Event_ID) = " & strFindUpdate & "));"
                  
               ' MsgBox strUpdateEventIDSQL
                
                'save the SQL as a qdef and then create a recordset to pass to the UPDATE EVENT ID Function
                    Set qdefUpdateEventID = db.CreateQueryDef("_Qry_UpdateEventID", strUpdateEventIDSQL)
                    Set rsAppend = db.OpenRecordset("_Qry_UpdateEventID")
                
                    If rsAppend.RecordCount > 0 Then
                        UpdateEventID rsAppend, GUIDMain, GUIDReplace, strTableName
                    End If
                'Delete the Select Query from the database after it has been used in the update event function
                    DoCmd.DeleteObject acQuery, qdefUpdateEventID.Name
    
    'Had to insert this code because when the record set was passed to the Append
    'function it was attempting to append too many records.
    'If events already exist additional information must be appended on a
    'Event by Event basis.
        
                    Dim strAppendSQL As String
                    Dim qdefAppend As QueryDef
        
        'main event id
                    Dim strFind As String
                    strFind = GUIDMain
                    
        'query only the records that were collected on this event
        
                    strAppendSQL = "SELECT [" & strAppend & "].*" _
                    & "FROM [" & strAppend & "] " _
                    & "WHERE ((([" & strAppend & "].Event_ID) = " & strFind & "));"
                
                    Set qdefAppend = db.CreateQueryDef("_Qry_AppendEventRecs", strAppendSQL)
      
        'reset the rsAppend variable to equal the query that only contains records from
        'targeted event
        
                    Set rsAppend = db.OpenRecordset("_Qry_AppendEventRecs")
       ' MsgBox rsAppend.Name
        
        'pass the event specific recordset to the append function
        
                    AppendtoTable rsAppend, rsMain, strAppendTableName
                
              'Delete the select query once the append function has been completed.
                
                    DoCmd.DeleteObject acQuery, qdefAppend.Name
                
                Else
                
                    GoTo AppendData:
                
                End If
            
            Else
                
                GoTo AppendData:
                
            End If
            
        Else
        
 'skip all of that crap about updating the event id and just append the damn records

AppendData:

 'We want to make sure to select only those records that are associated witht the events being imported.
  'skip this query if the current append table is tbl_Events
 
 If strAppTableName <> "tbl_Events" Then
 
 'Check to see if the table contains an Event_ID field
 
                'Dim boolEvent As Boolean
                boolEvent = False
            
                'Dim tdef As DAO.TableDef
                'Dim lCount As Long
                'Dim lCtr As Long
                'Dim strFieldName As String
                Set tdef = db.TableDefs(strAppend)
        
                With tdef
                    lCount = .Fields.Count
                        For lCtr = 0 To lCount - 1
                            strFieldName = .Fields(lCtr).Name
                                              
                            If strFieldName = "Event_ID" Then
                                boolEvent = True
                            End If
         
                        Next
                End With
        'If the data table has an Event_ID field, run it through the query that only selects data collected on the imported events. _
        If it does not have an Event_ID field, send it through the standard Append function for now.
        
        If boolEvent = True Then
                
                Dim strEvents As String
                strEvents = rsEvents.Name
     
                strSQL_FindImportedRecs = "SELECT [" & strAppend & "].* FROM [" & strEvents & "] INNER JOIN [" & strAppend & "] " _
                                        & "ON [" & strEvents & "].Event_ID = [" & strAppend & "].Event_ID;"
                                        
'Turn the SQL statement into a query to be used in the following append fuctions as the Append data set
           
                Dim qdef_NewEventRecs As QueryDef
                Set qdef_NewEventRecs = db.CreateQueryDef("_qry_NewEventData", strSQL_FindImportedRecs)
                
                Set rsAppend = db.OpenRecordset("_qry_NewEventData")
                
                'Send the two recordsets (rsMain and rsAppend) to the Append Function
                
                AppendtoTable rsAppend, rsMain, strAppendTableName
    
                'delete the query def so that it can be recreated as the code loops
                
                DoCmd.DeleteObject acQuery, qdef_NewEventRecs.Name
        Else
        
   'Run the append function for any table that does not have an Event_ID
        Set rsAppend = db.OpenRecordset(strAppend)
    
        AppendtoTable rsAppend, rsMain, strAppendTableName
        
        End If
        
Else

'Run the Append function for tbl_Events
                 
    Set rsAppend = db.OpenRecordset(strAppend)
    
   
    AppendtoTable rsAppend, rsMain, strAppendTableName
    
End If
          
End If
'Make sure that the Events table is checked even if you do not need to append any data to it.
    
ElseIf rsForm![Table_Name] = "tbl_Events" And rsForm![Append] = False Then
                MsgBox "The Events table needs to be included in the append operation." & vbNewLine & vbNewLine & _
                "Please go back and check the append box and select an imported events table.", , "Append Data"
                Exit Sub
Else
        
        GoTo NextRecord:
    
End If

NextRecord:

rsForm.MoveNext

Loop

MsgBox "Update and Appending complete!", , "Update and Append Data"

CleanUp:

Set rsAppend = Nothing
Set rsMain = Nothing
Set rsAppendLog = Nothing
Set rsEvents = Nothing

Set db = Nothing
Set rsForm = Nothing
Set qdefAppend = Nothing
Set qdefUpdateEventID = Nothing
Set qdef_NewEventRecs = Nothing

End Sub

Private Sub cmd_ViewUpdateLog_Click()
On Error GoTo Err_cmd_ViewUpdateLog_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Update_Log"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_cmd_ViewUpdateLog_Click:
    Exit Sub

Err_cmd_ViewUpdateLog_Click:
    MsgBox Err.Description
    Resume Exit_cmd_ViewUpdateLog_Click
    
End Sub

Private Sub Detail_Click()

If Me!optframe_Step2Append.Value = 2 Then
    If Me!cmbo_Select_Event = "" Then
        MsgBox "You must complete the necessary information above.", , "Append Data"
        Me!cmbo_Select_Event.SetFocus
    ElseIf Me!cmbo_Select_Import_Event_Table.Value = "" Then
        MsgBox "You must complete the necessary information above.", , "Append Data"
        Me!cmbo_Select_Import_Event_Table.SetFocus
    ElseIf Me!cmbo_Select_Import_Events.Value = "" Then
        MsgBox "You must complete the necessary information above.", , "Append Data"
        Me!cmbo_Select_Import_Events.SetFocus
    End If
End If

End Sub

Private Sub Form_Close()

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim ctrlCombo As ComboBox

Set db = CurrentDb
Set rs = db.OpenRecordset("tsys_Append_Tables")
Set ctrlCombo = Me!cmbo_Append_Table
Me!cmbo_Append_Table.RowSource = " "

rs.MoveFirst

Do While Not rs.EOF
    rs.Edit
    rs![Append_Table] = ""
    rs.Update
rs.MoveNext
Loop

Set db = Nothing
Set rs = Nothing

End Sub

Private Sub Form_Load()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim tdef As TableDef

Set db = CurrentDb

Set rs = Me.RecordsetClone
 
rs.MoveFirst
Do While Not rs.EOF

DoCmd.RunCommand acCmdSaveRecord
'Until all of the Slope and Aspect data are updated in the master locations table we want to update the locations
'table with the slope and aspect data collected in the field on the field data bases.

'    If Me!txt_Table_Name = "tbl_Locations" Then
'        Me!chk_Append.Value = 0
'    Else
'        Me!chk_Append.Value = 1
'    End If
'**************************************************
DoCmd.RunCommand acCmdSaveRecord
        For Each tdef In db.TableDefs
            Dim iTableName As Long
            'iTableName = Len(rs![Table_Name]) '(Me!txt_Table_Name.Value)

            Dim strTableName As String
            strTableName = rs![Table_Name]
            iTableName = Len(strTableName)
            Dim strAppTableName As String

            If Left(tdef.Name, 1) = "_" Then
                strAppTableName = Right(Left(tdef.Name, iTableName + 1), iTableName)
                
                'If it is tbl_Events make sure we are grabbing the events table from the primary tablet.
                 If strAppTableName = "tbl_Events" Then
                    If Right(tdef.Name, 9) = "SECONDARY" Then
                        GoTo NextRecord:
                    End If
                 End If
                 If strAppTableName = strTableName Then
                    rs.Edit
                    rs![Append_Table] = tdef.Name
                    rs.Update
                  End If
                
            Else
                GoTo NextRecord:

            End If

NextRecord:
        Next

rs.MoveNext
Loop
 
Me.OrderBy = "Append_Order ASC"
Me.OrderByOn = True

Me.optframe_Step1Append.SetFocus
Me.optframe_Step2Append.Enabled = False

Set db = Nothing
Set rs = Nothing
Set tdef = Nothing
 
End Sub

Private Sub cmd_Close_Click()
On Error GoTo Err_cmd_Close_Click

Dim strResponse As String

strResponse = MsgBox("Would you like to delete any of the imported tables?", vbYesNoCancel, "Delete Tables?")
    If strResponse = vbYes Then
        DoCmd.Close
        DoCmd.OpenForm "frm_Append_Delete_Tables"
    ElseIf strResponse = vbCancel Then
        Exit Sub
    Else
        DoCmd.Close
    End If
        
Exit_cmd_Close_Click:
    Exit Sub
Err_cmd_Close_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Close_Click
End Sub

Private Sub cmd_AppendLog_Click()
On Error GoTo Err_cmd_AppendLog_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Log"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_cmd_AppendLog_Click:
    Exit Sub
Err_cmd_AppendLog_Click:
    MsgBox Err.Description
    Resume Exit_cmd_AppendLog_Click
End Sub

Private Sub opt_frame_Select_Append_AfterUpdate()

If Me!opt_frame_Select_Append.Value = 1 Then
    
    Me!cmbo_Select_Event.Enabled = False
    Me!cmbo_Select_Import_Event_Table.Enabled = False
    Me!cmbo_Select_Import_Events.Enabled = False
    Me!Lbl_Step2_Finish.Visible = False
    
    Me.RecordSource = "tsys_Append_Tables"
    Me!cmbo_Select_Event = ""
    Me!cmbo_Select_Import_Event_Table = ""
    Me!cmbo_Select_Import_Events = ""
    
ElseIf Me!opt_frame_Select_Append.Value = 2 Then

    Me!cmbo_Select_Event.Enabled = True
    Me!cmbo_Select_Import_Event_Table.Enabled = True
    Me!cmbo_Select_Import_Events.Enabled = True
    Me!Lbl_Step2_Finish.Visible = True
    
    Me.RecordSource = "qry_Append"
    Me!cmbo_Select_Event = ""
    Me!cmbo_Select_Import_Event_Table = ""
    Me!cmbo_Select_Import_Events = ""
        
End If

End Sub

Private Sub optframe_Step1Append_AfterUpdate()

Select Case optframe_Step1Append.Value
Case 1
    Me.Detail.Visible = True
    
     Me!optframe_Step2Append.Value = 0
     optframe_Step2Append.Enabled = False
     
     Me!cmbo_Select_Event.Enabled = False
     Me!cmbo_Select_Import_Event_Table.Enabled = False
     Me!cmbo_Select_Import_Events.Enabled = False
     Me!Lbl_Step2_Finish.Visible = False
    
     Me.RecordSource = "tsys_Append_Tables"
     Me!cmbo_Select_Event = ""
     Me!cmbo_Select_Import_Event_Table = ""
     Me!cmbo_Select_Import_Events = ""
     
     'Order the append tables in the proper order so that there are no errors during the append sequence
     Me.OrderBy = "Append_Order ASC"
     Me.OrderByOn = True
     
     Me!cmd_Append_Event_Data.Enabled = True
     
     Case 2
    
    optframe_Step2Append.Enabled = True
    Me!optframe_Step2Append.Value = 0
    Me!optframe_Step2Append.SetFocus
    Me.Detail.Visible = False
    
    Me!cmd_Append_Event_Data.Enabled = False
    
Case Else
    optframe_Step2Import.Enabled = False
    Me!cmd_Append_Event_Data.Enabled = False
    
End Select
End Sub

Private Sub optframe_Step2Append_AfterUpdate()
Me.Detail.Visible = True

If Me!optframe_Step2Append.Value = 1 Then
  
    Me.RecordSource = "qry_Append_Primary_Tablet_Append"
    
    Me.OrderBy = "Append_Order ASC"
    Me.OrderByOn = True
    
    Me!cmbo_Append_Table.SetFocus
    
    Me!cmd_Append_Event_Data.Enabled = True
  
ElseIf Me!optframe_Step2Append.Value = 2 Then
       
    Me!cmbo_Select_Event.Enabled = True
    Me!cmbo_Select_Import_Event_Table.Enabled = True
    Me!cmbo_Select_Import_Events.Enabled = True
    Me!Lbl_Step2_Finish.Visible = True
    
    Me.RecordSource = "qry_Append_Secondary_Tablet_Append"
    Me!cmbo_Select_Event = ""
    Me!cmbo_Select_Import_Event_Table = ""
    Me!cmbo_Select_Import_Events = ""

    Me!cmbo_Select_Import_Event_Table.SetFocus
    
    Me!cmd_Append_Event_Data.Enabled = True
End If
End Sub
