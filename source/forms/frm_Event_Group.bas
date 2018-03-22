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
    ItemSuffix =19
    Left =2310
    Top =2715
    Right =12555
    Bottom =5580
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3244ec772e09e340
    End
    RecordSource ="tbl_Event_Group"
    Caption =" Event Groups"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Section
            Height =3240
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
                    Left =1086
                    Top =120
                    Width =8748
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtEvent_Group_ID"
                    ControlSource ="Event_Group_ID"
                    StatusBarText ="M. An identifier for the event group (Ev_Gp_ID)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =120
                            Width =210
                            Height =255
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblLocation_ID"
                            Caption ="ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Line
                    OverlapFlags =85
                    Left =120
                    Top =480
                    Width =9720
                    Name ="Line1"
                End
                Begin Line
                    OverlapFlags =85
                    Left =120
                    Top =960
                    Width =9720
                    Name ="Line2"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3540
                    Top =2820
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
                    Top =2700
                    Width =9720
                    Name ="Line6"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =600
                    ColumnWidth =2205
                    TabIndex =1
                    Name ="txtStart_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Ending date of the event group (Start_Date)"
                    Tag ="<data>"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =600
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="Start Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3600
                    Top =600
                    TabIndex =2
                    Name ="txtEnd_Date"
                    ControlSource ="End_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Ending date of the event group (End_Date)"
                    Tag ="<data>"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =600
                            Width =780
                            Height =240
                            Name ="Label10"
                            Caption ="End Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6780
                    Top =600
                    Width =3060
                    TabIndex =3
                    Name ="txtEvent_Group_Name"
                    ControlSource ="Event_Group_Name"
                    StatusBarText ="MA. Event group  (e.g. season, trip) name (Ev_Gp_Name)"
                    Tag ="<data>"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5220
                            Top =600
                            Width =1425
                            Height =240
                            Name ="Label11"
                            Caption ="Event Group Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =1080
                    Width =8760
                    TabIndex =4
                    Name ="txtEvent_Group_Desc"
                    ControlSource ="Event_Group_Desc"
                    StatusBarText ="MA. Event group description (Ev_Gp_Desc)"
                    Tag ="<data>"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1080
                            Width =870
                            Height =240
                            Name ="Label14"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =1440
                    Width =8760
                    Height =783
                    TabIndex =5
                    Name ="txtEvent_Group_Notes"
                    ControlSource ="Event_Group_Notes"
                    StatusBarText ="MA. Event group notes (Ev_Gp_Note)"
                    Tag ="<data>"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1440
                            Width =495
                            Height =240
                            Name ="Label15"
                            Caption ="Notes"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5100
                    Top =2820
                    Height =300
                    FontWeight =700
                    TabIndex =8
                    Name ="cmdDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =2340
                    Width =8760
                    ColumnWidth =4200
                    TabIndex =6
                    Name ="txtEvent_Group_Report"
                    ControlSource ="Event_Group_Report"
                    StatusBarText ="MA. Trip report, link to trip report or trip report name (Ev_Gp_Rept)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2340
                            Width =630
                            Height =240
                            Name ="Label18"
                            Caption ="Report"
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
' Description:  Event groups entry form
' Data source:  tbl_Event_Group
' Data access:  edit, add, delete
' Pages:        none
' Functions:    none
' References:   fxnGUIDGen
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Error_Handler

DoCmd.RunCommand acCmdDeleteRecord
DoCmd.Close acForm, Me.Name

MsgBox "Record deleted successfully", , "Record Deleted"

Exit_Handler:
    Exit Sub

Error_Handler:
    Select Case Err.Number
        Case 2046 'command not available
            MsgBox "Unable to delete record.", vbExclamation, "Cannot Delete Record"
            Resume Exit_Handler
        Case 2501 'user canceled delete
            MsgBox "Delete canceled", , "Delete Canceled"
            Resume Exit_Handler
        Case 3200 'related records
            MsgBox "There are related records that prevent this record from being deleted.  Delete all related records first and then delete this record.", vbInformation, "Cannot Delete Record"
            Resume Exit_Handler
        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error - Form: " & Me.Name & " - cmdDelete_Click"
            Resume Exit_Handler
    End Select

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'check to see if a primary key is needed and add it (used for string GUIDs)
If fxnFormCheck(Me) Then
    If Me.NewRecord Then
        If GetDataType("tbl_Event_Group", "Event_Group_ID") = dbText Then
            Me!txtEvent_Group_ID = fxnGUIDGen
        End If
    End If
Else
    Cancel = True
End If

End Sub

Private Sub Form_Close()
'update control on calling form to reflect new Event Group values
fxnUpdateControl Me.OpenArgs
End Sub
