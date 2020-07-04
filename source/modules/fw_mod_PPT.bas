Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_PPT
' Level:        Framework module
' Version:      1.01
' Description:  Framework-wide related powerpoint values, functions & subroutines
'
' Source/date:  Bonnie Campbell, January 2, 2018 for NCPN tools
' Revisions:    BLC, 1/2/2018 - 1.00 - initial version
'               BLC, 5/16/2019 - 1.01 - added fw_ module prefix
' =================================

'-----------------------------------------------------------------------
' Declarations
'-----------------------------------------------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

'-----------------------------------------------------------------------
' Functions
'-----------------------------------------------------------------------

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' Sub:          XX
' Description:  Uploads data into database from CSV file
' Assumptions:  -
' Parameters:   strFilename - name of file being uploaded (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/2/2018 - initial version
' ---------------------------------
Public Sub xx(strFileName As String)
On Error GoTo Err_Handler

' FIX: 1/23/2019 - comment out (no Office 2016 PPT)
'    Dim newPowerPoint As PowerPoint.Application
'    Dim pptPres As PowerPoint.Presentation
'    Dim activeSlide As PowerPoint.Slide
'    'Dim cht As Excel.ChartObject '<< requires Excel reference
'
'    Dim file As String
'    file = "C:\Users\jbain\Documents\PowerPoint template_Span.pptx"
'
'    Dim pptcht As PowerPoint.Chart
'
'    'Look for existing instance
'    'On Error Resume Next
''    Set newPowerPoint = GetObject(, "PowerPoint.Application")
'    On Error GoTo 0
'
'    'Let's create a new PowerPoint
'    If newPowerPoint Is Nothing Then
'        Set newPowerPoint = New PowerPoint.Application
'    End If
'
'    'Make a presentation in PowerPoint
'    If newPowerPoint.Presentations.Count = 0 Then
'        Set pptPres = newPowerPoint.Presentations.Open(file)
'    End If
'
'    'Show the PowerPoint
'    newPowerPoint.visible = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - XX[fw_mod_PPT])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          FetchPPTTemplate
' Description:  Retrieves ppt template file
' Assumptions:  -
' Parameters:   strTemplate - name of desired PPT template (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 2, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/2/2018 - initial version
' ---------------------------------
Public Sub FetchPPTTemplate(strTemplate As String)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FetchPPTTemplate[fw_mod_PPT])"
    End Select
    Resume Exit_Handler
End Sub