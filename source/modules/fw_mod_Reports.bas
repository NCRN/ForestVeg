Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_Reports
' Level:        Framework module
' Version:      1.03
' Description:  generic report functions & procedures
'
' Source/date:  Bonnie Campbell, 5/25/2016
' Revisions:    BLC - 5/25/2016 - 1.00 - initial version
'               BLC - 6/24/2016 - 1.01 - replaced Exit_Function > Exit_Handler
'               BLC - 10/6/2017 - 1.02 - added ReportIsLoaded() from mod_UI
'               BLC - 5/16/2019 - 1.03 - added fw_ module prefix
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------
' ---------------------------------
'  Reports
' ---------------------------------

' =================================
' FUNCTION:     ReportIsLoaded
' Description:  Returns whether the specified report is loaded
' Parameters:   strReportName - string for the name of the report to check
' Returns:      True if the specified report is open, False if not
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell - 5/17/2015 - for NCPN tools
' Revisions:    BLC, 5/17/2015 - initial version
' =================================
Public Function ReportIsLoaded(ByVal strReportName As String) As Boolean
On Error GoTo Err_Handler
 
    ' Possible states returned by SysCmd & CurrentView
    Const cObjStateClosed = 0
    Const cDesignView = 0
    Const cPrintView = 5
    Const cReportView = 6
    Const cLayoutView = 7

    ' check current state - not open or nonexistent, design, print, layout, or report view
    If SysCmd(acSysCmdGetObjectState, acReport, strReportName) <> cObjStateClosed Then
        ' check current view, return True if open and not in design view
        If Reports(strReportName).CurrentView <> cDesignView Then
            ReportIsLoaded = True
        End If
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReportIsLoaded[fw_mod_Reports])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     NoData
' Description:  report actions when no data is found
' Assumptions:  -
' Parameters:   rpt - report being referenced
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2015 - initial version
' ---------------------------------
Public Function NoData(rpt As Report)
On Error GoTo Err_Handler

    'Purpose: Called by report's NoData event.
    'Usage: =NoData([Report])
    Dim strCaption As String   'Caption of report.
    
    strCaption = rpt.Caption
    If strCaption = vbNullString Then
        strCaption = rpt.Name
    End If
    
'    DoCmd.CancelEvent
    MsgBox "There are no records to include in report """ & _
        strCaption & """.", vbInformation, "No Data..."


Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NoData[fw_mod_Reports])"
    End Select
    Resume Exit_Handler
End Function