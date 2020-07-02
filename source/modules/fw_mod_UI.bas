Option Compare Database
Option Explicit

' ---------------------------------
' MODULE:       fw_mod_UI
' Level:        Framework module
' Version:      1.21
' Description:  User interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/10/2015 - 1.01 - added GetRibbonXML()
'               BLC, 5/27/2015 - 1.02 - added functions
'               BLC, 6/30/2015 - 1.03 - moved to mod_Forms: FormIsOpen, FormIsLoaded, SwitchboardIsOpen
'                                       moved from mod_Forms: ChangeBackColor
'               BLC, 5/13/2016 - 1.04 - adapted for Big Rivers
'               BLC, 6/27/2016 - 1.05 - added acNormalSolid, acTransparent constants
'               BLC, 6/24/2016 - 1.06 - replaced Exit_Function > Exit_Handler
'               BLC, 7/6/2016  - 1.07 - added functions to hide VBE (shift off screen)
'                                       while the enum module is being updated
'               BLC, 9/1/2016  - 1.08 - updated ControlExists()
'               BLC, 12/8/2016 - 1.09 - added text alignment constants
'               BLC, 12/12/2016 - 1.10 - added scrolling constants & function
'               BLC, 1/11/2017 - 1.11 - added SetToggleCaption()
'               BLC, 1/26/2017 - 1.12 - added DisplayMsg()
'               BLC, 3/8/2017 - 1.13 - imported into invasives,
'                                      subs/functions not available in invasives
'                                      (missing reference/function):
'                                      OpenAndHideVBE(), ShowAndCloseVBE(), RepaintParentForm()
'                                      CircleControl(), ButtonHighlight(), ButtonUnHighlight()
'               BLC, 6/25/2017 - 1.14 - added SetNavGroup copied from invasives_rpts mod_UI (v 1.04)
'               BLC, 9/15/2017 - 1.15 - added heading for navigation
'               BLC, 10/4/2017 - 1.16 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC, 10/6/2017 - 1.17 - moved ReportIsLoaded() to mod_Reports,
'                                       SetWindowSize(), PopulateSubformControl(),
'                                       RepaintParentForm(), ChangeBackColor(), ResetHeaders(),
'                                       ShowControls(), AddFormControl() to mod_Forms
'               BLC, 11/24/2017 - 1.18 - revised to include general messages
'               BLC, 12/27/2017 - 1.19 - updated ToggleCaption to check for false text
'               BLC, 5/16/2019  - 1.20 - added fw_ module prefix
'               BLC, 3/9/2020   - 1.21 - 64-bit OS updates
' ---------------------------------

' ---------------------------------
' Declarations
' ---------------------------------
Public Const acNormalSolid As Integer = 1
Public Const acTransparent As Integer = 0

'text alignment
Public Const taGeneral As Integer = 0       'default alignment
Public Const taLeft As Integer = 1
Public Const taCenter As Integer = 2
Public Const taRight As Integer = 3
Public Const taDistribute As Integer = 4    'evenly distributed

' ---------------------------------
'  Scrollbars
' ---------------------------------
Declare PtrSafe Function FlatSB_SetScrollPos Lib "comctl32" (ByVal hwnd As Long, ByVal Code As Long, _
                                        ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
Declare PtrSafe Function FlatSB_GetScrollPos Lib "comctl32" (ByVal hwnd As Long, _
                                        ByVal Code As Long) As Long
'Get the Handle of a Control
Public Declare PtrSafe Function apiGetFocus Lib "user32" Alias "GetFocus" () As Long

'scroll bar alignments
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_BOTH = 3

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  VBE
' ---------------------------------

' ---------------------------------
'   NOTES:
'
'
'       Application.Echo False
'       SetWindowPos FindWindow("wndclass_desked_gsk", _
'           Application.VBE.MainWindow.Caption), 0&, 0&, 2000&, 1, 1, &H80 Or &H1
'       <run your code that calls the VBE>
'       Application.VBE.MainWindow.visible = False
'       Application.Echo True
'
' ---------------------------------

Declare PtrSafe Function SetWindowPos Lib "user32.dll" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal ClassName As String, _
    ByVal WindowName As String) As Long

'window positioning & sizing flags
Const HWND_NOTOPMOST = -2
Const SWP_HIDEWINDOW = &H80
Const SWP_NOSIZE = &H1

' ---------------------------------
' Function:     OpenAndHideVBE
' Description:  Opens then hides VBE
' Notes:        Call OpenAndHideVBE before writing to the project
'               and ShowAndCloseVBE when done.
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Peter Thornton, March 23, 2013
'   https://social.msdn.microsoft.com/Forums/en-US/197a9f1d-96cb-49d6-b08c-0dcae1eafc08/vbe-flashes-while-programming-in-the-vbe?forum=isvvba
'   AOB, September 5, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=252942
' Source/date:  Bonnie Campbell, July 6, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/6/2016 - initial version
' ---------------------------------
Public Sub OpenAndHideVBE()
On Error GoTo Err_Handler

    Dim hWndVBE As Long
    Dim objVBE As VBE

    Set objVBE = Application.VBE

    hWndVBE = FindWindow("wndclass_desked_gsk", _
                            Application.VBE.MainWindow.Caption)

    Call SetWindowPos(hWndVBE, 0&, 0&, 2000&, 1, 1, _
                        SWP_HIDEWINDOW Or SWP_NOSIZE)

    Application.VBE.MainWindow.visible = True
    'Application.Caption errors for Access w/ Method or data member not found
    'use "already open form caption", false instead
    'AppActivate Application.Caption
    AppActivate SWITCHBOARD, False
    DoCmd.OpenForm SWITCHBOARD, acNormal, , , , acDialog

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenAndHideVBE[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     ShowAndCloseVBE
' Description:  Displays VBE and closes it
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Peter Thornton, March 23, 2013
'   https://social.msdn.microsoft.com/Forums/en-US/197a9f1d-96cb-49d6-b08c-0dcae1eafc08/vbe-flashes-while-programming-in-the-vbe?forum=isvvba
'   AOB, September 5, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=252942
' Source/date:  Bonnie Campbell, July 6, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/6/2016 - initial version
' ---------------------------------
Public Sub ShowAndCloseVBE()
On Error GoTo Err_Handler

    Dim hWndVBE As Long
    Dim cbt As CommandBarButton
    Dim objVBE As VBIDE.VBE
    Dim objWin As VBIDE.Window

    Set objVBE = Application.VBE
    ' optionally close all module windows,
    ' or just the newly opened module Window

    For Each objWin In objVBE.Windows
        If objWin.Type = vbext_wt_CodeWindow Then
                objWin.Close
        ElseIf objWin.Type = vbext_wt_Designer Then
                objWin.Close
        End If
    Next

    objVBE.MainWindow.WindowState = vbext_ws_Minimize
    objVBE.MainWindow.visible = False

    hWndVBE = FindWindow("wndclass_desked_gsk", _
                            Application.VBE.MainWindow.Caption)

    Call SetWindowPos(hWndVBE, HWND_NOTOPMOST, 0, 0, 400, 300, 0)

    Set cbt = Application.VBE.CommandBars.FindControl(ID:=752)

    'Application.Caption errors for Access w/ Method or data member not found
    'use "already open form caption", false instead
    'AppActivate Application.Caption
    AppActivate SWITCHBOARD, False
    DoCmd.OpenForm SWITCHBOARD, acNormal, , , , acDialog

    cbt.Execute

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ShowAndCloseVBE[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Ribbon
' ---------------------------------
' ---------------------------------
' FUNCTION:     GetRibbonXML
' Description:  gets ribbon UI XML specified, if found
' Assumes:      USysRibbon table exists
' Parameters:   ribbon - name of the ribbon to retrieve, RibbonName in USysRibbon (string)
' Returns:      XML of the specified ribbon
' Throws:       none
' References:   none
' Source/date:  -
' Revisions:    BLC, 5/10/2015 - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetRibbonXML(strRibbon As String) As String
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Dim strSQL As String, strXML As String
    
    strSQL = "SELECT RibbonXML FROM USysRibbons WHERE RibbonName = '" & strRibbon & "';"
    strXML = ""
    
    Set rs = CurrDb.OpenRecordset(strSQL)
    If Not (rs.BOF And rs.EOF) Then
        strXML = rs!RibbonXML
    End If
    
    GetRibbonXML = strXML

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRibbonXML[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          RibbonOnLoad
' Description:  Callback function for ribbon customization
' Parameters:   ribbon - office ribbon control (IRibbonUI object)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from http://www.experts-exchange.com/Database/MS_Access/Q_28470268.html
'               by Christian, 7/7/2014.
' Revisions:    BLC, 5/17/2015 - initial version
' ---------------------------------
'Public objRibbon As IRibbonUI
Public Sub RibbonOnLoad(ribbon As Office.IRibbonUI)
On Error GoTo Err_Handler
Dim prv_Ribbon As IRibbonUI

    Set prv_Ribbon = ribbon

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RibbonOnLoad[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          GetRibbonVisibility
' Description:  Callback function to indicate if ribbon control should be displayed or not
' Parameters:   ctrl - office ribbon control (IRibbonControl object)
'               visible - true (boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from http://www.access-programmers.co.uk/forums/showthread.php?t=246015
'               by Mark K., 4/26/2013.
' Revisions:    BLC, 5/10/2015 - initial version
' ---------------------------------
Public Sub GetRibbonVisibility(ctrl As Office.IRibbonControl, ByRef visible)
On Error GoTo Err_Handler

    Select Case ctrl.ID
        Case "tabExportOptions"
            visible = True
            TempVars.Add "ribbon", True
        Case Else
            visible = False
            TempVars.Add "ribbon", False
    End Select
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRibbonVisibility[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Navigation
' ---------------------------------

' ---------------------------------
' SUB:          SetNavGroup
' Description:  Set the navigation group for the item
' Parameters:   strGroup - name of group to move object to (string)
'               stTable - name of table (object) to move (string)
'               strType - type of object to move (string)
' Returns:      -
' Throws:       none
' References:
'   Wayne G. Dunn, December 9, 2014
'   Phillippe R, February 9, 2016
'   https://stackoverflow.com/questions/27366038/change-navigation-pane-group-in-access-through-vba
' Source/date:  Bonnie Campbell June 25, 2017 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 6/25/2017 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' ---------------------------------
Public Function SetNavGroup(strGroup As String, strTable As String, strType As String) As String
On Error GoTo Err_Handler

    Dim strSQL          As String
    Dim dbs             As DAO.Database
    Dim rs              As DAO.Recordset
    Dim lCatID          As Long
    Dim lGrpID          As Long
    Dim lObjID          As Long
    Dim lType           As Long

    'default
    SetNavGroup = "Failed"
    
    Set dbs = CurrDb

    '-- Category Code --
    ' Ignore the following code unless you want to manage 'Categories'
    '    Table MSysNavPaneGroupCategories has fields: Filter, Flags, Id (AutoNumber), Name, Position, SelectedObjectID, Type
    '    strSQL = "SELECT Id, Name, Position, Type " & _
    '            "FROM MSysNavPaneGroupCategories " & _
    '            "WHERE (((MSysNavPaneGroupCategories.Name)='" & strGroup & "'));"
    '    Set rs = dbs.OpenRecordset(strSQL)
    '    If rs.EOF Then
    '        MsgBox "No group named '" & strGroup & "' found. Will quit now.", vbOKOnly, "No Group Found"
    '        rs.Close
    '        Set rs = Nothing
    '        dbs.Close
    '        Set dbs = Nothing
    '        Exit Function
    '    End If
    '    lCatID = rs!ID
    '    rs.Close

    ' New table's names are added to table 'MSysNavPaneObjectIDs'

    ' Types
        ' Type TypeDesc
        '-32768  Form                       1   Table - Local Access Tables
        '-32766  Macro                      2   Access object - Database
        '-32764  Reports                    3   Access object - Containers
        '-32761  Module                     4   Table - Linked ODBC Tables
        '-32758  Users                      5   Queries
        '-32757  Database Document          6   Table - Linked Access Tables
        '-32756  Data Access Pages          8   SubDataSheets
        
    If LCase(strType) = "table" Then
        lType = 1
    ElseIf LCase(strType) = "query" Then
        lType = 5
    ElseIf LCase(strType) = "form" Then
        lType = -32768
    ElseIf LCase(strType) = "report" Then
        lType = -32764
    ElseIf LCase(strType) = "module" Then
        lType = -32761
    ElseIf LCase(strType) = "macro" Then
        lType = -32766
    Else
        MsgBox "Add your own code to handle the object type of '" & strType & "'", vbOKOnly, _
                "Add Code"
        dbs.Close
        Set dbs = Nothing
        Exit Function
    End If

    ' Table MSysNavPaneGroups has fields: Flags, GroupCategoryID, Id, Name,
    '                                     Object, Type, Group, ObjectID, Position
'    Debug.Print "---------------------------------------"
'    Debug.Print "Add '" & strType & "' " & strTable & "' to Group '" & strGroup & "'"
    strSQL = "SELECT GroupCategoryID, Id, Name " & _
            "FROM MSysNavPaneGroups " & _
            "WHERE (((MSysNavPaneGroups.Name)='" & strGroup & "') " & _
            "AND ((MSysNavPaneGroups.Name) Not Like 'Unassigned*'));"
    Set rs = dbs.OpenRecordset(strSQL)
    If rs.EOF Then
        MsgBox "No group named '" & strGroup & "' found. Will quit now.", vbOKOnly, _
                "No Group Found"
        rs.Close
        Set rs = Nothing
        dbs.Close
        Set dbs = Nothing
        Exit Function
    End If
 '   Debug.Print rs!GroupCategoryID & vbTab & rs!ID & vbTab & rs!Name
    lGrpID = rs!ID
    rs.Close

    ' Get Table ID From MSysObjects
    strSQL = "SELECT * " & _
        "FROM MSysObjects " & _
        "WHERE (((MSysObjects.Name)='" & strTable & "') AND ((MSysObjects.Type)=" & lType & "));"
    Set rs = dbs.OpenRecordset(strSQL)
    If rs.EOF Then
        MsgBox "This is crazy! Table '" & strTable & "' not found in MSysObjects.", vbOKOnly, "No Table Found"
        rs.Close
        Set rs = Nothing
        dbs.Close
        Set dbs = Nothing
        Exit Function
    End If
    
    lObjID = rs!ID

    Debug.Print "Table found in MSysObjects " & lObjID & " . Lets compare to MSysNavPaneObjectIDs."

    ' Filter By Type
    strSQL = "SELECT Id, Name, Type " & _
            "FROM MSysNavPaneObjectIDs " & _
            "WHERE (((MSysNavPaneObjectIDs.ID)=" & lObjID & ") AND ((MSysNavPaneObjectIDs.Type)=" & lType & "));"
    Set rs = dbs.OpenRecordset(strSQL)
    If rs.EOF Then
        ' Seems to be a refresh issue / delay!  I have found no way to force a refresh.
        ' This table gets rebuilt at the whim of Access, so let's try a different approach....
        ' Lets add the record via this code.
        Debug.Print "Table not found in MSysNavPaneObjectIDs, add it from MSysObjects."
        strSQL = "INSERT INTO MSysNavPaneObjectIDs ( ID, Name, Type ) VALUES ( " & lObjID & ", '" & strTable & "', " & lType & ")"
        dbs.Execute strSQL
    End If
    Debug.Print lObjID & vbTab & strTable & vbTab & lType
    rs.Close

'Try_Again:
'    ' Filter By Type
'    strSQL = "SELECT Id, Name, Type " & _
'            "FROM MSysNavPaneObjectIDs " & _
'            "WHERE (((MSysNavPaneObjectIDs.Name)='" & strTable & "') " _
'            & "AND ((MSysNavPaneObjectIDs.Type)=" & lType & "));"
'
'    Set rs = dbs.OpenRecordset(strSQL)
'    If rs.EOF Then
'        ' Seems to be a refresh issue / delay!  I have found no way to force a refresh.
'        ' This table gets rebuilt at the whim of Access, so let's try a different approach....
'        ' Lets add the record vis code.
'        Debug.Print "Table not found in MSysNavPaneObjectIDs, try MSysObjects."
'         strSQL = "SELECT * " & _
'            "FROM MSysObjects " & _
'            "WHERE (((MSysObjects.Name)='" & strTable & "') AND " & _
'            "((MSysObjects.Type)=" & lType & "));"
'        Set rs = dbs.OpenRecordset(strSQL)
'        If rs.EOF Then
'            MsgBox "This is crazy! Table '" & strTable & "' not found in MSysObjects.", vbOKOnly, "No Table Found"
'            rs.Close
'            Set rs = Nothing
'            dbs.Close
'            Set dbs = Nothing
'            Exit Function
'        Else
'            Debug.Print "Table not found in MSysNavPaneObjectIDs, but was found in MSysObjects. Lets try to add via code."
'            strSQL = "INSERT INTO MSysNavPaneObjectIDs ( ID, Name, Type ) VALUES ( " & rs!ID & ", '" & strTable & "', " & lType & ")"
'            dbs.Execute strSQL
'            GoTo Try_Again
'        End If
'    End If
''    Debug.Print rs!ID & vbTab & rs!Name & vbTab & rs!Type
'    lObjID = rs!ID
'    rs.Close

    ' Add the table to the Custom group
    strSQL = "INSERT INTO MSysNavPaneGroupToObjects ( GroupID, ObjectID, Name ) VALUES ( " & lGrpID & ", " & lObjID & ", '" & strTable & "' )"
    dbs.Execute strSQL

    dbs.Close
    Set dbs = Nothing
    SetNavGroup = "Passed"

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetNavGroup[fw_mod_UI])"
    End Select
    Resume Exit_Handler

End Function

' ---------------------------------
'  Tabs
' ---------------------------------

' ---------------------------------
' SUB:          tabPageUnhide
' Description:  sets desired tab visible, all others hidden
' Parameters:   strTabName - tab page name to make visible
'               ctrl - tab control
'               blnHideOnly - true to hide tabs only (Boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from Tom's post comment, 9/12/2009
'               http://www.vbdotnetforums.com/gui/36561-loop-through-tab-pages-remove.html
'               Created 06/11/2014 blc; Last modified 06/11/2014 blc.
' Adapted:      Bonnie Campbell, June 11, 2014 - initial version
' Revisions:    BLC, June 11, 2014 - initial version
'               BLC, June 9, 2015  - adjust for hiding tabs only with blnHideOnly
' ---------------------------------
Public Sub tabPageUnhide(ctrl As TabControl, strTabName As String, Optional blnHideOnly As Boolean)
On Error GoTo Err_Handler

    Dim pg As Page
    
    For Each pg In ctrl.Pages
        If pg.Name = strTabName Then
            If Not blnHideOnly = True Then
                ctrl.Pages(pg.Name).visible = True
            End If
        Else
            ctrl.Pages(pg.Name).visible = False
        End If
    Next pg
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tabPageUnhide[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Controls
' ---------------------------------

' ---------------------------------
' FUNCTION:     HideObject
' Description:  Changes the hidden property of an object to hide / show in the database window
' Parameters:   strObjectName - name of the object (string)
'               blnHide - True to hide, False to show (default True)
'               varType - object type (default acTable)
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/25/2009
' Revisions:    JRB, 6/25/2009 - initial version
'               BLC, 4/30/2015 - move from mod_Utilities to mod_UI
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 10/6/2017 - changed from Function to Sub
' ---------------------------------
Public Sub HideObject(strObjectName As String, _
                        Optional blnHide As Boolean = True, _
                        Optional varType As Variant = acTable)
    On Error GoTo Err_Handler

    SetHiddenAttribute varType, strObjectName, blnHide

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HideObject[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          buttonHighlight
' Description:  Toggle button color to strColor or transparent if already colored
' Parameters:   btn      - name of the button to change
'                          accommodates command and label as control buttons
'               strColor - HTML color without # (string, optional)
'               solo - display only this control & leave others transparent (Boolean)
'               toggle - change the display for a control (Boolean)
'               intEffect - control display effect (integer)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub buttonHighlight(btn As Control, Optional solo As Boolean, _
                        Optional Toggle As Boolean, Optional intEffect As Integer, _
                        Optional strColor As String)
' Special Effects:  0 - flat, 1 - raised, 2 - sunken, 3 - etched, 4 - shadowed, 5 - chiseled
' Colors:
'   lime                   #9EFF00
'   chartreuse 1           #7FFF00 127 255 00  65407
'   dark olive green 1     #CAFF70 202 255 112 7405514
'   mint                   #BDFCC9 189 252 201 13237437
'   light lime (like)      #E6FABF 230 250 191
'   darker lt lime         #CFF583 207 245 131
On Error GoTo Err_Handler:

    'toggle button
    If Toggle Then
        buttonUnHighlight btn, Toggle
    End If
    
    'change all others to transparent if solo
    If solo Then
        buttonUnHighlight btn
    End If
    
    With btn
        If .backstyle = 1 Then
            GoTo Transparent
        End If
        
        If (Len(strColor) <> 6) Then
            strColor = "CFF583"
        End If
    
        If intEffect > -1 Or intEffect > 6 Then
            intEffect = 0 'flat
        End If
           
        'change button background to given color
        .backstyle = 1 'Normal - required to change color
        .backcolor = HTMLConvert("#" & strColor)
        .SpecialEffect = intEffect
    End With
    
Exit_Procedure:
    Exit Sub

Transparent:
    btn.backstyle = 0 'Transparent
    GoTo Exit_Procedure

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - buttonHighlight[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          buttonUnHighlight
' Description:  Toggles all other buttons to transparent if already colored
' Parameters:   btn - name of the button control to change
'                     accommodates command and label as control buttons
'               blnToggle - toggle only the identified button (Boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub buttonUnHighlight(btn As Control, Optional blnToggle As Boolean)
On Error GoTo Err_Handler:
Dim ctl As Control

    With btn
        'unhighlight only btn
        If blnToggle Then
            .backstyle = 0 'transparent
            .SpecialEffect = 0 'flat
            GoTo Exit_Procedure
        End If

        'unhighlight all other buttons
        For Each ctl In .Parent.Controls

            If ctl.Name <> btn.Name And _
                ctl.ControlType = acLabel Then
                With ctl
                    .backstyle = 0 'transparent
                End With
            End If

        Next

    End With

Exit_Procedure:
    'update display
    RepaintParentForm btn
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - buttonUnHighlight[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          DisableControl
' Description:  Set color scheme for labels so they appear disabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - moved from mod_List to mod_UI
' ---------------------------------
Public Sub DisableControl(ctrl As Control)

On Error GoTo Err_Handler
    
    ctrl.backcolor = lngLtGray
    ctrl.forecolor = lngGray
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = lngGray
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[fw_mod_UI])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          EnableControl
' Description:  Set color scheme for labels so they appear enabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
'               backColor - long value for desired back color
'               foreColor - long value for desired fore (text) color
'               optionally for command buttons:
'               borderColor - long value for desired border color
'               hoverColor - long value for desired hover color
'               pressedColor - long value for desired pressed button color
'               hoverForeColor - long value for desired hover fore (text) color
'               pressedForeColor - long value for desired pressed button fore (text) color
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - moved from mod_List to mod_UI
' ---------------------------------
Public Function EnableControl(ctrl As Control, backcolor As Long, forecolor As Long, _
                                Optional borderColor As Long, _
                                Optional hoverColor As Long, _
                                Optional pressColor As Long, _
                                Optional hoverForeColor As Long, _
                                Optional pressedForeColor As Long)
On Error GoTo Err_Handler
    
    ctrl.backcolor = backcolor
    ctrl.forecolor = forecolor
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = borderColor
        ctrl.hoverColor = hoverColor
        ctrl.pressedColor = pressColor
        ctrl.hoverForeColor = hoverForeColor
        ctrl.pressedForeColor = pressedForeColor
    End If

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableControl[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ToggleControl
' Description:  Toggles control font (fore) color & enables/disables
' Parameters:   frmName - name of parent form (string)
'               btnName - name of the button control to change
'                     accommodates command and label as control buttons (string)
'               color - optional color value (long)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub ToggleControl(frmName As String, btnName As String, Optional color As Variant = Null)
On Error GoTo Err_Handler:
    
    Dim ctrl As Control
    Set ctrl = Forms(frmName).Controls(btnName)
    
    'invert enabled value (change true -> false, false -> true) & change color
    With ctrl
    
        'enable/disable control (includes acCommandButton, acComboBox, acListBox, acTextBox, acToggleButton)
        If Not ctrl.ControlType = acLabel Then
            .Enabled = Not .Enabled
        End If
        
        If Not IsNull(color) Then
            ' change font color for appropriate controls with text
            Select Case ctrl.ControlType
                Case acCommandButton, acComboBox, acLabel, acListBox, acTextBox, acToggleButton
                    .forecolor = color
                Case Else
            End Select
        End If
    End With
    
Exit_Procedure:
    'update display
    RepaintParentForm Forms(frmName).Controls(btnName)
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleControl[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          ToggleCaption
' Description:  Sets toggle button caption based on button state
' Assumptions:  -
' Parameters:   ctrl - tgl (toggle button control)
'               TrueText - caption text to display when toggle is true (optional, string)
'               FalseText - caption text to display when toggle is false (optional, string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 11, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2017 - initial version
'   BLC - 12/27/2017 - revised to use control caption vs control on IF
'                     (avoids error #94 invalid use of Null)
' ---------------------------------
Public Sub ToggleCaption(ctrl As ToggleButton, blnCheckbox As Boolean, _
                Optional TrueText As String = "", _
                Optional FalseText As String = "")
On Error GoTo Err_Handler
    
    'set default if checkbox desired
    If blnCheckbox Then TrueText = StringFromCodepoint(uCheck)
    
    If Nz(ctrl.Caption, "") = FalseText Then
        ctrl.Caption = TrueText
    Else
        ctrl.Caption = FalseText
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleCaption[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Scrollbars
' ---------------------------------

' ---------------------------------
' FUNCTION:     fhWnd
' Description:  Returns the handle of a control
' Assumptions:
'   Used in combination w/ FlatSB_Set/GetScrollPos, apiGetFocus,
'   SB_ constants for synchronizing scrollbar positions in listboxes
' Parameters:   ctl - control whose handle is desired (control)
' Returns:      control handle (long)
' Throws:       none
' References:
'   CyberLynx, December 10, 2003
'   http://www.dbforums.com/showthread.php?973824-Sync-Scrolling-of-Two-Listboxes
' Source/date:  Bonnie Campbell, December 12, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/12/2016 - initial version
' ---------------------------------
Public Function fhWnd(ctl As Control) As Long
    On Error GoTo Err_Handler 'Resume Next
    
    ctl.SetFocus
    
    If Err Then
        fhWnd = 0
    Else
        fhWnd = apiGetFocus
    End If
    
'    On Error GoTo 0

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddControl[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Text
' ---------------------------------

' ---------------------------------
' FUNCTION:     CrumbsToArray
' Description:  Prepares breadcrumb elements from Me.OpenArgs values
' Parameters:   strCrumbs - Me.OpenArgs values from form open subs
'               delimiter - delimiter used for separating string values, default = | (pipe)
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    BLC, 6/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 12/12/2016 - revised to use Exit_handler vs. Exit_Function/Procedure
' ---------------------------------
Public Function CrumbsToArray(strCrumbs As String, Optional delimiter = "|") As Variant

On Error GoTo Err_Handler

    Dim strCrumbTrail As String

    If Len(strCrumbs) > 0 Then
        Dim aryCrumbs As Variant
        
        aryCrumbs = Split(strCrumbs, delimiter)
        
    End If

    CrumbsToArray = aryCrumbs
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CrumbsToArray[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:     PrepareCrumbs
' Description:  Sets breadcrumb label control captions & click events based on crumb element array
' Assumptions:  Breadcrumbs are displayed using label controls (lblCrumb01...)
'               & labels already exist on the targeted form
' Parameters:   frm - form holding crumb labels
'               aryCrumbs - breadcrumb array
'               separator - non-clickable value between crumbs, default = >
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    BLC, 6/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub PrepareCrumbs(frm As SubForm, aryCrumbs As Variant, Optional separator = ">")
 On Error GoTo Err_Handler
 
    Dim ctrl As Control
    Dim i As Integer, intLastCtrlWidth As Integer, intLastCtrlPosition As Integer
    Dim strNum As String, strCtrlName As String, strCtrlSeparator As String
    
    'initialize
    intLastCtrlPosition = 10
    
    'avoid flicker
    'Painting = False
    
    For i = 1 To UBound(aryCrumbs)
        ' set lbl caption
        If (i < 10) Then
            strNum = 0 & i
        Else
            strNum = i
        End If
        
        strCtrlName = "lblCrumb" & strNum
        
        With frm.Controls(strCtrlName)
       
            If .ControlType = acLabel Then
                'label control
                .Caption = aryCrumbs(i)
            Else
                'hyperlink control (displaytext vs caption)
                .Value = aryCrumbs(i)
            End If
            
            'set control position
            If intLastCtrlPosition > frm.Controls(strCtrlName).Parent.Width Then
                .Left = frm.Controls(strCtrlName).Parent.Width - .Width
            Else
                .Left = intLastCtrlPosition
            End If
            
            'set control width
'            setControlWidth frm.Controls(strCtrlName), , frm.Controls(strCtrlName).Parent.Width
            
            'save new ctrl width for setting separator position
            intLastCtrlWidth = .Width
        
        End With
        
        'display the separator
        If (i < UBound(aryCrumbs)) Then
          strCtrlSeparator = "lblSep" & strNum
          With frm.Controls(strCtrlSeparator)
            .Left = intLastCtrlPosition + intLastCtrlWidth + 10
            .Caption = separator
            .visible = True
            
            'determine position of next control
            intLastCtrlPosition = .Left + .Width + 10
          End With
        End If
        
    Next i
    
    'ready for viewing
    'Painting = True
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareCrumbs[fw_mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' FUNCTION:     ColorizeText
' Description:  Colors specific text items based on the colorizing type.
' Assumptions:  Text is used in Rich Text textboxes or HTML.
' Parameters:   InputText - text to colorize (string)
'               TextType - type of colorizing to do (string)
'               TextColor - text coloring to use (string)
' Returns:      -
' Throws:       none
' References:
'   TJ Poorman, August 15, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=251953
' Source/date:  Bonnie Campbell, January 10, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2017 - initial version
' ---------------------------------
Public Function ColorizeText(InputText As String, TextType As String, _
                                Optional TextColor As String = "red") As String
On Error GoTo Err_Handler
    
    Dim aryText() As Variant
    Dim colorizedText As String
    Dim colorizeStart As String
    Dim colorizeEnd As String
    Dim i As Integer
    
    Select Case TextType
        Case "SQL"
            aryText = Array("SELECT", "INSERT", "UPDATE", "DELETE", "FROM", "WHERE", _
                            "LEFT OUTER JOIN", "LEFT INNER JOIN", "RIGHT OUTER JOIN", _
                            "RIGHT INNER JOIN", _
                            "INTO", " BETWEEN ", " IN ", "INNER JOIN", "OUTER JOIN", _
                            "LEFT JOIN", "RIGHT JOIN", "JOIN", " AS ", "UNION", _
                            "UNION ALL", "PARAMETERS", "ORDER BY", "TOP", _
                            "VALUES", "IIf", "LCASE", "UCASE", "DISTINCT", _
                            " ON ", "COUNT", "DESC", "ASC", "AND")
        Case "NULL"
            aryText = Array("NULL")
        Case "NEGATIVE"
            aryText = Array("NULL", "NOT", "IS NOTHING")
        Case "COLTYPES"
            aryText = Array("INT", "TEXT", "BYTE", "DOUBLE", "LONG", "MEMO")
    End Select
    
    'setup colorizing
    colorizeStart = " <font color=" & TextColor & ">"
    colorizeEnd = "</font> "
    
    'begin
    colorizedText = InputText
    
    'iterate through text
    For i = 0 To UBound(aryText)
    
        If InStr(1, InputText, aryText(i)) Then
               
            colorizedText = Replace(colorizedText, aryText(i), colorizeStart & aryText(i) & colorizeEnd)
        
        End If
        
    Next

    ColorizeText = colorizedText

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ColorizeText[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Function


' ---------------------------------
'  Drawing
' ---------------------------------

' ---------------------------------
' SUB:          CircleControl
' Description:  Draws a circle around the control
' Assumptions:  -
' Parameters:   ctrl - control to circle (control)
'               ellipse - whether it should be an ellipse vs. circle (boolean)
' Returns:      -
' Throws:       none
' References:
'   Duane Hookom, October 6, 2008
'   http://www.pcreview.co.uk/threads/circle-a-word-in-access-report.3639434/
'
'   https://msdn.microsoft.com/en-us/library/office/aa195881(v=office.11).aspx
' Source/date:  Bonnie Campbell, May 10, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/10/2016 - initial version
' ---------------------------------
Public Sub CircleControl(ctrl As Control, Optional ellipse As Boolean = False)
On Error GoTo Err_Handler

    Dim iWidth As Integer, iHeight As Integer
    Dim iCenterX As Integer, iCenterY As Integer
    Dim iRadius As Integer
    Dim dblAspect As Double
    Dim sngStart As Single, sngEnd As Single

    iCenterX = ctrl.Left + ctrl.Width / 2
    iCenterY = ctrl.Top + ctrl.Height / 2
    iRadius = ctrl.Width '/ 3 '/ 2 + 100
    dblAspect = 1 'ctrl.Height / ctrl.Width

    sngStart = -0.00000001                    ' Start of pie slice.

    sngEnd = -2 * PI / 3                         ' End of pie slice.
    ctrl.Parent.FillColor = RGB(51, 51, 51)            ' Color pie slice red.
    ctrl.Parent.FillStyle = 0                          ' Fill pie slice.

    'add the circle to the parent
    ' X,Y center | radius | [ color, start, end, aspect ]
    ctrl.Parent.Circle (iCenterX, iCenterY), iRadius, lngLime, sngStart, sngEnd, dblAspect

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CircleControl[fw_mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Messages
' ---------------------------------

' ---------------------------------
' SUB:          DisplayMsg
' Description:  display a message specific for the database
' Assumptions:  general messages will include empty msg value,
'               plus text, type & title as desired
' Parameters:   msg - type of message to display (string)
'               msgText - text for message (optional, string, default "")
'               msgType - type for message sets background (optional, string, default "")
'               msgTitle - title for message (optional, string, default "")
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, January 26, 2017 - for NCPN tools
' Revisions:
'   BLC - 1/26/2017  - initial version
'   BLC - 11/24/2017 - revised to include general messages
' ---------------------------------
Public Sub DisplayMsg(msg As String, _
                        Optional msgText As String = "", _
                        Optional msgType As String = "", _
                        Optional msgTitle As String = "")
On Error GoTo Err_Handler
    
    Select Case msg
        Case "mx"   'fixing
            msgText = "Functionality currently under maintenance."
            msgType = "caution"
            msgTitle = "Feature Unavailable"
        Case "dev"  'in progress
            msgText = "Functionality under development."
            msgType = "caution"
            msgTitle = "Feature Unavailable"
        Case "undev" 'undeveloped functionality
            msgText = "Functionality not yet defined && developed."
            msgType = "caution"
            msgTitle = "Feature Unavailable"
        Case "" 'general messages - text, type, title passed in
    End Select

    'show msg
    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
        "msg" & PARAM_SEPARATOR & msgText & _
        "|Type" & PARAM_SEPARATOR & msgType & "|Title" & PARAM_SEPARATOR & msgTitle

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisplayMsg[fw_mod_UI])"
    End Select
    Resume Exit_Sub
End Sub