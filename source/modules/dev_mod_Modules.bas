Option Compare Database
Option Explicit
Option Private Module   'hides module from external users & makes it easier to see exposed function calls in intellisense

' =================================
' MODULE:       mod_Dev_Modules
' Level:        Development module
' Version:      1.02
' Description:  Dev Module functions & procedures
'
' Source/date:  Bonnie Campbell, 3/9/2020
' Revisions:    BLC, 3/9/2020 - 1.00 - initial version
'               BLC, 3/9/2020 - 1.01 - mod_Dev_xx to dev_mod_xx renaming
'               BLC, 3/10/2020 - 1.02 - added ImportModules()
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------
Public Property Get mdl() As String
    'module name
    mdl = VBE.ActiveCodePane.CodeModule.Parent.Name
End Property

Public Property Let mdl(val As String)
    val = VBE.ActiveCodePane.CodeModule.Parent.Name
    mdl = val
End Property

Public Property Get fxn() As String
    'function/subroutine name
    fxn = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Application.VBE.ActiveCodePane.TopLine, 0)
End Property

Public Property Let fxn(val As String)
    val = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Application.VBE.ActiveCodePane.TopLine, 0)
    fxn = val
End Property

' ---------------------------------
'  Methods
' ---------------------------------
' ---------------------------------
' SUB:          CopyModules
' Description:  Copy modules in from the current db to another
' Assumptions:  MSysObjects Type for modules is -32761
'                           ParentId for modules is -2147483646
' Parameters:   SourceDbPath - full path of source database (string)
'               DestDbPath - full path of destination databse (string) (unneeded)
'               Format paths as C:\dblocation\mydb.accdb
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Hasu, July 8, 2003
'   https://www.tek-tips.com/viewthread.cfm?qid=596803
'   lmb-hb, December 16, 2016
'   https://social.msdn.microsoft.com/Forums/vstudio/en-US/c2efb076-1ae2-44c3-82fb-6e01c992d13a/trying-to-retrieve-all-the-names-of-all-the-current-modules-in-an-access-2013-64bit-accdb-file?forum=accessdev
' Adapted:      Bonnie Campbell, March 9, 2020
' Revisions:
'   BLC - 3/9/2020  - initial version
' ---------------------------------
Public Function CopyModules(SourceDbPath As String, DestDbPath As String)
On Error GoTo Err_Handler
    
    Dim src As DAO.Database
    Dim dest As DAO.Database
    Dim rs As DAO.Recordset
    
DestDbPath = "C:\Projects\TEST_DATA\FORESTVEG\2020 PREP\Forest_Veg_TEST.accdb"
    
    Set src = OpenDatabase(SourceDbPath)
       
    'fetch modules
    Set rs = src.OpenRecordset("SELECT * FROM MSysObjects WHERE ParentId = -2147483646") 'TYPE=-32761")
    
    'iterate through source database modules
    With rs
    
        If Not rs.BOF And rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
        End If
        
        Do While Not .EOF
            Debug.Print .Fields("Name")
            DoCmd.CopyObject DestDbPath, .Fields("Name"), acModule, .Fields("Name")
            'DoCmd.TransferDatabase acImport, "Microsoft Access", SourceDbPath, acModule, .Fields("Name"), .Fields("Name")
            rs.MoveNext
        Loop
        
    End With
        
Exit_Handler:
    rs.Close
    Set rs = Nothing
    Set src = Nothing
    Set dest = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & fxn & "[" & mdl & "])" 'CopyModules[mod_Dev_Modules])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ExportModules
' Description:  Export modules from the current db to a destination folder
' Assumptions:  MSysObjects Type for modules is -32761
'                           ParentId for modules is -2147483646
' Parameters:   SourceDbPath - full path of source database (string)
'               DestPath - full path of destination directory (string)
'               origPrefix - original prefix (string, optional)
'               newPrefix - new prefix to add to module name (string, optional)
'               Format paths as C:\dblocation\mydb.accdb
'               When origPrefix and newPrefix are provided, module names will replace origPrefix with newPrefix in the destination folder
' Usage:        ?ExportModules(CurrentProject.Path & "\dev.accdb","","dev_mod_","dev_")
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Hasu, July 8, 2003
'   https://www.tek-tips.com/viewthread.cfm?qid=596803
'   lmb-hb, December 16, 2016
'   https://social.msdn.microsoft.com/Forums/vstudio/en-US/c2efb076-1ae2-44c3-82fb-6e01c992d13a/trying-to-retrieve-all-the-names-of-all-the-current-modules-in-an-access-2013-64bit-accdb-file?forum=accessdev
'   Duc Van Nguyen, December 9, 2014
'   https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files
' Adapted:      Bonnie Campbell, March 9, 2020
' Revisions:
'   BLC - 3/9/2020  - initial version
' ---------------------------------
Public Function ExportModules(SourceDbPath As String, DestPath As String, _
        Optional origPrefix As String, Optional newPrefix As String)
On Error GoTo Err_Handler

    mdl = VBE.ActiveCodePane.CodeModule.Parent.Name
    Debug.Print mdl
    
    DestPath = "C:\Projects\TEST_DATA\FORESTVEG\2020 PREP\dev\"
    
    'create directory if it doesn't exist
    If Len(dir(DestPath, vbDirectory)) = 0 Then
        MkDir DestPath
    End If
    
    Dim c As VBComponent
    Dim Sfx As String

    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            Dim DestFile As String
            DestFile = DestPath & IIf(Len(origPrefix) > 0, Replace(c.Name, origPrefix, newPrefix), c.Name) & Sfx
            c.Export FileName:=DestFile
        End If
    Next c
        
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & fxn & "[" & mdl & "])" 'ExportModules[mod_Dev_Modules])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ImportModules
' Description:  Import modules into the current db from a destination folder
' Assumptions:  MSysObjects Type for modules is -32761
'                           ParentId for modules is -2147483646
' Parameters:   SourcePath - full path of destination directory (string)
'               DestDbPath   - full path of source database (string)
'               SourceDb     - abbreviation of module source db - dev, vcs, framework (string, optional, default = dev)
'               DeleteOrigModules - whether to remove module of original name (boolean, optional, default = false)
'               origPrefix - original prefix (string, optional)
'               newPrefix - new prefix to add to module name (string, optional)
'               Format paths as C:\dblocation\mydb.accdb
'               When origPrefix and newPrefix are provided, module names will replace origPrefix with newPrefix in the destination folder
' Usage:        ?ExportModules(CurrentProject.Path & "\dev.accdb","","dev_mod_","dev_")
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Ronan Vico, May 2, 2019
'   https://stackoverflow.com/questions/55956116/mass-importing-modules-references-in-vba
'   Ron de Bruin, 2013
'   https://www.rondebruin.nl/win/s9/win002.htm
'   aka_BigRed, January 31, 2008
'   https://www.access-programmers.co.uk/forums/threads/problems-importing-a-module-using-vba.142802/
'   Bob Larson, December 9, 2010
'   https://answers.microsoft.com/en-us/msoffice/forum/all/using-vba-to-check-if-a-module-exists/82483c2c-406b-4b2b-882f-96e4612ef6fb
'   Randall Porter, July 16, 2015
'   https://stackoverflow.com/questions/3792134/get-name-of-current-vba-function
' Adapted:      Bonnie Campbell, March 9, 2020
' Revisions:
'   BLC - 3/10/2020  - initial version
'   BLC - 7/1/2020   - update to match w/ ExportModules functionality
' ---------------------------------
Public Function ImportModules(SourcePath As String, _
        DestDbPath As String, _
        DbProject As String, _
        Optional DeleteOrigModules As Boolean = False)
'        Optional SourceDb As String = "dev", _
'        Optional DeleteOrigModules As Boolean = False)', _
'        Optional origPrefix As String, Optional newPrefix As String)
On Error GoTo Err_Handler

    Dim strFile As String
    Dim c As VBComponent
    
    'add slash if none exists @ end of SourcePath
    strFile = dir(SourcePath & IIf(Right(SourcePath, 1) = "\", "*", "\*"))
    
    'iterate
    Do While Len(strFile) > 0
        Debug.Print strFile
        
Debug.Print Right(strFile, Len(strFile) - InStr(strFile, ","))
Dim ftype As Long

        'determine file type
        Select Case Right(strFile, Len(strFile) - InStr(strFile, ","))
            Case "cls" 'vbext_ct_ClassModule, vbext_ct_Document
                ftype = acModule
            Case "frm" 'vbext_ct_MSForm
                ftype = acForm
            Case "bas" 'vbext_ct_StdModule
                ftype = acModule
            Case Else
                ftype = acModule
        End Select
        
        'If DeleteOrigModules = True Then
            
            For Each c In Application.VBE.VBProjects(DbProject).VBComponents
            
                If c.Name = strFile Then Debug.Print c.Name & "- ALREADY HERE!"
                                
                                                
'                DoCmd.TransferDatabase acImport, "Microsoft Access", SourcePath, acModule, c.Name, DestName
                
            Next
        'End If
        
        strFile = dir
    Loop
    
    'import modules from SourcePath
    
    
    
'    'SourceDbPath = CurrentDb.Properties("Name")   '"C:\Projects\TEST_DATA\FORESTVEG\2020 PREP\dev\"
'
'    Dim c As VBComponent
'    Dim Sfx As String
'
'    'decide which source db to use - dev, vcs, framework or current db (i = 1 so SourceDb = "")
'    Dim i As Integer
'    i = 1   'current application database
'
'    Select Case SourceDb
'        Case "dev"
'            i = 4
'        Case "framework"
'            i = 2
'        Case "vcs"
'            i = 3
'    End Select
'
'    For Each c In Application.VBE.VBProjects(i).VBComponents
'        Select Case c.Type
'            Case vbext_ct_ClassModule, vbext_ct_Document
'                Sfx = ".cls"
'            Case vbext_ct_MSForm
'                Sfx = ".frm"
'            Case vbext_ct_StdModule
'                Sfx = ".bas"
'            Case Else
'                Sfx = ""
'        End Select
'
'        If Sfx <> "" Then
'            Dim DestName As String
'            DestName = IIf(Len(origPrefix) > 0, Replace(c.Name, origPrefix, newPrefix), c.Name)
'
'            'delete module if it exists
'            DoCmd.DeleteObject acModule, DestName
'
'            'delete the module w/ original name
'            If DeleteOrigModules = True Then _
'                DoCmd.DeleteObject acModule, c.Name
'
'            DoCmd.TransferDatabase acImport, "Microsoft Access", SourceDbPath, acModule, c.Name, DestName
'
'        End If
'    Next c
        
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case 7874 'module does not exist
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & fxn & "[" & mdl & "])" 'ImportModules[mod_Dev_Modules])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          CopyModuleToProject
' Description:  Copies module from one project to another
' Assumptions:  -
' Parameters:   FromVBProject - project to copy from (VBIDE.VBProject)
'               ToVBProject   - project to copy to(VBIDE.VBProject)
'               ModuleName    - name of module to copy(String)
'               OverwriteExisting - whether existing modules should be overwritten (Boolean, True = Yes/False=No)
'                                   if true and the module exists in the "to" project, it is overwritten
'                                   if false and the module exists in the "to" project the module is not copied and the function returns "false"
' Usage:        ?
' Returns:      True (successful) or False (if error occurs)
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Chip Pearson, May 6, 2018
'   http://www.cpearson.com/Excel/vbe.aspx
' Adapted:      Bonnie Campbell, July 1, 2020
' Revisions:
'   BLC - 3/10/2020  - initial version
' ---------------------------------
Function CopyModuleToProject(ModuleName As String, _
    FromVBProject As VBIDE.VBProject, _
    ToVBProject As VBIDE.VBProject, _
    OverwriteExisting As Boolean) As Boolean
    
    Dim vbComp As VBIDE.VBComponent
    Dim fName As String
    Dim CompName As String
    Dim s As String
    Dim SlashPos As Long
    Dim ExtPos As Long
    Dim TempVBComp As VBIDE.VBComponent
    Dim CopyModule As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Do some housekeeping validation.
    '''''''''''''''''''''''''''''''''''''''''''''
    If FromVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If
    
    If Trim(ModuleName) = vbNullString Then
        CopyModule = False
        Exit Function
    End If
    
    If ToVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If
    
    If FromVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If
    
    If ToVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If
    
    On Error Resume Next
    Set vbComp = FromVBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
        CopyModule = False
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName - name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    fName = Environ("Temp") & "\" & ModuleName & ".bas"
    If OverwriteExisting = True Then
        ''''''''''''''''''''''''''''''''''''''
        ' If OverwriteExisting = True, Kill
        ' existing temp file & remove
        ' xisting VBComponent from ToVBProject
        ''''''''''''''''''''''''''''''''''''''
        If dir(fName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            Err.Clear
            Kill fName
            If Err.Number <> 0 Then
                CopyModule = False
                Exit Function
            End If
        End If
        With ToVBProject.VBComponents
            .Remove .Item(ModuleName)
        End With
    Else
        '''''''''''''''''''''''''''''''''''''''''
        ' OverwriteExisting = False
        ' IF ModuleName VBComponent exists,
        ' exit with a return code of False
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set vbComp = ToVBProject.VBComponents(ModuleName)
        If Err.Number <> 0 Then
            If Err.Number = 9 Then
                ' module doesn't exist. ignore error.
            Else
                ' other error. get out with return value of False
                CopyModule = False
                Exit Function
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Do the Export and Import operation using FName
    ' and then Kill FName.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FromVBProject.VBComponents(ModuleName).Export FileName:=fName
    
    '''''''''''''''''''''''''''''''''''''
    ' Extract the module name from the
    ' export file name.
    '''''''''''''''''''''''''''''''''''''
    SlashPos = InStrRev(fName, "\")
    ExtPos = InStrRev(fName, ".")
    CompName = mid(fName, SlashPos + 1, ExtPos - SlashPos - 1)
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Document modules (SheetX and ThisWorkbook)
    ' cannot be removed. So, if we are working with
    ' a document object, delete all code in that
    ' component and add the lines of FName
    ' back in to the module.
    ''''''''''''''''''''''''''''''''''''''''''''''
    Set vbComp = Nothing
    Set vbComp = ToVBProject.VBComponents(CompName)
    
    If vbComp Is Nothing Then
        ToVBProject.VBComponents.Import FileName:=fName
    Else
        If vbComp.Type = vbext_ct_Document Then
            ' VBComp is destination module
            Set TempVBComp = ToVBProject.VBComponents.Import(fName)
            ' TempVBComp is source module
            With vbComp.CodeModule
                .DeleteLines 1, .CountOfLines
                s = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, s
            End With
            On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If
    Kill fName
    CopyModule = True

Exit_Handler:
    CopyModuleToProject = CopyModule
    Exit Function
    
Err_Handler:
    Select Case Err.Number
'      Case 7874 'module does not exist
'        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & fxn & "[" & mdl & "])" 'ImportModules[mod_Dev_Modules])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          CopyModulesToProject
' Description:  Copies modules from one project to another
' Assumptions:  CopyModuleToProject function is available
' Parameters:   FromVBProject - project to copy from (VBIDE.VBProject)
'               ToVBProject   - project to copy to(VBIDE.VBProject)
'               ModulePrefix    - prefix for copied modules(String)
'               OverwriteExisting - whether existing modules should be overwritten (Boolean, True = Yes/False=No)
'                                   if true and the module exists in the "to" project, it is overwritten
'                                   if false and the module exists in the "to" project the module is not copied and the function returns "false"
'               ModuleName - name of module to copy (string)
'                            if no name is given ALL modules/classes should be copied
'               AddToTable - name of table to add imported modules/classes to (string, optional, "" means no table is given and modules/class info isn't appended to a table)
' Usage:        ?
' Returns:      True (successful) or False (if error occurs)
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Chip Pearson, May 6, 2018
'   http://www.cpearson.com/Excel/vbe.aspx
' Adapted:      Bonnie Campbell, July 1, 2020
' Revisions:
'   BLC - 7/1/2020  - initial version
' ---------------------------------
Function CopyModulesToProject(ModulePrefix As String, _
    FromVBProject As VBIDE.VBProject, _
    ToVBProject As VBIDE.VBProject, _
    OverwriteExisting As Boolean, _
    ModuleName As String, _
    Optional AddToTable As String = "") As Boolean
    
    CopyModulesToProject = False
    
    If Len(ModuleName) > 0 Then
        'copy only module/class
        CopyModuleToProject ModuleName, FromVBProject, ToVBProject, OverwriteExisting
            
    Else
        'copy all modules/classes
        Dim c As VBComponent

        For Each c In Application.VBE.VBProjects(FromVBProject.Name).VBComponents
            Debug.Print c.Name
            
            CopyModuleToProject c.Name, FromVBProject, ToVBProject, OverwriteExisting
            
            ' append to table (tsys_db_components)
            If Len(AddToTable) > 0 Then
                Dim db As DAO.Database
                Dim rs As DAO.Recordset
                Dim DbFrom As DAO.Database
                Set db = DBEngine.OpenDatabase(ToVBProject.FileName) '(Application.VBE.VBProjects(ToVBProject).FileName)
                Set rs = db.OpenRecordset(AddToTable)
                
                rs.AddNew
                rs!ComponentName = c.Name
                rs!ComponentType = c.Type
                rs!ComponentFrom = FromVBProject.Name
                'rs!ComponentVersion = Nz(DbFrom.Properties("Db Version"), "N/A")
                Set DbFrom = DBEngine.OpenDatabase(FromVBProject.FileName)
                rs!ComponentVersion = Nz(DbFrom.Properties("Db Version"), "N/A")
                'rs.Update
                rs!ComponentType = c.Type
                rs.Update
            End If
        
        Next
        
    End If
    
    
    CopyModulesToProject = True
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & fxn & "[" & mdl & "])" 'ImportModules[mod_Dev_Modules])"
    End Select
    Resume Exit_Handler
End Function


' ---------------------------------
' SUB:          CopyModulesToProject
' Description:  Copies modules from one project to another
' Assumptions:  CopyModuleToProject function is available
' Parameters:   FromVBProject - project to copy from (VBIDE.VBProject)
'               ToVBProject   - project to copy to(VBIDE.VBProject)
'               ModulePrefix    - prefix for copied modules(String)
'               OverwriteExisting - whether existing modules should be overwritten (Boolean, True = Yes/False=No)
'                                   if true and the module exists in the "to" project, it is overwritten
'                                   if false and the module exists in the "to" project the module is not copied and the function returns "false"
'               ModuleName - name of module to copy (string)
'                            if no name is given ALL modules/classes should be copied
'               AddToTable - name of table to add imported modules/classes to (string, optional, "" means no table is given and modules/class info isn't appended to a table)
' Usage:        ?
' Returns:      True (successful) or False (if error occurs)
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Chip Pearson, May 6, 2018
'   http://www.cpearson.com/Excel/vbe.aspx
' Adapted:      Bonnie Campbell, July 1, 2020
' Revisions:
'   BLC - 7/1/2020  - initial version
' ---------------------------------
Public Function AddComponentsToTable( _
    FromVBProject As VBIDE.VBProject, _
    Optional AddToTable As String = "tsys_db_components") As Boolean
        
    AddComponentsToTable = False
        
    'copy all modules/classes
    Dim c As VBComponent
    
    For Each c In Application.VBE.VBProjects(FromVBProject.Name).VBComponents
        Debug.Print c.Name
                   
        ' append to table (tsys_db_components)
        If Len(AddToTable) > 0 Then
            Dim db As DAO.Database
            Dim rs As DAO.Recordset
            Dim DbFrom As DAO.Database
            Set db = DBEngine.OpenDatabase(FromVBProject.FileName)
            Set rs = db.OpenRecordset(AddToTable)
            
            'add new record
            rs.AddNew
            rs!ComponentName = c.Name
            rs!ComponentType = c.Type
            rs!ComponentFrom = FromVBProject.Name
            
            Set DbFrom = DBEngine.OpenDatabase(FromVBProject.FileName)
            rs!ComponentVersion = Nz(DbFrom.Properties("Db Version"), "N/A")
    
            rs.Update
        End If
    
    Next
    
    AddComponentsToTable = True
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & fxn & "[" & mdl & "])"
    End Select
    Resume Exit_Handler
End Function