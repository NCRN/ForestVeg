Option Compare Database
Option Explicit

Function GetSummaryInfo(strPropName As String, Optional varFileName As Variant) As String
  ' Comments: Get "Summary" properties of Database.  Taken
  '           from Access97 help file.
  ' Parameters: strPropName = Name of property
  ' Return: "None" is returned if the property hasn't already been set.
  '         If an unknown error occurs, a zero-length string (" ") is returned.
  ' Dependencies: None
  ' Created: 9/15/00 MAW
  ' Modified:6/19/01 SDK to accept varFileName as additional arg. to work with external db's.
  '
  ' --------------------------------------------------------
  Dim DBS As Database, cnt As Container
  Dim doc As Document, prp As Property

  ' Property not found error.
  Const conPropertyNotFound = 3270
  On Error GoTo GetSummary_Err
  
  If IsMissing(varFileName) Then
    Set DBS = CurrentDb
  Else
    On Error Resume Next
    Set DBS = DBEngine.OpenDatabase(varFileName)
    If Err <> 0 Then
        GetSummaryInfo = ""
        GoTo GetSummary_Bye
    End If
  End If
  
  Set cnt = DBS.Containers!Databases
  Set doc = cnt.Documents!SummaryInfo
  doc.Properties.Refresh
  GetSummaryInfo = doc.Properties(strPropName)

GetSummary_Bye:
  DBS.Close
  Set DBS = Nothing
  Set cnt = Nothing
  Set doc = Nothing
  Exit Function

GetSummary_Err:
  If Err = conPropertyNotFound Then
    Set prp = doc.CreateProperty(strPropName, dbText, "None")
    ' Append to collection.
    doc.Properties.Append prp
    Resume
  Else

' Unknown error.
    GetSummaryInfo = ""
    Resume GetSummary_Bye
  End If
End Function

Function GetCustomInfo(strPropName As String, Optional varFileName As Variant) As String
  ' Comments: Get "Summary" properties of Database.  Taken
  '           from Access97 help file.
  ' Parameters: strPropName = Name of property
  ' Return: "None" is returned if the property hasn't already been set.
  '         If an unknown error occurs, a zero-length string (" ") is returned.
  ' Dependencies: None
  ' Created: 9/15/00 MAW
  ' Modified:6/19/01 SDK to accept varFileName as additional arg. to work with external db's.
  '
  ' --------------------------------------------------------
  Dim DBS As Database, cnt As Container
  Dim doc As Document, prp As Property

  ' Property not found error.
  Const conPropertyNotFound = 3270
  On Error GoTo GetSummary_Err
  
  If IsMissing(varFileName) Then
    Set DBS = CurrentDb
  ElseIf Len(varFileName) = 0 Then
    GetCustomInfo = ""
    GoTo GetSummary_Bye
  Else
    On Error Resume Next
    Set DBS = DBEngine.OpenDatabase(varFileName)
    If Err <> 0 Then
        GetCustomInfo = ""
        GoTo GetSummary_Bye
    End If
  End If
  
  Set cnt = DBS.Containers!Databases
  Set doc = cnt.Documents!UserDefined
  doc.Properties.Refresh
  GetCustomInfo = doc.Properties(strPropName)

GetSummary_Bye:
  DBS.Close
  Set DBS = Nothing
  Set cnt = Nothing
  Set doc = Nothing
  Exit Function

GetSummary_Err:
' Unknown error.
    GetCustomInfo = ""
    Resume GetSummary_Bye
End Function

Function SetCustomProperty(strPropName As String, intPropType _
    As Integer, varPropValue As Variant, Optional varDBName As Variant) As Integer

    Dim DBS As Database, cnt As Container
    Dim wrkJet As Workspace

    Dim doc As Document, prp As Property

    Const conPropertyNotFound = 3270    ' Property not found error.
    ' Define Database object.
    If IsMissing(varDBName) Then
        Set DBS = CurrentDb
    Else
        Set wrkJet = CreateWorkspace("NewJetWorkspace", _
        "admin", "", dbUseJet)
        Set DBS = wrkJet.OpenDatabase(varDBName)
    End If
    Set cnt = DBS.Containers!Databases  ' Define Container object.
    Set doc = cnt.Documents!UserDefined ' Define Document object.
    On Error GoTo SetCustom_Err
    doc.Properties.Refresh
    ' Set custom property name. If error occurs here it means
    ' property doesn't exist and needs to be created and appended

' to Properties collection of Document object.
    Set prp = doc.Properties(strPropName)
    prp = varPropValue                  ' Set custom property value.
    SetCustomProperty = True

SetCustom_Bye:
    On Error Resume Next
    Set prp = Nothing
    Set doc = Nothing
    Set cnt = Nothing
    DBS.Close
    Set DBS = Nothing
    Set wrkJet = Nothing
    Exit Function

SetCustom_Err:
    If Err = conPropertyNotFound Then
        Set prp = doc.CreateProperty(strPropName, intPropType, varPropValue)
        doc.Properties.Append prp       ' Append to collection.
        Resume Next
    Else                                        ' Unknown error.
        SetCustomProperty = False
        Resume SetCustom_Bye
    End If
End Function

Public Function AddAppProperty(strName As String, varType As Variant, varValue As Variant) As Integer
' Description:  Subroutine for adding/editing some GUI aspects of an Access application
' Source/date:  Unknown
' Revisions:    Alan Williams, Oct 5, 2005 - documentation
'               Simon Kingston, Oct 24, 2006 - added cleanup of objects
' Example Calls:    intX = AddAppProperty("AppTitle", dbText, "My Custom Application")
'                   intX = AddAppProperty("AppIcon", dbText, strTempDir & "my.ICO")
    
    Dim DBS As Database, prp As Property
    Const conPropNotFoundError = 3270

    Set DBS = CurrentDb
    On Error GoTo AddProp_Err
    DBS.Properties(strName) = varValue

AddAppProperty = True

AddProp_Bye:
    On Error Resume Next
    Set DBS = Nothing
    Set prp = Nothing
    Exit Function

AddProp_Err:
    If Err = conPropNotFoundError Then
        Set prp = DBS.CreateProperty(strName, varType, varValue)
        DBS.Properties.Append prp
        Resume
    Else
        AddAppProperty = False
        Resume AddProp_Bye
    End If
End Function