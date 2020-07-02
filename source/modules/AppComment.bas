Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        AppComment
' Level:        Framework class
' Version:      1.05
'
' Description:  Comment object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/9/2016   - 1.01 - added SaveToDb() revised to AppComment (Comment reserved word)
'               --------------- Reference Library ------------------
'               BLC - 9/19/2017  - 1.02 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/4/2017 - 1.03 - SaveToDb() code cleanup
'               BLC - 10/6/2017 - 1.04 - removed GetClass() after Factory class instatiation implemented
'               BLC - 10/17/2017 - 1.05 - code cleanup
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_CommentType As String
Private m_TypeID As Integer
Private m_Comment As String
Private m_CommentDate As Date
Private m_CommentorID As Integer    'Long??
Private m_MaxLength As Integer

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    If IsNumeric(Value) Then
        m_ID = Value
    End If
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let TypeID(Value As Integer)
    If IsNumeric(Value) Then
        m_TypeID = Value
    End If
End Property

Public Property Get TypeID() As Integer
    TypeID = m_TypeID
End Property

Public Property Let CommentType(Value As String)
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_CommentType = Value
    End If
End Property

Public Property Get CommentType() As String
    CommentType = m_CommentType
End Property

Public Property Let Comment(Value As String)
    If ValidateString(Value, "paragraph") Then
        m_Comment = Value
    End If
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let CommentorID(Value As Integer)
    If IsNumeric(Value) Then
        m_CommentorID = Value
    End If
End Property

Public Property Get CommentorID() As Integer
    ID = m_CommentorID
End Property

Public Property Let MaxLength(Value As Integer)
    If IsNumeric(Value) Then
        m_MaxLength = Value
    End If
End Property

Public Property Get MaxLength() As Integer
    MaxLength = m_MaxLength
End Property

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========
' handled by Factory class

'======== Standard Methods ===========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[AppComment class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    'Set m_ID = 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[AppComment class])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          AddComment
' Description:  Add new Comment item
' Assumptions:  -
' Parameters:   context - what the Comment is about/Comment type (string)
'               Comment
'               recordID - ID for the record the Comment references (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 19, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/19/2015 - initial version
'   BLC - 10/17/2017 - code cleanup
' ---------------------------------
Public Sub AddComment()
On Error GoTo Err_Handler

    Me.SaveToDb False

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddComment[AppComment class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  -
' Parameters:   IsUpdate - indicates if data is an update vs. an insert (boolean, optional)
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 8/9/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_comment"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "Comment"
        Params(1) = .CommentType
        Params(2) = .TypeID
        Params(3) = .Comment
        Params(4) = .CommentorID
        
        If IsUpdate Then
            Template = "u_comment"
            Params(5) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[AppComment class])"
    End Select
    Resume Exit_Handler
End Sub