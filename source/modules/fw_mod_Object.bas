Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_Ojbect
' Level:        Framework module
' Version:      1.02
' Description:  Framework-wide related object values, functions & subroutines
'
' Source/date:  Bonnie Campbell, January 22, 2018 for NCPN tools
' Revisions:    BLC, 1/22/2018 - 1.00 - initial version
'               BLC, 5/16/2019 - 1.01 - added fw_ module prefix
'               BLC, 3/9/2020  - 1.02 - 64-bit OS updates
' =================================

'-----------------------------------------------------------------------
' Declarations
'-----------------------------------------------------------------------
Private Const POINTERSIZE As Long = 4
Private Const ZEROPOINTER As Long = 0
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, _
                                                  ByRef Source As Any, _
                                                  ByVal Length As Long)

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
' Function:     GetPointer
' Description:  Retrieves the pointer (long) of an object
' Assumptions:  Transfers 32-bit object pointer value to long value
' Parameters:   obj - object whose pointer should be retrieved (object)
' Returns:      -
' Throws:       none
' References:
'   ChrisO, July 11, 2011
'   https://access-programmers.co.uk/forums/showthread.php?t=212556
' Source/date:  Bonnie Campbell, January 22, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/22/2018 - initial version
' ---------------------------------
Public Function GetPointer(ByRef obj As Object) As Long
On Error GoTo Err_Handler

    Dim ptr As Long
    
    RtlMoveMemory ptr, obj, POINTERSIZE
    
    GetPointer = ptr

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getPointer[fw_mod_Ojbect])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     GetObject
' Description:  Retrieves object based on pointer (long)
' Assumptions:  Transfers 32-bit object long value back to pointer value
' Parameters:   ptr - long pointer to object (long)
' Returns:      -
' Throws:       none
' References:
'   ChrisO, July 11, 2011
'   https://access-programmers.co.uk/forums/showthread.php?t=212556
' Source/date:  Bonnie Campbell, January 22, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/22/2018 - initial version
' ---------------------------------
Public Function GetObject(ByVal ptr As Long) As Object
On Error GoTo Err_Handler

    Dim obj As Object
    
    RtlMoveMemory obj, ptr, POINTERSIZE
    
    Set GetObject = obj

    'cleanup
    RtlMoveMemory obj, ZEROPOINTER, POINTERSIZE
    Set obj = Nothing

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetObject[fw_mod_Ojbect])"
    End Select
    Resume Exit_Handler
End Function